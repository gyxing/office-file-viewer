import { createContentAbortError, createContentStoreError } from './errors';
import type {
  OfficeContentMetaRecord,
  OfficeContentRecord,
  OfficeContentStore,
} from './types';

/** 浏览器内容缓存使用的 IndexedDB 数据库名称。 */
const DATABASE_NAME = 'office-file-viewer-content-cache';
/** 浏览器内容缓存数据库的结构版本。 */
const DATABASE_VERSION = 1;
/** IndexedDB 中保存内容记录的对象仓库名称。 */
const RECORD_STORE = 'records';
/** IndexedDB 中保存解析会话的对象仓库名称。 */
const SESSION_STORE = 'sessions';
/** 解析会话缓存的有效期，单位为毫秒。 */
const SESSION_TTL = 24 * 60 * 60 * 1000;
/** 活跃会话写回访问时间的最小间隔，避免连续读取放大 IndexedDB 写入。 */
const SESSION_TOUCH_INTERVAL = 60 * 1000;
/** 过期会话全表扫描的最小间隔；会话有效期较长，无需为每个 Store 重复清理。 */
const SESSION_CLEANUP_INTERVAL = 5 * 60 * 1000;

/** 同一页面生命周期内最近一次成功完成过期清理的时间。 */
let lastSessionCleanupAt = 0;
/** 复用并发 Store 初始化触发的过期会话清理任务。 */
let sessionCleanupPromise: Promise<void> | undefined;

/** IndexedDB 中持久化的内容记录。 */
type StoredRecord<TMeta, TValue> = OfficeContentRecord<TMeta, TValue> & {
  /** 当前解析或预览会话的标识。 */
  sessionId: string;
  /** 当前记录所属的内容存储命名空间。 */
  namespace: string;
};

/** IndexedDB 中持久化的解析会话。 */
type StoredSession = {
  /** 当前解析或预览会话的标识。 */
  sessionId: string;
  /** 当前会话所属的内容存储命名空间。 */
  namespace: string;
  /** 记录创建时的时间戳。 */
  createdAt: number;
  /** 会话最近一次访问的时间戳。 */
  lastAccessAt: number;
};

function requestResult<T>(request: IDBRequest<T>) {
  return new Promise<T>((resolve, reject) => {
    request.addEventListener('success', () => resolve(request.result), {
      once: true,
    });
    request.addEventListener('error', () => reject(request.error), {
      once: true,
    });
  });
}

function transactionComplete(transaction: IDBTransaction) {
  return new Promise<void>((resolve, reject) => {
    transaction.addEventListener('complete', () => resolve(), { once: true });
    transaction.addEventListener('abort', () => reject(transaction.error), {
      once: true,
    });
    transaction.addEventListener('error', () => reject(transaction.error), {
      once: true,
    });
  });
}

function sessionRange(sessionId: string, namespace: string) {
  return IDBKeyRange.bound(
    [sessionId, namespace, ''],
    [sessionId, namespace, '\uffff'],
  );
}

/** 按固定间隔清理所有命名空间中的过期会话，避免多格式 Store 重复全表扫描。 */
function cleanupExpiredSessions(database: IDBDatabase) {
  const now = Date.now();
  if (now - lastSessionCleanupAt < SESSION_CLEANUP_INTERVAL) {
    return Promise.resolve();
  }
  if (sessionCleanupPromise) return sessionCleanupPromise;

  const cleanup = (async () => {
    const transaction = database.transaction(
      [RECORD_STORE, SESSION_STORE],
      'readwrite',
    );
    const records = transaction.objectStore(RECORD_STORE);
    const sessions = transaction.objectStore(SESSION_STORE);
    const existingSessions = await requestResult(
      sessions.getAll() as IDBRequest<StoredSession[]>,
    );
    existingSessions.forEach((session) => {
      if (now - session.lastAccessAt <= SESSION_TTL) return;
      records.delete(sessionRange(session.sessionId, session.namespace));
      sessions.delete([session.sessionId, session.namespace]);
    });
    await transactionComplete(transaction);
    lastSessionCleanupAt = now;
  })();
  sessionCleanupPromise = cleanup.finally(() => {
    sessionCleanupPromise = undefined;
  });
  return sessionCleanupPromise;
}

/** 一次内容事务可选择携带的会话访问时间写回凭证。 */
type SessionTouchReservation = {
  /** 区分并发事务的递增标识。 */
  sequence: number;
  /** 本次访问发生的时间。 */
  touchedAt: number;
};

/** 使用会话隔离的 IndexedDB 保存可 structured-clone 的冷内容。 */
export class IndexedDbContentStore<TMeta, TValue>
  implements OfficeContentStore<TMeta, TValue>
{
  private readonly metaRecords = new Map<
    string,
    OfficeContentMetaRecord<TMeta>
  >();
  private readonly createdAt = Date.now();
  private readonly databasePromise: Promise<IDBDatabase>;
  /** 最近一次成功写回当前会话访问时间的时间戳。 */
  private lastTouchedAt = 0;
  /** 区分并发会话写回事务的递增序号。 */
  private touchSequence = 0;
  /** 尚未完成的会话写回事务序号。 */
  private pendingTouchSequence: number | undefined;
  private disposed = false;
  private disposePromise: Promise<void> | undefined;

  constructor(
    private readonly sessionId: string,
    private readonly namespace: string,
  ) {
    if (typeof indexedDB === 'undefined') {
      this.databasePromise = Promise.reject(
        new Error('当前环境不支持 IndexedDB'),
      );
      return;
    }
    this.databasePromise = this.openDatabase();
  }

  private ensureAvailable(signal?: AbortSignal) {
    if (this.disposed) {
      throw createContentStoreError(
        'CONTENT_STORE_DISPOSED',
        'ContentStore 已释放',
      );
    }
    if (signal?.aborted) throw createContentAbortError();
  }

  private async openDatabase() {
    const request = indexedDB.open(DATABASE_NAME, DATABASE_VERSION);
    request.addEventListener('upgradeneeded', () => {
      const database = request.result;
      if (!database.objectStoreNames.contains(RECORD_STORE)) {
        database.createObjectStore(RECORD_STORE, {
          keyPath: ['sessionId', 'namespace', 'key'],
        });
      }
      if (!database.objectStoreNames.contains(SESSION_STORE)) {
        database.createObjectStore(SESSION_STORE, {
          keyPath: ['sessionId', 'namespace'],
        });
      }
    });
    const database = await requestResult(request);
    await cleanupExpiredSessions(database);
    await this.initializeSession(database);
    return database;
  }

  /** 初始化当前命名空间的会话记录，并建立后续访问时间节流基线。 */
  private async initializeSession(database: IDBDatabase) {
    const transaction = database.transaction(SESSION_STORE, 'readwrite');
    const now = Date.now();
    transaction.objectStore(SESSION_STORE).put({
      sessionId: this.sessionId,
      namespace: this.namespace,
      createdAt: this.createdAt,
      lastAccessAt: now,
    } satisfies StoredSession);
    await transactionComplete(transaction);
    this.lastTouchedAt = now;
  }

  /** 只允许一个并发内容事务携带到期的会话访问时间写回。 */
  private reserveSessionTouch(): SessionTouchReservation | undefined {
    const touchedAt = Date.now();
    if (
      this.pendingTouchSequence !== undefined ||
      touchedAt - this.lastTouchedAt < SESSION_TOUCH_INTERVAL
    ) {
      return undefined;
    }
    const sequence = ++this.touchSequence;
    this.pendingTouchSequence = sequence;
    return { sequence, touchedAt };
  }

  /** 把访问时间写入已经包含会话仓库的内容事务。 */
  private writeSessionTouch(
    transaction: IDBTransaction,
    reservation: SessionTouchReservation,
  ) {
    transaction.objectStore(SESSION_STORE).put({
      sessionId: this.sessionId,
      namespace: this.namespace,
      createdAt: this.createdAt,
      lastAccessAt: reservation.touchedAt,
    } satisfies StoredSession);
  }

  /** 根据事务结果提交或释放访问时间写回凭证。 */
  private settleSessionTouch(
    reservation: SessionTouchReservation | undefined,
    completed: boolean,
  ) {
    if (!reservation || this.pendingTouchSequence !== reservation.sequence) {
      return;
    }
    if (completed) this.lastTouchedAt = reservation.touchedAt;
    this.pendingTouchSequence = undefined;
  }

  getMeta(key: string) {
    this.ensureAvailable();
    return this.metaRecords.get(key);
  }

  async get(key: string, signal?: AbortSignal) {
    this.ensureAvailable(signal);
    const database = await this.databasePromise;
    this.ensureAvailable(signal);
    const sessionTouch = this.reserveSessionTouch();
    let stored: StoredRecord<TMeta, TValue> | undefined;
    try {
      const transaction = database.transaction(
        sessionTouch ? [RECORD_STORE, SESSION_STORE] : RECORD_STORE,
        sessionTouch ? 'readwrite' : 'readonly',
      );
      stored = await requestResult(
        transaction
          .objectStore(RECORD_STORE)
          .get([this.sessionId, this.namespace, key]) as IDBRequest<
          StoredRecord<TMeta, TValue> | undefined
        >,
      );
      if (sessionTouch) this.writeSessionTouch(transaction, sessionTouch);
      await transactionComplete(transaction);
      this.settleSessionTouch(sessionTouch, true);
    } catch (error) {
      this.settleSessionTouch(sessionTouch, false);
      throw error;
    }
    this.ensureAvailable(signal);
    if (!stored) return undefined;
    const record: OfficeContentRecord<TMeta, TValue> = {
      key: stored.key,
      revision: stored.revision,
      meta: stored.meta,
      value: stored.value,
      updatedAt: stored.updatedAt,
    };
    this.metaRecords.set(key, {
      key,
      revision: record.revision,
      meta: record.meta,
      updatedAt: record.updatedAt,
    });
    return record;
  }

  async put(record: OfficeContentRecord<TMeta, TValue>) {
    this.ensureAvailable();
    const current = this.metaRecords.get(record.key);
    if (current && record.revision <= current.revision) {
      throw createContentStoreError(
        'STALE_CONTENT_REVISION',
        `内容 ${record.key} 的 revision 必须递增`,
      );
    }
    const database = await this.databasePromise;
    this.ensureAvailable();
    const sessionTouch = this.reserveSessionTouch();
    try {
      const transaction = database.transaction(
        sessionTouch ? [RECORD_STORE, SESSION_STORE] : RECORD_STORE,
        'readwrite',
      );
      transaction.objectStore(RECORD_STORE).put({
        ...record,
        sessionId: this.sessionId,
        namespace: this.namespace,
      } satisfies StoredRecord<TMeta, TValue>);
      if (sessionTouch) this.writeSessionTouch(transaction, sessionTouch);
      await transactionComplete(transaction);
      this.settleSessionTouch(sessionTouch, true);
    } catch (error) {
      this.settleSessionTouch(sessionTouch, false);
      throw error;
    }
    this.metaRecords.set(record.key, {
      key: record.key,
      revision: record.revision,
      meta: record.meta,
      updatedAt: record.updatedAt,
    });
  }

  pin() {
    this.ensureAvailable();
    return () => undefined;
  }

  async delete(key: string) {
    this.ensureAvailable();
    const database = await this.databasePromise;
    this.ensureAvailable();
    const sessionTouch = this.reserveSessionTouch();
    try {
      const transaction = database.transaction(
        sessionTouch ? [RECORD_STORE, SESSION_STORE] : RECORD_STORE,
        'readwrite',
      );
      transaction
        .objectStore(RECORD_STORE)
        .delete([this.sessionId, this.namespace, key]);
      if (sessionTouch) this.writeSessionTouch(transaction, sessionTouch);
      await transactionComplete(transaction);
      this.settleSessionTouch(sessionTouch, true);
    } catch (error) {
      this.settleSessionTouch(sessionTouch, false);
      throw error;
    }
    this.metaRecords.delete(key);
  }

  dispose() {
    if (this.disposePromise) return this.disposePromise;
    this.disposed = true;
    this.metaRecords.clear();
    this.disposePromise = (async () => {
      let database: IDBDatabase;
      try {
        database = await this.databasePromise;
      } catch {
        return;
      }
      const transaction = database.transaction(
        [RECORD_STORE, SESSION_STORE],
        'readwrite',
      );
      transaction
        .objectStore(RECORD_STORE)
        .delete(sessionRange(this.sessionId, this.namespace));
      transaction
        .objectStore(SESSION_STORE)
        .delete([this.sessionId, this.namespace]);
      await transactionComplete(transaction);
      database.close();
    })();
    return this.disposePromise;
  }
}
