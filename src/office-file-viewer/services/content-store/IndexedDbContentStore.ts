import { createContentAbortError, createContentStoreError } from './errors';
import type {
  OfficeContentMetaRecord,
  OfficeContentRecord,
  OfficeContentStore,
} from './types';

const DATABASE_NAME = 'office-file-viewer-content-cache';
const DATABASE_VERSION = 1;
const RECORD_STORE = 'records';
const SESSION_STORE = 'sessions';
const SESSION_TTL = 24 * 60 * 60 * 1000;

type StoredRecord<TMeta, TValue> = OfficeContentRecord<TMeta, TValue> & {
  sessionId: string;
  namespace: string;
};

type StoredSession = {
  sessionId: string;
  namespace: string;
  createdAt: number;
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
    await this.cleanupAndTouchSession(database);
    return database;
  }

  private async cleanupAndTouchSession(database: IDBDatabase) {
    const transaction = database.transaction(
      [RECORD_STORE, SESSION_STORE],
      'readwrite',
    );
    const records = transaction.objectStore(RECORD_STORE);
    const sessions = transaction.objectStore(SESSION_STORE);
    const existingSessions = await requestResult(
      sessions.getAll() as IDBRequest<StoredSession[]>,
    );
    const now = Date.now();
    existingSessions.forEach((session) => {
      if (now - session.lastAccessAt <= SESSION_TTL) return;
      records.delete(sessionRange(session.sessionId, session.namespace));
      sessions.delete([session.sessionId, session.namespace]);
    });
    sessions.put({
      sessionId: this.sessionId,
      namespace: this.namespace,
      createdAt: this.createdAt,
      lastAccessAt: now,
    } satisfies StoredSession);
    await transactionComplete(transaction);
  }

  private async touchSession(transaction: IDBTransaction) {
    transaction.objectStore(SESSION_STORE).put({
      sessionId: this.sessionId,
      namespace: this.namespace,
      createdAt: this.createdAt,
      lastAccessAt: Date.now(),
    } satisfies StoredSession);
  }

  getMeta(key: string) {
    this.ensureAvailable();
    return this.metaRecords.get(key);
  }

  async get(key: string, signal?: AbortSignal) {
    this.ensureAvailable(signal);
    const database = await this.databasePromise;
    this.ensureAvailable(signal);
    const transaction = database.transaction(
      [RECORD_STORE, SESSION_STORE],
      'readwrite',
    );
    const stored = await requestResult(
      transaction
        .objectStore(RECORD_STORE)
        .get([this.sessionId, this.namespace, key]) as IDBRequest<
        StoredRecord<TMeta, TValue> | undefined
      >,
    );
    await this.touchSession(transaction);
    await transactionComplete(transaction);
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
    if (typeof structuredClone !== 'function') {
      throw new Error('当前环境不支持 structuredClone');
    }
    const cloned = structuredClone(record);
    const database = await this.databasePromise;
    this.ensureAvailable();
    const transaction = database.transaction(
      [RECORD_STORE, SESSION_STORE],
      'readwrite',
    );
    transaction.objectStore(RECORD_STORE).put({
      ...cloned,
      sessionId: this.sessionId,
      namespace: this.namespace,
    } satisfies StoredRecord<TMeta, TValue>);
    await this.touchSession(transaction);
    await transactionComplete(transaction);
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
    const transaction = database.transaction(
      [RECORD_STORE, SESSION_STORE],
      'readwrite',
    );
    transaction
      .objectStore(RECORD_STORE)
      .delete([this.sessionId, this.namespace, key]);
    await this.touchSession(transaction);
    await transactionComplete(transaction);
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
