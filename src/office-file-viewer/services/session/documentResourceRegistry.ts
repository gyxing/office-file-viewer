import type { OfficeDocumentSession } from './types';

const ownerSessions = new WeakMap<object, OfficeDocumentSession>();
const ownerDisposals = new WeakMap<object, Promise<void>>();

/** 把文档会话附加到不允许增加内部字段的公开模型。 */
export function attachDocumentSession(
  owner: object,
  session: OfficeDocumentSession,
) {
  const previousSession = ownerSessions.get(owner);
  if (previousSession === session) return;
  if (previousSession) void previousSession.dispose();

  ownerDisposals.delete(owner);
  ownerSessions.set(owner, session);
}

/** 幂等释放公开模型附加的文档会话。 */
export function disposeDocumentSession(owner: object | undefined) {
  if (!owner) return Promise.resolve();

  const existingDisposal = ownerDisposals.get(owner);
  if (existingDisposal) return existingDisposal;

  const session = ownerSessions.get(owner);
  const disposal = session?.dispose() ?? Promise.resolve();
  ownerSessions.delete(owner);
  ownerDisposals.set(owner, disposal);
  return disposal;
}
