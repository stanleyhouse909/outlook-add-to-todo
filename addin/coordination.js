import { mailboxIdentity } from './auth.js';
import { config } from './config.js';
import { createExactlyOnceBestEffort, findExistingTask, requireId, validateMessage, UncertainCreateError } from './graph.js';
export async function createCoordinatedTask(token, list, message, ctx, environment = {}) {
  const email = (environment.mailboxIdentity || mailboxIdentity)();
  requireId(list?.id, 'Default list ID'); validateMessage(message);
  const locks = environment.locks || globalThis.navigator?.locks;
  const cryptoApi = environment.crypto || globalThis.crypto;
  let storage;
  try { storage = environment.storage || globalThis.localStorage; } catch { /* fail closed below */ }
  if (!locks?.request || !cryptoApi?.subtle || !storage) throw new Error('This Outlook environment cannot provide duplicate safeguards. No task was created. Use a supported environment after review.');
  const raw = JSON.stringify([config.tenantId, email, list.id, message.id]);
  const digest = await cryptoApi.subtle.digest('SHA-256', new TextEncoder().encode(raw));
  const key = 'addtodo:v1:' + Array.from(new Uint8Array(digest), x => x.toString(16).padStart(2, '0')).join('');
  return locks.request(key, { mode: 'exclusive', signal: ctx.signal }, async () => {
    if (ctx.signal?.aborted || ctx.deadline?.expired()) throw new Error('Capture expired while waiting for duplicate safeguards. No task was created.');
    let previous;
    try { previous = storage.getItem(key); } catch { throw new Error('Duplicate history cannot be read. No task was created.'); }
    if (previous !== null) {
      const existing = await findExistingTask(token, list.id, `outlook-immutable:${message.id}`, ctx);
      if (existing) return { status: 'existing', task: existing };
      throw new UncertainCreateError('A previous capture attempt is recorded, but its task could not be confirmed. Check To Do and request a reviewed recovery; automatic retry is blocked and no new task was created.');
    }
    return createExactlyOnceBestEffort(token, list, message, { ...ctx, onDefinitiveFailure() {
      try { storage.removeItem(key); } catch { /* Retain conservative retry block if removal fails. */ }
    }, beforePost() {
      if (ctx.signal?.aborted || ctx.deadline?.expired()) throw new Error('Capture expired before task creation.');
      try {
        storage.setItem(key, 'attempted');
        if (storage.getItem(key) !== 'attempted') throw new Error('readback failed');
      } catch { throw new Error('Duplicate history cannot be saved. No task was created.'); }
    } });
  });
}
