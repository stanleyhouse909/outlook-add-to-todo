import { config } from "./config.js";
export class BudgetExceededError extends Error { constructor(message) { super(message); this.name = "BudgetExceededError"; this.code = "BUDGET_EXCEEDED"; } }
export class UncertainCreateError extends Error { constructor(message, cause) { super(message, { cause }); this.name = "UncertainCreateError"; this.code = "UNCERTAIN_CREATE"; } }
export function requireId(value, label) { if (typeof value !== "string" || !value.trim()) throw new Error(`${label} is missing.`); return value; }
export function validateMessage(message) {
  requireId(message?.id, "Email ID");
  let url; try { url = new URL(message?.webLink); } catch { throw new Error("Email backlink is missing or invalid."); }
  if (url.protocol !== "https:" || !["outlook.office.com", "outlook.office365.com", "outlook.live.com"].includes(url.hostname.toLowerCase()) || url.port || url.username || url.password) throw new Error("Email backlink must use a supported Microsoft Outlook HTTPS host.");
  return message;
}
export function createDeadline(totalMs, now = () => Date.now()) { const end = now() + totalMs; return { remaining: () => Math.max(0, end - now()), expired: () => now() >= end }; }
function combineSignals(parent, timeoutMs) {
  const controller = new AbortController();
  const abort = () => controller.abort(parent?.reason || new Error("Cancelled"));
  if (parent) { if (parent.aborted) abort(); else parent.addEventListener("abort", abort, { once: true }); }
  const timer = setTimeout(() => controller.abort(new Error("Request timeout")), timeoutMs);
  return { signal: controller.signal, cleanup: () => { clearTimeout(timer); parent?.removeEventListener?.("abort", abort); } };
}
export async function graphRequest(token, path, options = {}, ctx = {}) {
  const method = options.method || "GET";
  if (typeof path !== "string" || !path.startsWith("/me/") || path.includes("#") || path.includes("\\")) throw new Error("Unsupported Graph request path.");
  if (ctx.signal?.aborted) throw ctx.signal.reason || new Error("Operation cancelled.");
  const deadline = ctx.deadline;
  const remaining = deadline ? deadline.remaining() : Infinity;
  if (remaining <= 0) throw new BudgetExceededError("Operation time budget expired before the request began.");
  const timeout = Math.max(1, Math.min(ctx.requestTimeoutMs ?? config.requestTimeoutMs, remaining));
  const linked = combineSignals(ctx.signal, timeout);
  try {
    const response = await (ctx.fetchImpl || fetch)(`${config.graphBaseUrl}${path}`, { ...options, redirect: "error", signal: linked.signal, headers: { Authorization: `Bearer ${token}`, Accept: "application/json", ...(options.body ? { "Content-Type": "application/json" } : {}), ...options.headers } });
    const text = await response.text();
    const body = text ? JSON.parse(text) : undefined;
    if (!response.ok) { const error = new Error(`Graph ${method} ${response.status}: ${body?.error?.message || response.statusText}`); error.status = response.status; throw error; }
    return body;
  } catch (error) {
    if (deadline?.expired() && method !== "POST") throw new BudgetExceededError("Operation time budget expired.");
    throw error;
  } finally { linked.cleanup(); }
}
export async function getMessage(token, id, ctx) {
  const select = "id,internetMessageId,subject,from,receivedDateTime,webLink";
  return graphRequest(token, `/me/messages/${encodeURIComponent(id)}?$select=${select}`, { headers: { Prefer: 'IdType="ImmutableId"' } }, ctx);
}
export function nextPagePath(next, collectionPath) {
  if (!next) return null;
  const url = new URL(next);
  const expected = new URL(config.graphBaseUrl + collectionPath.split("?")[0]);
  if (url.origin !== expected.origin || url.pathname !== expected.pathname || url.username || url.password || url.hash) throw new Error("Untrusted Graph pagination link.");
  return url.pathname.slice(new URL(config.graphBaseUrl).pathname.length) + url.search;
}
export async function* pages(token, path, ctx) {
  const seen = new Set(); let next = path;
  while (next) {
    if (seen.has(next)) throw new Error("Graph pagination loop; capture stopped.");
    seen.add(next);
    const data = await graphRequest(token, next, {}, ctx);
    if (!data || !Array.isArray(data.value)) throw new Error("Graph collection response is invalid.");
    yield data.value;
    next = nextPagePath(data["@odata.nextLink"], path);
  }
}
export async function getDefaultList(token, ctx) {
  for await (const lists of pages(token, "/me/todo/lists", ctx)) {
    const value = lists.find(x => x.wellknownListName === "defaultList");
    if (value) { requireId(value.id, "Default list ID"); return value; }
  }
  throw new Error("The default Microsoft To Do list is unavailable.");
}
export async function findExistingTask(token, listId, externalId, ctx) {
  for await (const tasks of pages(token, `/me/todo/lists/${encodeURIComponent(listId)}/tasks?$top=${config.duplicateTaskLimit}&$orderby=createdDateTime desc`, ctx)) {
  let next = 0, found = null, failure;
  const checkActive = () => {
    if (ctx.signal?.aborted) throw ctx.signal.reason || new Error("Duplicate checking cancelled.");
    if (ctx.deadline?.expired()) throw new BudgetExceededError("Duplicate checking did not finish within the capture budget.");
  };
  async function worker() {
    while (!failure && !found && next < tasks.length) {
      try {
        checkActive();
        const task = tasks[next++];
        requireId(task.id, "Task ID");
        for await (const resources of pages(token, `/me/todo/lists/${encodeURIComponent(listId)}/tasks/${encodeURIComponent(task.id)}/linkedResources`, ctx)) {
          checkActive();
          if (resources.some(x => x.externalId === externalId)) found = task;
          if (found) break;
        }
      } catch (error) { failure ||= error; }
    }
  }
  // Join all started reads before returning, including failure paths.
  await Promise.all(Array.from({ length: Math.min(5, tasks.length) }, () => worker()));
  if (failure) throw failure;
  checkActive();
  if (found) return found;
  }
  return null;
}

export function taskPayload(message) {
  validateMessage(message);
  const sender = message.from?.emailAddress?.name || message.from?.emailAddress?.address || "Unknown";
  const details = [`From: ${sender}`, `Received: ${message.receivedDateTime || "Unknown"}`, `Immutable Graph message ID: ${message.id}`, `Internet Message ID: ${message.internetMessageId || "Unavailable"}`, `Open original email: ${message.webLink}`].join("\n");
  return { title: message.subject || "Email follow-up", body: { contentType: "text", content: details }, linkedResources: [{ webUrl: message.webLink, applicationName: "Microsoft Outlook", displayName: message.subject || "Original email", externalId: `outlook-immutable:${message.id}` }] };
}
export async function createExactlyOnceBestEffort(token, list, message, ctx) {
  validateMessage(message); requireId(list?.id, "Default list ID");
  const externalId = `outlook-immutable:${message.id}`;
  const existing = await findExistingTask(token, list.id, externalId, ctx);
  if (existing) return { status: "existing", task: existing };
  if (ctx.deadline?.expired()) throw new BudgetExceededError("Duplicate checking was incomplete; no task was created.");
  let created;
  ctx.beforePost?.();
  try {
    created = await graphRequest(token, `/me/todo/lists/${encodeURIComponent(list.id)}/tasks`, { method: "POST", body: JSON.stringify(taskPayload(message)) }, ctx);
    requireId(created?.id, "Created task confirmation ID");
    return { status: "created", task: created };
  } catch (error) {
    if ([400, 401, 403, 404, 405, 413, 415, 422].includes(error?.status)) { ctx.onDefinitiveFailure?.(); throw error; }
    const reconciliationMs = Math.min(ctx.reconciliationBudgetMs ?? config.reconciliationBudgetMs, ctx.deadline?.remaining() ?? Infinity);
    const reconciliation = createDeadline(reconciliationMs, ctx.now);
    const reconciliationSignals = combineSignals(ctx.signal, Math.max(1, reconciliationMs));
    const reconcileCtx = { ...ctx, signal: reconciliationSignals.signal, deadline: reconciliation, requestTimeoutMs: config.reconciliationTimeoutMs };
    try {
      const found = await findExistingTask(token, list.id, externalId, reconcileCtx);
      if (found) return { status: "reconciled", task: found };
    } catch { /* preserve uncertainty */ }
    finally { reconciliationSignals.cleanup(); }
    throw new UncertainCreateError("Task creation may have reached Microsoft Graph. No second creation request was sent.", error);
  }
}
