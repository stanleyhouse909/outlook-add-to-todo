import { config } from "./config.js";
import { selectedMessageGraphId } from "./outlook.js";
import { silentToken } from "./auth.js";
import { BudgetExceededError, requireId, validateMessage, createDeadline, getMessage, getDefaultList, createExactlyOnceBestEffort } from "./graph.js";
const inFlight = new Map();
export async function runCapture(restId, deps = {}) {
  requireId(restId, "Selected email ID");
  const deadline = createDeadline(deps.captureBudgetMs ?? config.captureBudgetMs, deps.now);
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(new BudgetExceededError("Capture budget expired before task creation.")), deps.captureBudgetMs ?? config.captureBudgetMs);
  const ctx = { deadline, signal: controller.signal, fetchImpl: deps.fetchImpl, now: deps.now, reconciliationBudgetMs: deps.reconciliationBudgetMs };
  try {
    const checkActive = () => { if (controller.signal.aborted || deadline.expired()) throw new BudgetExceededError("Capture budget expired before task creation."); };
    // Only race authentication: no Graph write can exist yet. Late auth results
    // settle this detached promise but cannot resume the workflow.
    let abortHandler;
    const cancelled = new Promise((_, reject) => { abortHandler = () => reject(controller.signal.reason); controller.signal.addEventListener("abort", abortHandler, { once: true }); });
    let token;
    try { token = await Promise.race([Promise.resolve().then(() => (deps.silentToken || silentToken)()), cancelled]); }
    finally { controller.signal.removeEventListener("abort", abortHandler); }
    checkActive();
    const message = await (deps.getMessage || getMessage)(token, restId, ctx);
    checkActive(); validateMessage(message);
    const list = await (deps.getDefaultList || getDefaultList)(token, ctx);
    checkActive(); requireId(list?.id, "Default list ID");
    return await (deps.createTask || createExactlyOnceBestEffort)(token, list, message, ctx);
  } finally { clearTimeout(timer); controller.abort(); }
}
export function capture(deps = {}) {
  const id = (deps.selectedMessageGraphId || selectedMessageGraphId)();
  const active = inFlight.get(id);
  if (active) return active.then(result => ({ ...result, joined: true }));
  const operation = runCapture(id, deps);
  inFlight.set(id, operation);
  return operation.finally(() => { if (inFlight.get(id) === operation) inFlight.delete(id); });
}
export const __testing = { inFlight };
