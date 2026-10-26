import { capture } from "./workflow.js";
import { notify } from "./outlook.js";
export async function addToDo(event, deps = {}) {
  const captureAction = deps.capture || capture;
  const notifyAction = deps.notify || notify;
  try {
    const result = await captureAction();
    const message = result.status === "created" ? "To Do task created." : result.status === "existing" ? "A linked To Do task already exists; no duplicate was created." : result.status === "reconciled" ? "To Do task creation was reconciled; no second task was created." : "Add to To Do completed.";
    await notifyAction("addToDoResult", result.joined ? `Shared in-progress result: ${message}` : message);
  } catch (error) {
    const text = error?.code === "SETUP_REQUIRED" ? error.message : error?.code === "UNCERTAIN_CREATE" ? error.message : error?.code === "BUDGET_EXCEEDED" ? `${error.message} No task was created.` : `Add to To Do failed: ${error?.message || "Unknown error"}`;
    await notifyAction("addToDoResult", text);
  } finally { event?.completed?.(); }
}
if (globalThis.Office?.actions?.associate) globalThis.Office.actions.associate("addToDo", addToDo);
