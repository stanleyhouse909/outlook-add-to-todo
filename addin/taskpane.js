import { interactiveSetup } from "./auth.js";
async function setup() { const status = document.getElementById("status"); try { await interactiveSetup(); status.textContent = "Setup completed. You can now close this pane."; } catch (e) { status.textContent = `Setup did not complete: ${e.message}`; } }
Office.onReady(() => document.getElementById("setup").addEventListener("click", setup));
