import { config } from "./config.js";
import { createNestablePublicClientApplication, isInteractionRequired } from "./msal-adapter.js";
let clientPromise;
async function getClient() {
  if (!clientPromise) clientPromise = createNestablePublicClientApplication({ auth: { clientId: config.clientId, authority: config.authority } });
  return clientPromise;
}
export async function interactiveSetup() {
  const client = await getClient();
  return client.acquireTokenPopup({ scopes: [...config.scopes] });
}
export async function silentToken() {
  const client = await getClient();
  try {
    const accounts = client.getAllAccounts?.() || [];
    const result = await client.acquireTokenSilent({ scopes: [...config.scopes], account: accounts[0] });
    return result.accessToken;
  } catch (error) {
    if (isInteractionRequired(error)) {
      const wrapped = new Error("Setup is required before Add to To Do can run.");
      wrapped.code = "SETUP_REQUIRED";
      throw wrapped;
    }
    throw error;
  }
}
