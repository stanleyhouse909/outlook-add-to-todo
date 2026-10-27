import { config } from "./config.js";
import { createNestablePublicClientApplication, isInteractionRequired } from "./msal-adapter.js";
let clientPromise;
const normalized = value => typeof value === "string" ? value.trim().toLowerCase() : "";
function setupError(message) { const error = new Error(message); error.code = "SETUP_REQUIRED"; return error; }
export function mailboxIdentity() {
  const email = normalized(globalThis.Office?.context?.mailbox?.userProfile?.emailAddress);
  if (!email || !email.includes("@")) throw setupError("Cannot verify the Outlook mailbox. Open your own work mailbox and run Setup.");
  return email;
}
export function verifyAccount(account, expectedEmail, tenantId = config.tenantId) {
  if (!account || normalized(account.username) !== normalized(expectedEmail) || normalized(account.tenantId || account.idTokenClaims?.tid) !== normalized(tenantId)) {
    throw setupError("The Microsoft sign-in does not match this Outlook mailbox and organization. Run Setup with the matching work account. Aliases and shared mailboxes are not supported by this pilot.");
  }
  return account;
}
async function getClient() {
  if (!clientPromise) clientPromise = createNestablePublicClientApplication({ auth: { clientId: config.clientId, authority: config.authority } }).catch(error => { clientPromise = undefined; throw error; });
  return clientPromise;
}
export async function interactiveSetup() {
  const email = mailboxIdentity();
  const client = await getClient();
  const result = await client.acquireTokenPopup({ scopes: [...config.scopes], loginHint: email });
  verifyAccount(result.account, email);
  if (mailboxIdentity() !== email) throw setupError("The mailbox changed during sign-in. Run Setup again.");
  return result;
}
export async function silentToken() {
  const email = mailboxIdentity();
  const client = await getClient();
  try {
    const accounts = (client.getAllAccounts?.() || []).filter(account => normalized(account.username) === email && normalized(account.tenantId || account.idTokenClaims?.tid) === normalized(config.tenantId));
    if (accounts.length > 1) throw setupError("Multiple matching sign-in sessions were found. Run Setup before capturing a task.");
    // NAA brokers may have no local account cache. Supply a mailbox hint and
    // verify the broker's returned identity before releasing a token.
    const result = await client.acquireTokenSilent({ scopes: [...config.scopes], ...(accounts.length ? { account: accounts[0] } : { loginHint: email }) });
    verifyAccount(result.account, email);
    if (mailboxIdentity() !== email) throw setupError("The mailbox changed during sign-in. Try again.");
    if (typeof result.accessToken !== "string" || !result.accessToken) throw setupError("Microsoft did not return a usable sign-in token.");
    return result.accessToken;
  } catch (error) {
    if (isInteractionRequired(error)) throw setupError("Setup is required before Add to To Do can run.");
    throw error;
  }
}
