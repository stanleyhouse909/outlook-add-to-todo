import { PublicClientApplication } from "https://cdn.jsdelivr.net/npm/@azure/msal-browser@4.15.0/+esm";

const CLIENT_ID = "7133549f-d39f-4bf6-b48d-68d91f3c6818";
const TENANT_ID = "962ee2ef-4803-4b8d-beb0-8c8368a24a4c";
const REDIRECT_URI = "https://stanleyhouse909.github.io/outlook-add-to-todo/auth.html";
const SCOPES = ["User.Read", "Tasks.ReadWrite"];

const status = document.getElementById("status");
const testButton = document.getElementById("testButton");
const signOutButton = document.getElementById("signOutButton");

const msal = new PublicClientApplication({
  auth: {
    clientId: CLIENT_ID,
    authority: `https://login.microsoftonline.com/${TENANT_ID}`,
    redirectUri: REDIRECT_URI,
    postLogoutRedirectUri: REDIRECT_URI
  },
  cache: {
    cacheLocation: "sessionStorage",
    storeAuthStateInCookie: false
  }
});

await msal.initialize();

function show(message, detail = "") {
  status.textContent = detail ? `${message}\n\n${detail}` : message;
}

async function getToken() {
  let account = msal.getAllAccounts()[0];
  if (!account) {
    const login = await msal.loginPopup({ scopes: SCOPES, prompt: "select_account" });
    account = login.account;
  }
  msal.setActiveAccount(account);
  try {
    return await msal.acquireTokenSilent({ scopes: SCOPES, account });
  } catch {
    return await msal.acquireTokenPopup({ scopes: SCOPES, account });
  }
}

async function runTest() {
  testButton.disabled = true;
  show("Signing in and requesting delegated To Do access...");
  try {
    const token = await getToken();
    show("Signed in. Reading the names and identifiers of your To Do lists...");
    const response = await fetch("https://graph.microsoft.com/v1.0/me/todo/lists?$select=id,displayName,wellknownListName,isOwner,isShared", {
      headers: { Authorization: `Bearer ${token.accessToken}` }
    });
    const body = await response.json();
    if (!response.ok) throw new Error(`${response.status} ${response.statusText}\n${JSON.stringify(body, null, 2)}`);
    const defaultList = body.value.find(list => list.wellknownListName === "defaultList");
    if (!defaultList) throw new Error(`Graph access succeeded, but no list marked defaultList was returned.\n${JSON.stringify(body, null, 2)}`);
    const summary = body.value.map(list => `• ${list.displayName} (${list.wellknownListName || "none"})`).join("\n");
    show(
      "SUCCESS: Microsoft Graph access works.\nThe built-in main To Do list was found.",
      `Default list: ${defaultList.displayName}\nDefault list ID: ${defaultList.id}\n\nLists returned:\n${summary}`
    );
    signOutButton.hidden = false;
  } catch (error) {
    const message = error?.errorMessage || error?.message || String(error);
    show("TEST DID NOT COMPLETE. Copy or screenshot the details below.", message);
  } finally {
    testButton.disabled = false;
  }
}

testButton.addEventListener("click", runTest);
signOutButton.addEventListener("click", async () => {
  sessionStorage.clear();
  show("Local session cleared. Reload the page to test again.");
  signOutButton.hidden = true;
});
