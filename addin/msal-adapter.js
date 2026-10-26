function api() {
  const value = globalThis.msal;
  if (!value) throw new Error("MSAL browser bundle did not create globalThis.msal.");
  return value;
}
export function createNestablePublicClientApplication(options) {
  const fn = api().createNestablePublicClientApplication;
  if (typeof fn !== "function") throw new Error("This MSAL bundle does not expose NAA creation.");
  return fn(options);
}
export function isInteractionRequired(error) {
  const Ctor = api().InteractionRequiredAuthError;
  return Boolean(Ctor && error instanceof Ctor) || error?.errorCode === "interaction_required";
}
