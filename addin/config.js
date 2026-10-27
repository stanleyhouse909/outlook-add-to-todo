export const config = Object.freeze({
  clientId: "7133549f-d39f-4bf6-b48d-68d91f3c6818",
  tenantId: "962ee2ef-4803-4b8d-beb0-8c8368a24a4c",
  authority: "https://login.microsoftonline.com/962ee2ef-4803-4b8d-beb0-8c8368a24a4c",
  scopes: Object.freeze(["Mail.ReadBasic", "Tasks.ReadWrite"]),
  graphBaseUrl: "https://graph.microsoft.com/v1.0",
  requestTimeoutMs: 20000,
  reconciliationTimeoutMs: 30000,
  captureBudgetMs: 45000,
  reconciliationBudgetMs: 30000,
  duplicateTaskLimit: 100
});
