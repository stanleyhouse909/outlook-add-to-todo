export function selectedMessageGraphId(OfficeApi = globalThis.Office) {
  const mailbox = OfficeApi?.context?.mailbox;
  const itemId = mailbox?.item?.itemId;
  if (!itemId) throw new Error("Open or select a saved email before using Add to To Do.");
  if (typeof mailbox.convertToRestId !== "function") throw new Error("Outlook ID conversion is unavailable.");
  return mailbox.convertToRestId(itemId, OfficeApi.MailboxEnums.RestVersion.v2_0);
}
export function notify(key, message, OfficeApi = globalThis.Office, timeoutMs = 1500) {
  return new Promise(resolve => {
    let timer;
    const finish = () => { clearTimeout(timer); resolve(); };
    const notifications = OfficeApi?.context?.mailbox?.item?.notificationMessages;
    if (!notifications?.replaceAsync) { finish(); return; }
    timer = setTimeout(finish, timeoutMs);
    try { notifications.replaceAsync(key, { type: OfficeApi.MailboxEnums.ItemNotificationMessageType.InformationalMessage, message: String(message).slice(0, 250), icon: "Icon.16", persistent: false }, finish); }
    catch { finish(); }
  });
}
