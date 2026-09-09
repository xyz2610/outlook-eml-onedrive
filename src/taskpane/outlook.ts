/* global Office */

export type CurrentMessage = {
  graphId: string;
  subject: string;
  sender: string;
  received: Date;
};

export function getGraphMessageId(): string {
  const item = Office.context.mailbox.item as Office.MessageRead;
  const itemId = item.itemId;
  if (!itemId) throw new Error("Für die geöffnete Nachricht ist keine Item-ID verfügbar.");

  const host = Office.context.mailbox.diagnostics.hostName;
  if (host === "OutlookIOS" || host === "OutlookAndroid") {
    return itemId;
  }

  return Office.context.mailbox.convertToRestId(itemId, Office.MailboxEnums.RestVersion.v2_0);
}

export function getCurrentMessage(): CurrentMessage {
  const item = Office.context.mailbox.item as Office.MessageRead;
  const from = item.from;
  const sender = from?.displayName || from?.emailAddress || "Unbekannt";

  return {
    graphId: getGraphMessageId(),
    subject: item.subject || "Ohne Betreff",
    sender,
    received: item.dateTimeCreated ? new Date(item.dateTimeCreated) : new Date(),
  };
}

export function makeDefaultFilename(message: CurrentMessage): string {
  const d = message.received;
  const two = (n: number) => String(n).padStart(2, "0");
  const stamp = `${d.getFullYear()}-${two(d.getMonth() + 1)}-${two(d.getDate())}_${two(d.getHours())}${two(d.getMinutes())}`;
  const raw = `${stamp}_${message.sender}_${message.subject}.eml`;
  return sanitizeFilename(raw);
}

export function sanitizeFilename(value: string): string {
  const cleaned = value
    .replace(/[\\/:*?"<>|#%]/g, "-")
    .replace(/[\u0000-\u001F]/g, "")
    .replace(/\s+/g, " ")
    .replace(/\.+$/g, "")
    .trim();

  const withExtension = cleaned.toLowerCase().endsWith(".eml") ? cleaned : `${cleaned}.eml`;
  return `${withExtension.slice(0, -4).slice(0, 176) || "Nachricht"}.eml`;
}
