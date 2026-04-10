// chat-adapter-outlook-email
// Microsoft Outlook / M365 email adapter for Vercel Chat SDK

export { OutlookEmailAdapter, createOutlookEmailAdapter } from "./adapter.js";
export { OutlookEmailFormatConverter } from "./format-converter.js";
export { GraphClient } from "./graph-client.js";
export { SubscriptionManager } from "./subscription-manager.js";
export type { SubscriptionInfo } from "./subscription-manager.js";
export {
	encodeThreadId,
	decodeThreadId,
	channelIdFromThreadId,
} from "./thread-id.js";
export {
	handleGraphWebhook,
	extractMessageId,
	extractMailbox,
} from "./webhook-handler.js";
export type { WebhookResult } from "./webhook-handler.js";
export {
	parseGraphMessage,
	isSelfEcho,
	isAutoReply,
	mapAttachments,
} from "./message-parser.js";
export type { MessageParserOptions } from "./message-parser.js";
export { renderMessage } from "./message-renderer.js";
export { normalizeEmailBody, stripHtml } from "./body-normalizer.js";
export { handleGraphError, NotImplementedError } from "./errors.js";
export type {
	OutlookEmailAdapterConfig,
	OutlookThreadId,
	GraphMessage,
	GraphAttachment,
	GraphNotification,
	GraphNotificationPayload,
	GraphSubscription,
} from "./types.js";
