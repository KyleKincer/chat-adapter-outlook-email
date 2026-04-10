import type {
	Adapter,
	AdapterPostableMessage,
	EmojiValue,
	FetchOptions,
	FetchResult,
	FormattedContent,
	RawMessage,
	ThreadInfo,
	WebhookOptions,
} from "chat";
import { Message } from "chat";
import { NotImplementedError } from "./errors.js";
import { OutlookEmailFormatConverter } from "./format-converter.js";
import { GraphClient } from "./graph-client.js";
import {
	isAutoReply,
	isSelfEcho,
	mapAttachments,
	parseGraphMessage,
} from "./message-parser.js";
import { renderMessage } from "./message-renderer.js";
import { SubscriptionManager } from "./subscription-manager.js";
import {
	channelIdFromThreadId,
	decodeThreadId,
	encodeThreadId,
} from "./thread-id.js";
import type {
	GraphMessage,
	GraphNotification,
	OutlookEmailAdapterConfig,
	OutlookThreadId,
} from "./types.js";
import {
	extractMailbox,
	extractMessageId,
	handleGraphWebhook,
} from "./webhook-handler.js";

interface ChatInstance {
	processMessage(
		adapter: Adapter,
		threadId: string,
		message: Message,
		options?: WebhookOptions,
	): void;
}

const NEW_THREAD_PREFIX = "new-";

export class OutlookEmailAdapter
	implements Adapter<OutlookThreadId, GraphMessage>
{
	readonly name = "outlook-email";
	readonly userName: string;
	readonly botUserId: string;
	readonly persistMessageHistory = false;

	private chat: ChatInstance | null = null;
	readonly config: OutlookEmailAdapterConfig;
	private readonly formatConverter = new OutlookEmailFormatConverter();
	private readonly mailboxes: string[];
	private graphClient: GraphClient | null = null;
	private subscriptionManager: SubscriptionManager | null = null;
	private readonly processedMessageIds = new Set<string>();

	constructor(config: OutlookEmailAdapterConfig) {
		this.config = config;
		this.userName =
			config.botName ||
			process.env.OUTLOOK_BOT_NAME ||
			process.env.BOT_USERNAME ||
			"Bot";

		const mailboxEnv = process.env.OUTLOOK_MAILBOX || "";
		const mailboxConfig = config.mailbox;
		if (Array.isArray(mailboxConfig)) {
			this.mailboxes = mailboxConfig;
		} else if (mailboxConfig) {
			this.mailboxes = [mailboxConfig];
		} else if (mailboxEnv) {
			this.mailboxes = mailboxEnv.split(",").map((m) => m.trim());
		} else {
			this.mailboxes = [];
		}

		this.botUserId = this.mailboxes[0] || "";
	}

	// --- Lifecycle ---

	async initialize(chat: ChatInstance): Promise<void> {
		this.chat = chat;

		const tenantId =
			this.config.tenantId || process.env.OUTLOOK_TENANT_ID;
		const clientId =
			this.config.clientId || process.env.OUTLOOK_CLIENT_ID;
		const clientSecret =
			this.config.clientSecret || process.env.OUTLOOK_CLIENT_SECRET;

		if (!tenantId || !clientId || !clientSecret) {
			throw new Error(
				"Missing required credentials: tenantId, clientId, and clientSecret must be provided via config or environment variables",
			);
		}

		this.graphClient = new GraphClient(tenantId, clientId, clientSecret);

		// Set up subscriptions if webhook configuration is present
		const notificationUrl =
			this.config.notificationUrl || process.env.OUTLOOK_NOTIFICATION_URL;
		const clientState =
			this.config.clientState || process.env.OUTLOOK_CLIENT_STATE;

		if (notificationUrl && clientState && this.mailboxes.length > 0) {
			const renewalInterval =
				this.config.renewalIntervalMinutes ?? 4000;
			this.subscriptionManager = new SubscriptionManager(
				this.graphClient,
				notificationUrl,
				clientState,
				renewalInterval,
				this.config.logger,
			);

			await this.subscriptionManager.initialize(this.mailboxes);
		}
	}

	async disconnect(): Promise<void> {
		if (this.subscriptionManager) {
			await this.subscriptionManager.cleanup();
			this.subscriptionManager = null;
		}
	}

	// --- Thread ID ---

	encodeThreadId(platformData: OutlookThreadId): string {
		return encodeThreadId(platformData);
	}

	decodeThreadId(threadId: string): OutlookThreadId {
		return decodeThreadId(threadId);
	}

	channelIdFromThreadId(threadId: string): string {
		return channelIdFromThreadId(threadId);
	}

	// --- Webhook ---

	async handleWebhook(
		request: Request,
		options?: WebhookOptions,
	): Promise<Response> {
		const clientState =
			this.config.clientState ||
			process.env.OUTLOOK_CLIENT_STATE ||
			"";

		const result = await handleGraphWebhook(request, clientState);

		// Validation challenge — return immediately
		if (result.type === "validation") {
			return result.response;
		}

		// Process notifications asynchronously via waitUntil
		if (result.notifications && result.notifications.length > 0) {
			const processingPromise = this.processNotifications(
				result.notifications,
				options,
			);

			if (options?.waitUntil) {
				options.waitUntil(processingPromise);
			} else {
				processingPromise.catch((error) => {
					this.config.logger?.error(
						"Notification processing failed",
						error,
					);
				});
			}
		}

		return result.response;
	}

	private async processNotifications(
		notifications: GraphNotification[],
		options?: WebhookOptions,
	): Promise<void> {
		if (!this.graphClient || !this.chat) return;

		for (const notification of notifications) {
			if (notification.changeType !== "created") continue;

			const messageId = extractMessageId(notification);
			const mailbox =
				extractMailbox(notification) || this.botUserId;

			if (!messageId) continue;

			// Deduplication: skip already-processed messages
			if (this.processedMessageIds.has(messageId)) continue;
			this.processedMessageIds.add(messageId);

			// Evict oldest entries when the set grows too large
			if (this.processedMessageIds.size > 10_000) {
				const iterator = this.processedMessageIds.values();
				for (let i = 0; i < 5_000; i++) {
					const val = iterator.next().value;
					if (val) this.processedMessageIds.delete(val);
				}
			}

			try {
				const raw = await this.graphClient.getMessage(
					mailbox,
					messageId,
				);

				// Skip drafts
				if (raw.isDraft) continue;

				// Self-echo detection (bot's own messages)
				if (isSelfEcho(raw, this.mailboxes)) continue;

				// Auto-reply detection (OOF, read receipts)
				if (isAutoReply(raw)) continue;

				const message = parseGraphMessage(raw, {
					mailboxes: this.mailboxes,
					botUserId: this.botUserId,
				});

				// Fetch and map attachments if present
				if (raw.hasAttachments) {
					const attachments =
						await this.graphClient.getAttachments(
							mailbox,
							messageId,
						);
					message.attachments = mapAttachments(attachments.value);
				}

				this.chat.processMessage(
					this,
					message.threadId,
					message,
					options,
				);
			} catch (error) {
				this.config.logger?.error(
					`Failed to process notification for message ${messageId}`,
					error,
				);
			}
		}
	}

	// --- Message Parsing ---

	parseMessage(raw: GraphMessage): Message<GraphMessage> {
		return parseGraphMessage(raw, {
			mailboxes: this.mailboxes,
			botUserId: this.botUserId,
		});
	}

	// --- Messaging ---

	async postMessage(
		threadId: string,
		message: AdapterPostableMessage,
	): Promise<RawMessage<GraphMessage>> {
		if (!this.graphClient) {
			throw new Error("Adapter not initialized");
		}

		const decoded = decodeThreadId(threadId);
		const { html, files } = renderMessage(message);

		// New outbound thread (placeholder from openDM)
		if (decoded.conversationId.startsWith(NEW_THREAD_PREFIX)) {
			return this.sendNewMail(decoded.mailbox, html, files);
		}

		// Reply to existing thread
		const strategy = this.config.replyStrategy ?? "createReply";
		if (strategy === "createReply") {
			return this.replyViaCreateReply(
				decoded.mailbox,
				decoded.conversationId,
				html,
				files,
			);
		}
		return this.replyViaSendMail(
			decoded.mailbox,
			decoded.conversationId,
			html,
			files,
		);
	}

	private async replyViaCreateReply(
		mailbox: string,
		conversationId: string,
		html: string,
		files: Array<{
			data: Buffer | Blob | ArrayBuffer;
			filename: string;
			mimeType?: string;
		}>,
	): Promise<RawMessage<GraphMessage>> {
		if (!this.graphClient) throw new Error("Adapter not initialized");

		// Find latest message to reply to
		const messages = await this.graphClient.getConversationMessages(
			mailbox,
			conversationId,
			{ top: 1, orderBy: "receivedDateTime desc" },
		);

		if (!messages.value.length) {
			throw new Error(
				`No messages found in conversation ${conversationId}`,
			);
		}

		const latestMessage = messages.value[0] as GraphMessage;

		// Create reply draft with body and bot header set upfront
		// (internetMessageHeaders can only be set at creation, not via PATCH)
		const draft = await this.graphClient.createReply(
			mailbox,
			latestMessage.id,
			{
				body: { contentType: "html", content: html },
				internetMessageHeaders: [
					{ name: "X-Chat-SDK-Bot", value: "true" },
				],
			},
		);

		// Attach files
		for (const file of files) {
			const contentBytes = await toBase64(file.data);
			await this.graphClient.addAttachment(mailbox, draft.id, {
				"@odata.type": "#microsoft.graph.fileAttachment",
				name: file.filename,
				contentType: file.mimeType || "application/octet-stream",
				contentBytes,
			});
		}

		// Send the draft
		await this.graphClient.sendDraft(mailbox, draft.id);

		return {
			id: draft.id,
			threadId: encodeThreadId({ mailbox, conversationId }),
			raw: draft,
		};
	}

	private async replyViaSendMail(
		mailbox: string,
		conversationId: string,
		html: string,
		files: Array<{
			data: Buffer | Blob | ArrayBuffer;
			filename: string;
			mimeType?: string;
		}>,
	): Promise<RawMessage<GraphMessage>> {
		if (!this.graphClient) throw new Error("Adapter not initialized");

		// Find latest message for subject and recipient
		const messages = await this.graphClient.getConversationMessages(
			mailbox,
			conversationId,
			{ top: 1, orderBy: "receivedDateTime desc" },
		);

		if (!messages.value.length) {
			throw new Error(
				`No messages found in conversation ${conversationId}`,
			);
		}

		const latest = messages.value[0] as GraphMessage;
		const replyTo = latest.from.emailAddress.address;

		const attachments = await Promise.all(
			files.map(async (file) => ({
				"@odata.type": "#microsoft.graph.fileAttachment",
				name: file.filename,
				contentType: file.mimeType || "application/octet-stream",
				contentBytes: await toBase64(file.data),
			})),
		);

		const headers: Array<{ name: string; value: string }> = [
			{ name: "X-Chat-SDK-Bot", value: "true" },
		];

		if (latest.internetMessageId) {
			headers.push({
				name: "In-Reply-To",
				value: latest.internetMessageId,
			});
		}

		const messagePayload: Record<string, unknown> = {
			subject: `Re: ${latest.subject.replace(/^Re:\s*/i, "")}`,
			body: { contentType: "html", content: html },
			toRecipients: [{ emailAddress: { address: replyTo } }],
			internetMessageHeaders: headers,
		};

		if (attachments.length > 0) {
			messagePayload.attachments = attachments;
		}

		const saveToSent = this.config.saveToSentItems ?? true;
		await this.graphClient.sendMail(mailbox, messagePayload, saveToSent);

		return {
			id: `sent-${Date.now()}`,
			threadId: encodeThreadId({ mailbox, conversationId }),
			raw: latest,
		};
	}

	private async sendNewMail(
		recipientEmail: string,
		html: string,
		files: Array<{
			data: Buffer | Blob | ArrayBuffer;
			filename: string;
			mimeType?: string;
		}>,
	): Promise<RawMessage<GraphMessage>> {
		if (!this.graphClient) throw new Error("Adapter not initialized");

		const attachments = await Promise.all(
			files.map(async (file) => ({
				"@odata.type": "#microsoft.graph.fileAttachment",
				name: file.filename,
				contentType: file.mimeType || "application/octet-stream",
				contentBytes: await toBase64(file.data),
			})),
		);

		const messagePayload: Record<string, unknown> = {
			subject: "New message",
			body: { contentType: "html", content: html },
			toRecipients: [
				{ emailAddress: { address: recipientEmail } },
			],
			internetMessageHeaders: [
				{ name: "X-Chat-SDK-Bot", value: "true" },
			],
		};

		if (attachments.length > 0) {
			messagePayload.attachments = attachments;
		}

		const saveToSent = this.config.saveToSentItems ?? true;
		await this.graphClient.sendMail(
			this.botUserId,
			messagePayload,
			saveToSent,
		);

		const placeholderId = `sent-${Date.now()}`;
		return {
			id: placeholderId,
			threadId: encodeThreadId({
				mailbox: recipientEmail,
				conversationId: `${NEW_THREAD_PREFIX}${placeholderId}`,
			}),
			raw: {} as GraphMessage,
		};
	}

	editMessage(
		_threadId: string,
		_messageId: string,
		_message: AdapterPostableMessage,
	): Promise<RawMessage<GraphMessage>> {
		throw new NotImplementedError("editMessage");
	}

	deleteMessage(_threadId: string, _messageId: string): Promise<void> {
		throw new NotImplementedError("deleteMessage");
	}

	// --- Reactions (not applicable to email) ---

	addReaction(
		_threadId: string,
		_messageId: string,
		_emoji: EmojiValue | string,
	): Promise<void> {
		throw new NotImplementedError("addReaction");
	}

	removeReaction(
		_threadId: string,
		_messageId: string,
		_emoji: EmojiValue | string,
	): Promise<void> {
		throw new NotImplementedError("removeReaction");
	}

	// --- Typing (not applicable to email) ---

	startTyping(_threadId: string, _status?: string): Promise<void> {
		throw new NotImplementedError("startTyping");
	}

	// --- History ---

	async fetchMessages(
		threadId: string,
		options?: FetchOptions,
	): Promise<FetchResult<GraphMessage>> {
		if (!this.graphClient) {
			throw new Error("Adapter not initialized");
		}

		const decoded = decodeThreadId(threadId);

		// New threads created via openDM have no messages yet
		if (decoded.conversationId.startsWith(NEW_THREAD_PREFIX)) {
			return { messages: [] };
		}

		const limit = options?.limit ?? 25;
		const skip = options?.cursor
			? Number.parseInt(options.cursor, 10)
			: 0;

		const result = await this.graphClient.getConversationMessages(
			decoded.mailbox,
			decoded.conversationId,
			{ top: limit, skip },
		);

		const messages = result.value.map((raw) =>
			parseGraphMessage(raw, {
				mailboxes: this.mailboxes,
				botUserId: this.botUserId,
			}),
		);

		const nextCursor =
			result.value.length === limit
				? String(skip + limit)
				: undefined;

		return { messages, nextCursor };
	}

	async fetchThread(threadId: string): Promise<ThreadInfo> {
		const decoded = decodeThreadId(threadId);
		return {
			id: threadId,
			channelId: channelIdFromThreadId(threadId),
			metadata: {
				mailbox: decoded.mailbox,
				conversationId: decoded.conversationId,
			},
		};
	}

	// --- Formatting ---

	renderFormatted(content: FormattedContent): string {
		return this.formatConverter.fromAst(content);
	}

	// --- Optional: New thread creation ---

	async openDM(recipientEmail: string): Promise<string> {
		const placeholder = `${NEW_THREAD_PREFIX}${Date.now()}-${Math.random().toString(36).slice(2, 10)}`;
		return encodeThreadId({
			mailbox: recipientEmail,
			conversationId: placeholder,
		});
	}
}

/**
 * Convert file data to base64 string for Graph API attachment upload.
 */
async function toBase64(data: Buffer | Blob | ArrayBuffer): Promise<string> {
	if (Buffer.isBuffer(data)) {
		return data.toString("base64");
	}
	if (data instanceof ArrayBuffer) {
		return Buffer.from(data).toString("base64");
	}
	// Blob
	const arrayBuffer = await data.arrayBuffer();
	return Buffer.from(arrayBuffer).toString("base64");
}

/**
 * Create a new Outlook Email adapter instance.
 */
export function createOutlookEmailAdapter(
	config: OutlookEmailAdapterConfig = {},
): OutlookEmailAdapter {
	return new OutlookEmailAdapter(config);
}
