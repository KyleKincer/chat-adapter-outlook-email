import type { Attachment } from "chat";
import { Message, parseMarkdown } from "chat";
import { normalizeEmailBody, stripHtml } from "./body-normalizer.js";
import { encodeThreadId } from "./thread-id.js";
import type { GraphAttachment, GraphMessage } from "./types.js";

export interface MessageParserOptions {
	/** Bot mailbox addresses for self-echo detection */
	mailboxes: string[];
	/** Bot user ID (primary mailbox) */
	botUserId: string;
}

/**
 * Parse a Microsoft Graph message into a Chat SDK Message.
 * Normalizes the body, detects authorship, and maps to the Chat SDK format.
 */
export function parseGraphMessage(
	raw: GraphMessage,
	options: MessageParserOptions,
): Message<GraphMessage> {
	const senderAddress = raw.from.emailAddress.address.toLowerCase();
	const isMe = options.mailboxes.some(
		(m) => m.toLowerCase() === senderAddress,
	);

	// Extract and normalize text content
	const rawText =
		raw.body.contentType === "text"
			? raw.body.content
			: stripHtml(raw.body.content);

	const text = normalizeEmailBody(rawText);

	return new Message<GraphMessage>({
		id: raw.id,
		threadId: encodeThreadId({
			mailbox: options.botUserId,
			conversationId: raw.conversationId,
		}),
		text,
		formatted: parseMarkdown(text),
		raw,
		author: {
			userId: senderAddress,
			userName: senderAddress,
			fullName: raw.from.emailAddress.name,
			isBot: false,
			isMe,
		},
		metadata: {
			dateSent: new Date(raw.receivedDateTime),
			edited: false,
		},
		attachments: [],
		isMention: true, // Email to a mailbox is always a direct mention
	});
}

/**
 * Check if a message was sent by the bot (self-echo).
 * Uses two layers of detection:
 * 1. From address matches a configured bot mailbox
 * 2. Custom X-Chat-SDK-Bot header is present
 */
export function isSelfEcho(
	raw: GraphMessage,
	mailboxes: string[],
): boolean {
	// Layer 1: From address matches bot mailbox
	const senderAddress = raw.from.emailAddress.address.toLowerCase();
	if (mailboxes.some((m) => m.toLowerCase() === senderAddress)) {
		return true;
	}

	// Layer 2: Custom header set on all outbound messages
	if (raw.internetMessageHeaders) {
		const botHeader = raw.internetMessageHeaders.find(
			(h) => h.name.toLowerCase() === "x-chat-sdk-bot",
		);
		if (botHeader?.value === "true") {
			return true;
		}
	}

	return false;
}

/**
 * Check if a message is an auto-reply (OOF, read receipts, etc.).
 * These should be skipped to prevent loops.
 */
export function isAutoReply(raw: GraphMessage): boolean {
	if (!raw.internetMessageHeaders) {
		return false;
	}

	for (const header of raw.internetMessageHeaders) {
		const name = header.name.toLowerCase();

		// RFC 3834 Auto-Submitted header
		if (name === "auto-submitted" && header.value.toLowerCase() !== "no") {
			return true;
		}

		// Outlook Out-of-Office
		if (name === "x-auto-reply-from") {
			return true;
		}

		// Exchange auto-generated messages
		if (name === "x-ms-exchange-generated-message-source") {
			return true;
		}

		// Microsoft auto-response header
		if (
			name === "x-auto-response-suppress" &&
			header.value.toLowerCase() !== "none"
		) {
			return true;
		}
	}

	return false;
}

/**
 * Map Graph attachment objects to Chat SDK Attachment format.
 * Filters out inline images (CID-referenced) for now.
 */
export function mapAttachments(
	graphAttachments: GraphAttachment[],
): Attachment[] {
	return graphAttachments
		.filter((att) => !att.isInline)
		.map((att) => {
			const type = inferAttachmentType(att.contentType);
			const attachment: Attachment = {
				type,
				name: att.name,
				mimeType: att.contentType,
				size: att.size,
			};
			if (att.contentBytes) {
				attachment.data = Buffer.from(att.contentBytes, "base64");
			}
			return attachment;
		});
}

function inferAttachmentType(
	contentType: string,
): "image" | "file" | "video" | "audio" {
	if (contentType.startsWith("image/")) return "image";
	if (contentType.startsWith("video/")) return "video";
	if (contentType.startsWith("audio/")) return "audio";
	return "file";
}
