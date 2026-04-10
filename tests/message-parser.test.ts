import { describe, expect, it } from "vitest";
import {
	isAutoReply,
	isSelfEcho,
	mapAttachments,
	parseGraphMessage,
} from "../src/message-parser.js";
import type { GraphAttachment, GraphMessage } from "../src/types.js";

function makeGraphMessage(
	overrides: Partial<GraphMessage> = {},
): GraphMessage {
	return {
		id: "msg-001",
		conversationId: "conv-001",
		subject: "Test Subject",
		body: { contentType: "text", content: "Hello world" },
		from: {
			emailAddress: { name: "Alice", address: "alice@contoso.com" },
		},
		toRecipients: [
			{
				emailAddress: {
					name: "Bot",
					address: "bot@contoso.com",
				},
			},
		],
		receivedDateTime: "2026-04-10T12:00:00Z",
		sentDateTime: "2026-04-10T11:59:55Z",
		internetMessageId: "<ABC123@contoso.com>",
		hasAttachments: false,
		isRead: false,
		isDraft: false,
		...overrides,
	};
}

const PARSER_OPTIONS = {
	mailboxes: ["bot@contoso.com"],
	botUserId: "bot@contoso.com",
};

describe("parseGraphMessage", () => {
	it("parses a plain text email into a Message", () => {
		const raw = makeGraphMessage();
		const msg = parseGraphMessage(raw, PARSER_OPTIONS);

		expect(msg.id).toBe("msg-001");
		expect(msg.text).toBe("Hello world");
		expect(msg.author.userId).toBe("alice@contoso.com");
		expect(msg.author.fullName).toBe("Alice");
		expect(msg.author.isMe).toBe(false);
		expect(msg.isMention).toBe(true);
	});

	it("strips HTML for HTML emails", () => {
		const raw = makeGraphMessage({
			body: {
				contentType: "html",
				content: "<p>Hello <strong>world</strong></p>",
			},
		});
		const msg = parseGraphMessage(raw, PARSER_OPTIONS);

		expect(msg.text).toContain("Hello");
		expect(msg.text).toContain("world");
		expect(msg.text).not.toContain("<p>");
		expect(msg.text).not.toContain("<strong>");
	});

	it("normalizes quoted content in email bodies", () => {
		const raw = makeGraphMessage({
			body: {
				contentType: "text",
				content:
					"Sure thing!\n\nOn Thu, Apr 10, 2026 Alice wrote:\n> Original message",
			},
		});
		const msg = parseGraphMessage(raw, PARSER_OPTIONS);

		expect(msg.text).toBe("Sure thing!");
	});

	it("sets isMe=true when sender matches bot mailbox", () => {
		const raw = makeGraphMessage({
			from: {
				emailAddress: {
					name: "Bot",
					address: "bot@contoso.com",
				},
			},
		});
		const msg = parseGraphMessage(raw, PARSER_OPTIONS);

		expect(msg.author.isMe).toBe(true);
	});

	it("encodes thread ID with bot mailbox and conversationId", () => {
		const raw = makeGraphMessage({ conversationId: "conv-xyz" });
		const msg = parseGraphMessage(raw, PARSER_OPTIONS);

		expect(msg.threadId).toMatch(/^outlook-email:/);
		expect(msg.threadId).toContain(":");
	});

	it("sets dateSent from receivedDateTime", () => {
		const raw = makeGraphMessage({
			receivedDateTime: "2026-04-10T15:30:00Z",
		});
		const msg = parseGraphMessage(raw, PARSER_OPTIONS);

		expect(msg.metadata.dateSent.toISOString()).toBe(
			"2026-04-10T15:30:00.000Z",
		);
	});

	it("preserves raw Graph message", () => {
		const raw = makeGraphMessage();
		const msg = parseGraphMessage(raw, PARSER_OPTIONS);

		expect(msg.raw).toBe(raw);
	});
});

describe("isSelfEcho", () => {
	it("returns true when from address matches bot mailbox", () => {
		const raw = makeGraphMessage({
			from: {
				emailAddress: {
					name: "Bot",
					address: "bot@contoso.com",
				},
			},
		});
		expect(isSelfEcho(raw, ["bot@contoso.com"])).toBe(true);
	});

	it("returns true when X-Chat-SDK-Bot header is present", () => {
		const raw = makeGraphMessage({
			internetMessageHeaders: [
				{ name: "X-Chat-SDK-Bot", value: "true" },
			],
		});
		expect(isSelfEcho(raw, ["bot@contoso.com"])).toBe(true);
	});

	it("is case-insensitive for mailbox check", () => {
		const raw = makeGraphMessage({
			from: {
				emailAddress: {
					name: "Bot",
					address: "BOT@Contoso.com",
				},
			},
		});
		expect(isSelfEcho(raw, ["bot@contoso.com"])).toBe(true);
	});

	it("returns false for normal external email", () => {
		const raw = makeGraphMessage();
		expect(isSelfEcho(raw, ["bot@contoso.com"])).toBe(false);
	});

	it("checks multiple configured mailboxes", () => {
		const raw = makeGraphMessage({
			from: {
				emailAddress: {
					name: "Bot 2",
					address: "bot2@contoso.com",
				},
			},
		});
		expect(
			isSelfEcho(raw, ["bot@contoso.com", "bot2@contoso.com"]),
		).toBe(true);
	});
});

describe("isAutoReply", () => {
	it("returns true for Auto-Submitted header", () => {
		const raw = makeGraphMessage({
			internetMessageHeaders: [
				{ name: "Auto-Submitted", value: "auto-replied" },
			],
		});
		expect(isAutoReply(raw)).toBe(true);
	});

	it("returns false when Auto-Submitted is 'no'", () => {
		const raw = makeGraphMessage({
			internetMessageHeaders: [
				{ name: "Auto-Submitted", value: "no" },
			],
		});
		expect(isAutoReply(raw)).toBe(false);
	});

	it("returns true for X-Auto-Reply-From header", () => {
		const raw = makeGraphMessage({
			internetMessageHeaders: [
				{
					name: "X-Auto-Reply-From",
					value: "alice@contoso.com",
				},
			],
		});
		expect(isAutoReply(raw)).toBe(true);
	});

	it("returns true for Exchange generated messages", () => {
		const raw = makeGraphMessage({
			internetMessageHeaders: [
				{
					name: "X-MS-Exchange-Generated-Message-Source",
					value: "Mailbox Rules Agent",
				},
			],
		});
		expect(isAutoReply(raw)).toBe(true);
	});

	it("returns false when no headers present", () => {
		const raw = makeGraphMessage({ internetMessageHeaders: undefined });
		expect(isAutoReply(raw)).toBe(false);
	});

	it("returns false for normal email headers", () => {
		const raw = makeGraphMessage({
			internetMessageHeaders: [
				{ name: "Content-Type", value: "text/html" },
			],
		});
		expect(isAutoReply(raw)).toBe(false);
	});
});

describe("mapAttachments", () => {
	it("maps file attachments correctly", () => {
		const graphAttachments: GraphAttachment[] = [
			{
				id: "att-001",
				name: "report.pdf",
				contentType: "application/pdf",
				size: 12345,
				isInline: false,
				contentBytes: Buffer.from("pdf content").toString("base64"),
			},
		];

		const result = mapAttachments(graphAttachments);
		expect(result).toHaveLength(1);
		expect(result[0]!.type).toBe("file");
		expect(result[0]!.name).toBe("report.pdf");
		expect(result[0]!.mimeType).toBe("application/pdf");
		expect(result[0]!.size).toBe(12345);
		expect(result[0]!.data).toBeInstanceOf(Buffer);
	});

	it("maps image attachments with correct type", () => {
		const graphAttachments: GraphAttachment[] = [
			{
				id: "att-002",
				name: "photo.png",
				contentType: "image/png",
				size: 5000,
				isInline: false,
			},
		];

		const result = mapAttachments(graphAttachments);
		expect(result[0]!.type).toBe("image");
	});

	it("maps video attachments with correct type", () => {
		const graphAttachments: GraphAttachment[] = [
			{
				id: "att-003",
				name: "clip.mp4",
				contentType: "video/mp4",
				size: 50000,
				isInline: false,
			},
		];

		const result = mapAttachments(graphAttachments);
		expect(result[0]!.type).toBe("video");
	});

	it("maps audio attachments with correct type", () => {
		const graphAttachments: GraphAttachment[] = [
			{
				id: "att-004",
				name: "recording.mp3",
				contentType: "audio/mpeg",
				size: 30000,
				isInline: false,
			},
		];

		const result = mapAttachments(graphAttachments);
		expect(result[0]!.type).toBe("audio");
	});

	it("filters out inline attachments", () => {
		const graphAttachments: GraphAttachment[] = [
			{
				id: "att-inline",
				name: "logo.png",
				contentType: "image/png",
				size: 1000,
				isInline: true,
				contentId: "cid:logo",
			},
			{
				id: "att-file",
				name: "doc.pdf",
				contentType: "application/pdf",
				size: 5000,
				isInline: false,
			},
		];

		const result = mapAttachments(graphAttachments);
		expect(result).toHaveLength(1);
		expect(result[0]!.name).toBe("doc.pdf");
	});

	it("handles empty attachment list", () => {
		expect(mapAttachments([])).toEqual([]);
	});

	it("handles attachments without contentBytes", () => {
		const graphAttachments: GraphAttachment[] = [
			{
				id: "att-005",
				name: "file.txt",
				contentType: "text/plain",
				size: 100,
				isInline: false,
			},
		];

		const result = mapAttachments(graphAttachments);
		expect(result[0]!.data).toBeUndefined();
	});
});
