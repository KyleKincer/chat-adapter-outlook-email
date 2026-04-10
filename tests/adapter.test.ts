import { describe, expect, it, vi } from "vitest";
import {
	OutlookEmailAdapter,
	createOutlookEmailAdapter,
} from "../src/adapter.js";
import type { GraphMessage } from "../src/types.js";

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

describe("OutlookEmailAdapter", () => {
	describe("constructor", () => {
		it("creates adapter with config", () => {
			const adapter = new OutlookEmailAdapter({
				mailbox: "bot@contoso.com",
				botName: "TestBot",
			});

			expect(adapter.name).toBe("outlook-email");
			expect(adapter.userName).toBe("TestBot");
			expect(adapter.botUserId).toBe("bot@contoso.com");
			expect(adapter.persistMessageHistory).toBe(false);
		});

		it("supports multiple mailboxes as array", () => {
			const adapter = new OutlookEmailAdapter({
				mailbox: ["bot1@contoso.com", "bot2@contoso.com"],
			});

			expect(adapter.botUserId).toBe("bot1@contoso.com");
		});

		it("falls back to env vars for bot name", () => {
			const original = process.env.OUTLOOK_BOT_NAME;
			process.env.OUTLOOK_BOT_NAME = "EnvBot";

			const adapter = new OutlookEmailAdapter({});
			expect(adapter.userName).toBe("EnvBot");

			if (original !== undefined) {
				process.env.OUTLOOK_BOT_NAME = original;
			} else {
				delete process.env.OUTLOOK_BOT_NAME;
			}
		});

		it("defaults bot name to Bot", () => {
			const original = process.env.OUTLOOK_BOT_NAME;
			const original2 = process.env.BOT_USERNAME;
			delete process.env.OUTLOOK_BOT_NAME;
			delete process.env.BOT_USERNAME;

			const adapter = new OutlookEmailAdapter({});
			expect(adapter.userName).toBe("Bot");

			if (original !== undefined)
				process.env.OUTLOOK_BOT_NAME = original;
			if (original2 !== undefined)
				process.env.BOT_USERNAME = original2;
		});
	});

	describe("createOutlookEmailAdapter", () => {
		it("returns an adapter instance", () => {
			const adapter = createOutlookEmailAdapter({
				mailbox: "bot@contoso.com",
			});
			expect(adapter).toBeInstanceOf(OutlookEmailAdapter);
		});

		it("works with empty config", () => {
			const adapter = createOutlookEmailAdapter();
			expect(adapter.name).toBe("outlook-email");
		});
	});

	describe("thread ID delegation", () => {
		const adapter = new OutlookEmailAdapter({
			mailbox: "bot@contoso.com",
		});

		it("encodes and decodes thread IDs", () => {
			const data = {
				mailbox: "bot@contoso.com",
				conversationId: "conv-123",
			};
			const encoded = adapter.encodeThreadId(data);
			const decoded = adapter.decodeThreadId(encoded);
			expect(decoded).toEqual(data);
		});

		it("derives channel ID from thread ID", () => {
			const threadId = adapter.encodeThreadId({
				mailbox: "bot@contoso.com",
				conversationId: "conv-123",
			});
			const channelId = adapter.channelIdFromThreadId(threadId);
			expect(channelId).toMatch(/^outlook-email:/);
		});
	});

	describe("parseMessage", () => {
		it("parses a Graph message into a Chat SDK Message", () => {
			const adapter = new OutlookEmailAdapter({
				mailbox: "bot@contoso.com",
			});
			const raw = makeGraphMessage();
			const msg = adapter.parseMessage(raw);

			expect(msg.id).toBe("msg-001");
			expect(msg.text).toBe("Hello world");
			expect(msg.author.userId).toBe("alice@contoso.com");
			expect(msg.author.isMe).toBe(false);
		});

		it("detects messages from the bot", () => {
			const adapter = new OutlookEmailAdapter({
				mailbox: "bot@contoso.com",
			});
			const raw = makeGraphMessage({
				from: {
					emailAddress: {
						name: "Bot",
						address: "bot@contoso.com",
					},
				},
			});
			const msg = adapter.parseMessage(raw);
			expect(msg.author.isMe).toBe(true);
		});
	});

	describe("fetchThread", () => {
		it("returns thread info with metadata", async () => {
			const adapter = new OutlookEmailAdapter({
				mailbox: "bot@contoso.com",
			});
			const threadId = adapter.encodeThreadId({
				mailbox: "bot@contoso.com",
				conversationId: "conv-123",
			});
			const info = await adapter.fetchThread(threadId);

			expect(info.id).toBe(threadId);
			expect(info.channelId).toMatch(/^outlook-email:/);
			expect(info.metadata).toEqual({
				mailbox: "bot@contoso.com",
				conversationId: "conv-123",
			});
		});
	});

	describe("openDM", () => {
		it("returns an encoded thread ID with placeholder conversation", async () => {
			const adapter = new OutlookEmailAdapter({
				mailbox: "bot@contoso.com",
			});
			const threadId = await adapter.openDM("alice@contoso.com");

			expect(threadId).toMatch(/^outlook-email:/);
			const decoded = adapter.decodeThreadId(threadId);
			expect(decoded.mailbox).toBe("alice@contoso.com");
			expect(decoded.conversationId).toMatch(/^new-/);
		});
	});

	describe("renderFormatted", () => {
		it("converts mdast AST to HTML", () => {
			const adapter = new OutlookEmailAdapter({
				mailbox: "bot@contoso.com",
			});
			const content = {
				type: "root" as const,
				children: [
					{
						type: "paragraph" as const,
						children: [
							{ type: "text" as const, value: "Hello" },
						],
					},
				],
			};
			const html = adapter.renderFormatted(content);
			expect(html).toContain("Hello");
		});
	});

	describe("not-implemented methods", () => {
		const adapter = new OutlookEmailAdapter({
			mailbox: "bot@contoso.com",
		});

		it("editMessage throws NotImplementedError", () => {
			expect(() =>
				adapter.editMessage("t", "m", "msg"),
			).toThrow("not supported");
		});

		it("deleteMessage throws NotImplementedError", () => {
			expect(() => adapter.deleteMessage("t", "m")).toThrow(
				"not supported",
			);
		});

		it("addReaction throws NotImplementedError", () => {
			expect(() =>
				adapter.addReaction("t", "m", "thumbsup"),
			).toThrow("not supported");
		});

		it("removeReaction throws NotImplementedError", () => {
			expect(() =>
				adapter.removeReaction("t", "m", "thumbsup"),
			).toThrow("not supported");
		});

		it("startTyping throws NotImplementedError", () => {
			expect(() => adapter.startTyping("t")).toThrow(
				"not supported",
			);
		});
	});

	describe("initialize", () => {
		it("throws when credentials are missing", async () => {
			const adapter = new OutlookEmailAdapter({});
			const mockChat = {
				processMessage: vi.fn(),
			};

			// Clear env vars
			const origTenant = process.env.OUTLOOK_TENANT_ID;
			const origClient = process.env.OUTLOOK_CLIENT_ID;
			const origSecret = process.env.OUTLOOK_CLIENT_SECRET;
			delete process.env.OUTLOOK_TENANT_ID;
			delete process.env.OUTLOOK_CLIENT_ID;
			delete process.env.OUTLOOK_CLIENT_SECRET;

			await expect(adapter.initialize(mockChat)).rejects.toThrow(
				"Missing required credentials",
			);

			// Restore
			if (origTenant) process.env.OUTLOOK_TENANT_ID = origTenant;
			if (origClient) process.env.OUTLOOK_CLIENT_ID = origClient;
			if (origSecret) process.env.OUTLOOK_CLIENT_SECRET = origSecret;
		});
	});

	describe("handleWebhook", () => {
		it("handles validation token challenge", async () => {
			const adapter = new OutlookEmailAdapter({
				mailbox: "bot@contoso.com",
				clientState: "test-secret",
			});

			const request = new Request(
				"https://example.com/webhook?validationToken=test-token",
				{ method: "POST" },
			);
			const response = await adapter.handleWebhook(request);

			expect(response.status).toBe(200);
			expect(await response.text()).toBe("test-token");
		});

		it("returns 202 for notification payload", async () => {
			const adapter = new OutlookEmailAdapter({
				mailbox: "bot@contoso.com",
				clientState: "test-secret",
			});

			const payload = {
				value: [
					{
						subscriptionId: "sub-001",
						changeType: "created",
						clientState: "test-secret",
						resource:
							"users/bot@contoso.com/messages/msg-001",
						subscriptionExpirationDateTime:
							"2026-04-17T12:00:00Z",
						tenantId: "tenant-001",
					},
				],
			};
			const request = new Request("https://example.com/webhook", {
				method: "POST",
				headers: { "Content-Type": "application/json" },
				body: JSON.stringify(payload),
			});
			const response = await adapter.handleWebhook(request);

			expect(response.status).toBe(202);
		});
	});

	describe("postMessage without initialization", () => {
		it("throws when not initialized", async () => {
			const adapter = new OutlookEmailAdapter({
				mailbox: "bot@contoso.com",
			});
			const threadId = adapter.encodeThreadId({
				mailbox: "bot@contoso.com",
				conversationId: "conv-123",
			});

			await expect(
				adapter.postMessage(threadId, "Hello"),
			).rejects.toThrow("Adapter not initialized");
		});
	});

	describe("fetchMessages without initialization", () => {
		it("throws when not initialized", async () => {
			const adapter = new OutlookEmailAdapter({
				mailbox: "bot@contoso.com",
			});
			const threadId = adapter.encodeThreadId({
				mailbox: "bot@contoso.com",
				conversationId: "conv-123",
			});

			await expect(
				adapter.fetchMessages(threadId),
			).rejects.toThrow("Adapter not initialized");
		});
	});
});
