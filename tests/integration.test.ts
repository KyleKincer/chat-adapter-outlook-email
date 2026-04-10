/**
 * Integration tests against a real Microsoft 365 tenant.
 *
 * These tests exercise the full Graph API path: authenticate, send mail,
 * read messages, create replies, and fetch conversation history.
 *
 * Requires .env.test with valid credentials. Skipped if credentials are missing.
 * Run with: pnpm test:integration
 */
import { config } from "dotenv";
import { afterAll, beforeAll, describe, expect, it } from "vitest";
import { GraphClient } from "../src/graph-client.js";
import {
	isSelfEcho,
	isAutoReply,
	parseGraphMessage,
} from "../src/message-parser.js";
import { renderMessage } from "../src/message-renderer.js";
import {
	encodeThreadId,
	decodeThreadId,
} from "../src/thread-id.js";
import type { GraphMessage } from "../src/types.js";

config({ path: ".env.test" });

const tenantId = process.env.OUTLOOK_TENANT_ID;
const clientId = process.env.OUTLOOK_CLIENT_ID;
const clientSecret = process.env.OUTLOOK_CLIENT_SECRET;
const mailbox = process.env.OUTLOOK_MAILBOX;

const canRun = tenantId && clientId && clientSecret && mailbox;

describe.skipIf(!canRun)("Integration: Microsoft Graph API", () => {
	let graph: GraphClient;

	// Track messages we create so we can clean up
	const sentMessageSubjects: string[] = [];
	const testRunId = `test-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;

	beforeAll(() => {
		graph = new GraphClient(tenantId!, clientId!, clientSecret!);
	});

	// --- Authentication ---

	it("authenticates and can access the mailbox", async () => {
		// Simple query to verify auth works — list top 1 message (no filter)
		const result: { value: GraphMessage[] } = await (graph as any).request(
			"GET",
			`/users/${encodeURIComponent(mailbox!)}/messages?$top=1`,
		);
		expect(Array.isArray(result.value)).toBe(true);
	});

	// --- Send Mail ---

	let sentSubject: string;

	it("sends a new email via sendMail", async () => {
		sentSubject = `Integration Test: ${testRunId}`;
		sentMessageSubjects.push(sentSubject);

		const { html } = renderMessage(`Hello from integration test ${testRunId}`);

		await graph.sendMail(
			mailbox!,
			{
				subject: sentSubject,
				body: { contentType: "html", content: html },
				toRecipients: [
					{ emailAddress: { address: mailbox! } },
				],
				internetMessageHeaders: [
					{ name: "X-Chat-SDK-Bot", value: "true" },
					{ name: "X-Test-Run", value: testRunId },
				],
			},
			true,
		);

		// sendMail returns 202 with no body — if we get here it succeeded
		expect(true).toBe(true);
	});

	// --- Read Mail ---

	let receivedMessage: GraphMessage | null = null;

	it("receives the sent email in the mailbox", async () => {
		// Graph mail delivery can take a few seconds — poll until it arrives
		let attempts = 0;
		const maxAttempts = 20;

		while (attempts < maxAttempts) {
			attempts++;

			// Search Inbox specifically (not Sent Items, which doesn't support createReply)
			const searchResult: { value: GraphMessage[] } = await (graph as any).request(
				"GET",
				`/users/${encodeURIComponent(mailbox!)}/mailFolders/Inbox/messages?$filter=${encodeURIComponent(`subject eq '${sentSubject}'`)}&$top=1`,
			);

			if (searchResult.value.length > 0) {
				receivedMessage = searchResult.value[0]!;
				break;
			}

			await new Promise((r) => setTimeout(r, 2000));
		}

		expect(receivedMessage).not.toBeNull();
		expect(receivedMessage!.subject).toBe(sentSubject);
		expect(receivedMessage!.conversationId).toBeTruthy();
	}, 60_000);

	// --- Parse Message ---

	it("parses the received message into a Chat SDK Message", () => {
		expect(receivedMessage).not.toBeNull();

		const parsed = parseGraphMessage(receivedMessage!, {
			mailboxes: [mailbox!],
			botUserId: mailbox!,
		});

		expect(parsed.id).toBe(receivedMessage!.id);
		expect(parsed.text).toContain("Hello from integration test");
		expect(parsed.author.userId).toBe(mailbox!.toLowerCase());
		expect(parsed.author.isMe).toBe(true);
		expect(parsed.threadId).toMatch(/^outlook-email:/);
	});

	// --- Self-Echo Detection ---

	it("detects the sent message as self-echo", () => {
		expect(receivedMessage).not.toBeNull();
		expect(isSelfEcho(receivedMessage!, [mailbox!])).toBe(true);
	});

	it("does not flag the message as auto-reply", () => {
		expect(receivedMessage).not.toBeNull();
		expect(isAutoReply(receivedMessage!)).toBe(false);
	});

	// --- Thread ID Round-trip ---

	it("encodes and decodes thread ID with real conversationId", () => {
		expect(receivedMessage).not.toBeNull();

		const threadId = encodeThreadId({
			mailbox: mailbox!,
			conversationId: receivedMessage!.conversationId,
		});

		const decoded = decodeThreadId(threadId);
		expect(decoded.mailbox).toBe(mailbox!);
		expect(decoded.conversationId).toBe(receivedMessage!.conversationId);
	});

	// --- Reply to Thread ---

	let replyDraftId: string | null = null;

	it("creates a reply draft on the conversation", async () => {
		expect(receivedMessage).not.toBeNull();

		const { html } = renderMessage(`Reply from integration test ${testRunId}`);

		const draft = await graph.createReply(mailbox!, receivedMessage!.id, {
			body: { contentType: "html", content: html },
			internetMessageHeaders: [
				{ name: "X-Chat-SDK-Bot", value: "true" },
			],
		});
		replyDraftId = draft.id;

		expect(draft.id).toBeTruthy();
		expect(draft.isDraft).toBe(true);
	});

	it("sends the reply draft", async () => {
		expect(replyDraftId).not.toBeNull();

		await graph.sendDraft(mailbox!, replyDraftId!);

		// If we get here, the reply was sent successfully
		expect(true).toBe(true);
	});

	// --- Fetch Conversation History ---

	it("fetches conversation messages in chronological order", async () => {
		expect(receivedMessage).not.toBeNull();

		// Wait for the reply to appear
		let messages: GraphMessage[] = [];
		let attempts = 0;

		while (attempts < 15) {
			attempts++;
			const result = await graph.getConversationMessages(
				mailbox!,
				receivedMessage!.conversationId,
				{ top: 10, orderBy: "receivedDateTime asc" },
			);
			messages = result.value;

			if (messages.length >= 2) break;
			await new Promise((r) => setTimeout(r, 2000));
		}

		expect(messages.length).toBeGreaterThanOrEqual(2);

		// Should contain the original message
		const original = messages.find((m) => m.subject === sentSubject);
		expect(original).toBeTruthy();

		// Should contain the reply
		const reply = messages.find((m) => /^Re:/i.test(m.subject));
		expect(reply).toBeTruthy();

		// All messages share the same conversationId
		const convIds = new Set(messages.map((m) => m.conversationId));
		expect(convIds.size).toBe(1);
	}, 45_000);

	// --- Attachments ---

	it("sends a message with an attachment and reads it back", async () => {
		const attachmentSubject = `Attachment Test: ${testRunId}`;
		sentMessageSubjects.push(attachmentSubject);

		const fileContent = Buffer.from("Hello from integration test attachment");

		await graph.sendMail(
			mailbox!,
			{
				subject: attachmentSubject,
				body: { contentType: "text", content: "See attached file." },
				toRecipients: [
					{ emailAddress: { address: mailbox! } },
				],
				attachments: [
					{
						"@odata.type": "#microsoft.graph.fileAttachment",
						name: "test-file.txt",
						contentType: "text/plain",
						contentBytes: fileContent.toString("base64"),
					},
				],
			},
			true,
		);

		// Wait for delivery
		let msg: GraphMessage | null = null;
		let attempts = 0;

		while (attempts < 15) {
			attempts++;
			const searchResult: { value: GraphMessage[] } = await (graph as any).request(
				"GET",
				`/users/${encodeURIComponent(mailbox!)}/messages?$filter=${encodeURIComponent(`subject eq '${attachmentSubject}'`)}&$top=1`,
			);
			if (searchResult.value.length > 0) {
				msg = searchResult.value[0]!;
				break;
			}
			await new Promise((r) => setTimeout(r, 2000));
		}

		expect(msg).not.toBeNull();
		expect(msg!.hasAttachments).toBe(true);

		// Fetch the attachment
		const attachments = await graph.getAttachments(mailbox!, msg!.id);
		expect(attachments.value.length).toBeGreaterThanOrEqual(1);

		const att = attachments.value[0]!;
		expect(att.name).toBe("test-file.txt");
		expect(att.contentType).toBe("text/plain");
		expect(att.contentBytes).toBeTruthy();

		const decoded = Buffer.from(att.contentBytes!, "base64").toString("utf-8");
		expect(decoded).toBe("Hello from integration test attachment");
	}, 45_000);

	// --- Subscription Management ---

	let subscriptionId: string | null = null;

	it("can list subscriptions (initially empty or with existing)", async () => {
		const result = await graph.listSubscriptions();
		expect(Array.isArray(result.value)).toBe(true);
	});

	// Note: subscription creation requires a publicly reachable notificationUrl,
	// so we skip that in CI. The unit tests cover the SubscriptionManager logic.

	// --- Cleanup ---

	afterAll(async () => {
		if (!graph || !mailbox) return;

		// Clean up test messages by searching and deleting them
		for (const subject of sentMessageSubjects) {
			try {
				const searchResult: { value: GraphMessage[] } = await (graph as any).request(
					"GET",
					`/users/${encodeURIComponent(mailbox)}/messages?$filter=${encodeURIComponent(`subject eq '${subject}'`)}&$top=10`,
				);
				for (const msg of searchResult.value) {
					await (graph as any).request(
						"DELETE",
						`/users/${encodeURIComponent(mailbox)}/messages/${encodeURIComponent(msg.id)}`,
					);
				}
				// Also clean up replies
				const replySearchResult: { value: GraphMessage[] } = await (graph as any).request(
					"GET",
					`/users/${encodeURIComponent(mailbox)}/messages?$filter=${encodeURIComponent(`subject eq 'Re: ${subject}'`)}&$top=10`,
				);
				for (const msg of replySearchResult.value) {
					await (graph as any).request(
						"DELETE",
						`/users/${encodeURIComponent(mailbox)}/messages/${encodeURIComponent(msg.id)}`,
					);
				}
			} catch {
				// Best-effort cleanup
			}
		}
	});
});
