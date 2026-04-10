import { describe, expect, it } from "vitest";
import {
	extractMailbox,
	extractMessageId,
	handleGraphWebhook,
} from "../src/webhook-handler.js";
import type { GraphNotification } from "../src/types.js";

describe("handleGraphWebhook", () => {
	const CLIENT_STATE = "test-secret";

	describe("validation challenge", () => {
		it("echoes validationToken as text/plain", async () => {
			const request = new Request(
				"https://example.com/webhook?validationToken=abc-123-token",
				{ method: "POST" },
			);
			const result = await handleGraphWebhook(request, CLIENT_STATE);

			expect(result.type).toBe("validation");
			expect(result.response.status).toBe(200);
			expect(result.response.headers.get("Content-Type")).toBe(
				"text/plain",
			);
			expect(await result.response.text()).toBe("abc-123-token");
		});

		it("handles URL-encoded validationToken", async () => {
			const token = "token with spaces+and/special=chars";
			const encoded = encodeURIComponent(token);
			const request = new Request(
				`https://example.com/webhook?validationToken=${encoded}`,
				{ method: "POST" },
			);
			const result = await handleGraphWebhook(request, CLIENT_STATE);

			expect(result.type).toBe("validation");
			expect(await result.response.text()).toBe(token);
		});
	});

	describe("notification handling", () => {
		it("returns 202 for valid notifications", async () => {
			const payload = {
				value: [
					{
						subscriptionId: "sub-001",
						changeType: "created",
						clientState: CLIENT_STATE,
						resource: "users/bot@contoso.com/messages/msg-001",
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
			const result = await handleGraphWebhook(request, CLIENT_STATE);

			expect(result.type).toBe("notification");
			expect(result.response.status).toBe(202);
			expect(result.notifications).toHaveLength(1);
		});

		it("filters notifications with invalid clientState", async () => {
			const payload = {
				value: [
					{
						subscriptionId: "sub-001",
						changeType: "created",
						clientState: "wrong-secret",
						resource: "users/bot@contoso.com/messages/msg-001",
						subscriptionExpirationDateTime:
							"2026-04-17T12:00:00Z",
						tenantId: "tenant-001",
					},
					{
						subscriptionId: "sub-002",
						changeType: "created",
						clientState: CLIENT_STATE,
						resource: "users/bot@contoso.com/messages/msg-002",
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
			const result = await handleGraphWebhook(request, CLIENT_STATE);

			expect(result.notifications).toHaveLength(1);
			expect(result.notifications![0]!.resource).toContain("msg-002");
		});

		it("returns 400 for invalid JSON", async () => {
			const request = new Request("https://example.com/webhook", {
				method: "POST",
				headers: { "Content-Type": "application/json" },
				body: "not valid json",
			});
			const result = await handleGraphWebhook(request, CLIENT_STATE);

			expect(result.response.status).toBe(400);
			expect(result.notifications).toEqual([]);
		});

		it("returns 400 for missing value array", async () => {
			const request = new Request("https://example.com/webhook", {
				method: "POST",
				headers: { "Content-Type": "application/json" },
				body: JSON.stringify({ foo: "bar" }),
			});
			const result = await handleGraphWebhook(request, CLIENT_STATE);

			expect(result.response.status).toBe(400);
			expect(result.notifications).toEqual([]);
		});
	});
});

describe("extractMessageId", () => {
	it("extracts from resourceData.id", () => {
		const notification: GraphNotification = {
			subscriptionId: "sub-001",
			changeType: "created",
			clientState: "test",
			resource: "users/bot@contoso.com/messages/AAMkAGI2TG93AAA=",
			resourceData: {
				"@odata.type": "#Microsoft.Graph.Message",
				"@odata.id":
					"users/bot@contoso.com/messages/AAMkAGI2TG93AAA=",
				id: "AAMkAGI2TG93AAA=",
			},
			subscriptionExpirationDateTime: "2026-04-17T12:00:00Z",
			tenantId: "tenant-001",
		};

		expect(extractMessageId(notification)).toBe("AAMkAGI2TG93AAA=");
	});

	it("falls back to parsing resource path", () => {
		const notification: GraphNotification = {
			subscriptionId: "sub-001",
			changeType: "created",
			clientState: "test",
			resource: "users/bot@contoso.com/messages/msg-123",
			subscriptionExpirationDateTime: "2026-04-17T12:00:00Z",
			tenantId: "tenant-001",
		};

		expect(extractMessageId(notification)).toBe("msg-123");
	});

	it("returns null when no message ID found", () => {
		const notification: GraphNotification = {
			subscriptionId: "sub-001",
			changeType: "created",
			clientState: "test",
			resource: "users/bot@contoso.com",
			subscriptionExpirationDateTime: "2026-04-17T12:00:00Z",
			tenantId: "tenant-001",
		};

		expect(extractMessageId(notification)).toBeNull();
	});
});

describe("extractMailbox", () => {
	it("extracts mailbox from resource path", () => {
		const notification: GraphNotification = {
			subscriptionId: "sub-001",
			changeType: "created",
			clientState: "test",
			resource: "users/bot@contoso.com/messages/msg-123",
			subscriptionExpirationDateTime: "2026-04-17T12:00:00Z",
			tenantId: "tenant-001",
		};

		expect(extractMailbox(notification)).toBe("bot@contoso.com");
	});

	it("returns null for unexpected resource format", () => {
		const notification: GraphNotification = {
			subscriptionId: "sub-001",
			changeType: "created",
			clientState: "test",
			resource: "subscriptions/sub-001",
			subscriptionExpirationDateTime: "2026-04-17T12:00:00Z",
			tenantId: "tenant-001",
		};

		expect(extractMailbox(notification)).toBeNull();
	});
});
