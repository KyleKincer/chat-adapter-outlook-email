import { beforeEach, describe, expect, it, vi } from "vitest";
import { SubscriptionManager } from "../src/subscription-manager.js";
import type { GraphClient } from "../src/graph-client.js";

function createMockGraphClient(): GraphClient {
	return {
		listSubscriptions: vi.fn().mockResolvedValue({ value: [] }),
		createSubscription: vi.fn().mockImplementation((sub) =>
			Promise.resolve({
				id: `sub-${Date.now()}`,
				...sub,
			}),
		),
		renewSubscription: vi.fn().mockImplementation((id, expiry) =>
			Promise.resolve({
				id,
				expirationDateTime: expiry,
			}),
		),
		deleteSubscription: vi.fn().mockResolvedValue(undefined),
	} as unknown as GraphClient;
}

describe("SubscriptionManager", () => {
	let mockClient: GraphClient;
	let manager: SubscriptionManager;

	const NOTIFICATION_URL = "https://example.com/webhook";
	const CLIENT_STATE = "test-secret";
	const RENEWAL_INTERVAL = 4000;

	beforeEach(() => {
		vi.useFakeTimers();
		mockClient = createMockGraphClient();
		manager = new SubscriptionManager(
			mockClient,
			NOTIFICATION_URL,
			CLIENT_STATE,
			RENEWAL_INTERVAL,
		);
	});

	describe("initialize", () => {
		it("creates subscriptions for each mailbox", async () => {
			await manager.initialize(["bot@contoso.com"]);

			expect(mockClient.createSubscription).toHaveBeenCalledWith(
				expect.objectContaining({
					resource: "users/bot@contoso.com/messages",
					changeType: "created",
					notificationUrl: NOTIFICATION_URL,
					clientState: CLIENT_STATE,
				}),
			);
		});

		it("creates subscriptions for multiple mailboxes", async () => {
			await manager.initialize([
				"bot1@contoso.com",
				"bot2@contoso.com",
			]);

			expect(mockClient.createSubscription).toHaveBeenCalledTimes(2);
		});

		it("reuses existing valid subscription", async () => {
			const futureExpiry = new Date();
			futureExpiry.setHours(futureExpiry.getHours() + 24);

			(mockClient.listSubscriptions as ReturnType<typeof vi.fn>).mockResolvedValue({
				value: [
					{
						id: "existing-sub",
						resource: "users/bot@contoso.com/messages",
						changeType: "created",
						notificationUrl: NOTIFICATION_URL,
						clientState: CLIENT_STATE,
						expirationDateTime: futureExpiry.toISOString(),
					},
				],
			});

			await manager.initialize(["bot@contoso.com"]);

			expect(mockClient.createSubscription).not.toHaveBeenCalled();
			expect(manager.hasSubscription("existing-sub")).toBe(true);
		});

		it("renews near-expiry subscription", async () => {
			const nearExpiry = new Date();
			nearExpiry.setMinutes(nearExpiry.getMinutes() + 30); // 30 min from now

			(mockClient.listSubscriptions as ReturnType<typeof vi.fn>).mockResolvedValue({
				value: [
					{
						id: "expiring-sub",
						resource: "users/bot@contoso.com/messages",
						changeType: "created",
						notificationUrl: NOTIFICATION_URL,
						clientState: CLIENT_STATE,
						expirationDateTime: nearExpiry.toISOString(),
					},
				],
			});

			await manager.initialize(["bot@contoso.com"]);

			expect(mockClient.renewSubscription).toHaveBeenCalledWith(
				"expiring-sub",
				expect.any(String),
			);
		});

		it("ignores subscriptions from other apps", async () => {
			(mockClient.listSubscriptions as ReturnType<typeof vi.fn>).mockResolvedValue({
				value: [
					{
						id: "other-sub",
						resource: "users/bot@contoso.com/messages",
						changeType: "created",
						notificationUrl: "https://other-app.com/webhook",
						clientState: "other-secret",
						expirationDateTime: new Date(
							Date.now() + 86400000,
						).toISOString(),
					},
				],
			});

			await manager.initialize(["bot@contoso.com"]);

			expect(mockClient.createSubscription).toHaveBeenCalled();
		});
	});

	describe("cleanup", () => {
		it("deletes all managed subscriptions", async () => {
			await manager.initialize(["bot@contoso.com"]);

			await manager.cleanup();

			expect(mockClient.deleteSubscription).toHaveBeenCalled();
		});

		it("clears subscription tracking", async () => {
			await manager.initialize(["bot@contoso.com"]);
			const subs = manager.getSubscriptions();
			const subId = subs[0]!.id;

			await manager.cleanup();

			expect(manager.hasSubscription(subId)).toBe(false);
			expect(manager.getSubscriptions()).toHaveLength(0);
		});

		it("continues cleanup even if delete fails", async () => {
			(mockClient.deleteSubscription as ReturnType<typeof vi.fn>).mockRejectedValue(
				new Error("Delete failed"),
			);

			await manager.initialize([
				"bot1@contoso.com",
				"bot2@contoso.com",
			]);

			// Should not throw
			await manager.cleanup();

			expect(mockClient.deleteSubscription).toHaveBeenCalledTimes(2);
		});
	});

	describe("hasSubscription", () => {
		it("returns true for managed subscription", async () => {
			(mockClient.createSubscription as ReturnType<typeof vi.fn>).mockResolvedValue({
				id: "my-sub-123",
				resource: "users/bot@contoso.com/messages",
				expirationDateTime: new Date(
					Date.now() + 86400000,
				).toISOString(),
			});

			await manager.initialize(["bot@contoso.com"]);

			expect(manager.hasSubscription("my-sub-123")).toBe(true);
		});

		it("returns false for unknown subscription", async () => {
			await manager.initialize(["bot@contoso.com"]);

			expect(manager.hasSubscription("unknown-sub")).toBe(false);
		});
	});

	describe("renewAll", () => {
		it("renews all managed subscriptions", async () => {
			(mockClient.createSubscription as ReturnType<typeof vi.fn>)
				.mockResolvedValueOnce({
					id: "sub-1",
					resource: "users/bot1@contoso.com/messages",
					expirationDateTime: new Date(
						Date.now() + 86400000,
					).toISOString(),
				})
				.mockResolvedValueOnce({
					id: "sub-2",
					resource: "users/bot2@contoso.com/messages",
					expirationDateTime: new Date(
						Date.now() + 86400000,
					).toISOString(),
				});

			await manager.initialize([
				"bot1@contoso.com",
				"bot2@contoso.com",
			]);

			await manager.renewAll();

			expect(mockClient.renewSubscription).toHaveBeenCalledTimes(2);
		});

		it("recreates subscription if renewal fails", async () => {
			(mockClient.createSubscription as ReturnType<typeof vi.fn>).mockResolvedValue({
				id: "sub-1",
				resource: "users/bot@contoso.com/messages",
				expirationDateTime: new Date(
					Date.now() + 86400000,
				).toISOString(),
			});

			await manager.initialize(["bot@contoso.com"]);

			(mockClient.renewSubscription as ReturnType<typeof vi.fn>).mockRejectedValue(
				new Error("Subscription not found"),
			);

			await manager.renewAll();

			// Should have called createSubscription again (once during init, once during recovery)
			expect(mockClient.createSubscription).toHaveBeenCalledTimes(2);
		});
	});
});
