import type { Logger } from "chat";
import type { GraphClient } from "./graph-client.js";

export interface SubscriptionInfo {
	id: string;
	mailbox: string;
	expirationDateTime: string;
}

/**
 * Manages Microsoft Graph subscription lifecycle:
 * creation, renewal, reconciliation on startup, and cleanup on disconnect.
 */
export class SubscriptionManager {
	private subscriptions = new Map<string, SubscriptionInfo>();
	private renewalTimer: ReturnType<typeof setInterval> | null = null;

	constructor(
		private readonly graphClient: GraphClient,
		private readonly notificationUrl: string,
		private readonly clientState: string,
		private readonly renewalIntervalMinutes: number,
		private readonly logger?: Logger,
	) {}

	/**
	 * Initialize subscriptions for all configured mailboxes.
	 * Reconciles with existing subscriptions: reuses valid ones,
	 * renews near-expiry ones, and creates missing ones.
	 */
	async initialize(mailboxes: string[]): Promise<void> {
		const existing = await this.graphClient.listSubscriptions();

		// Find subscriptions that belong to us (matching notificationUrl and clientState)
		const ourSubscriptions = existing.value.filter(
			(sub) =>
				sub.notificationUrl === this.notificationUrl &&
				sub.clientState === this.clientState,
		);

		for (const mailbox of mailboxes) {
			const resource = `users/${mailbox}/messages`;
			const existingSub = ourSubscriptions.find(
				(sub) => sub.resource === resource,
			);

			if (existingSub) {
				const expiry = new Date(existingSub.expirationDateTime);
				const hoursUntilExpiry =
					(expiry.getTime() - Date.now()) / (1000 * 60 * 60);

				if (hoursUntilExpiry < 1) {
					this.logger?.info(
						`Renewing near-expiry subscription for ${mailbox}`,
					);
					await this.renewSubscription(existingSub.id, mailbox);
				} else {
					this.logger?.info(
						`Reusing existing subscription for ${mailbox}: ${existingSub.id}`,
					);
					this.subscriptions.set(mailbox, {
						id: existingSub.id,
						mailbox,
						expirationDateTime: existingSub.expirationDateTime,
					});
				}
			} else {
				await this.createSubscription(mailbox);
			}
		}

		this.startRenewalTimer();
	}

	private async createSubscription(mailbox: string): Promise<void> {
		const expirationDateTime = this.getExpirationDateTime();

		try {
			const sub = await this.graphClient.createSubscription({
				resource: `users/${mailbox}/messages`,
				changeType: "created",
				notificationUrl: this.notificationUrl,
				expirationDateTime,
				clientState: this.clientState,
			});

			this.subscriptions.set(mailbox, {
				id: sub.id,
				mailbox,
				expirationDateTime: sub.expirationDateTime,
			});

			this.logger?.info(
				`Created subscription for ${mailbox}: ${sub.id}`,
			);
		} catch (error) {
			this.logger?.error(
				`Failed to create subscription for ${mailbox}`,
				error,
			);
			throw error;
		}
	}

	private async renewSubscription(
		subscriptionId: string,
		mailbox: string,
	): Promise<void> {
		const expirationDateTime = this.getExpirationDateTime();

		try {
			const sub = await this.graphClient.renewSubscription(
				subscriptionId,
				expirationDateTime,
			);

			this.subscriptions.set(mailbox, {
				id: sub.id,
				mailbox,
				expirationDateTime: sub.expirationDateTime,
			});

			this.logger?.info(
				`Renewed subscription for ${mailbox}: ${sub.id}`,
			);
		} catch (error) {
			// Subscription may have been deleted externally — recreate
			this.logger?.warn(
				`Renewal failed for ${mailbox}, recreating subscription`,
				error,
			);
			await this.createSubscription(mailbox);
		}
	}

	/**
	 * Renew all active subscriptions.
	 */
	async renewAll(): Promise<void> {
		for (const [mailbox, info] of this.subscriptions) {
			await this.renewSubscription(info.id, mailbox);
		}
	}

	private startRenewalTimer(): void {
		if (this.renewalTimer) {
			clearInterval(this.renewalTimer);
		}

		// Renew at half the configured interval for a safety margin
		const intervalMs = (this.renewalIntervalMinutes / 2) * 60 * 1000;
		this.renewalTimer = setInterval(() => {
			this.renewAll().catch((error) => {
				this.logger?.error("Subscription renewal failed", error);
			});
		}, intervalMs);

		// Allow the process to exit even if the timer is running
		if (this.renewalTimer && typeof this.renewalTimer === "object" && "unref" in this.renewalTimer) {
			this.renewalTimer.unref();
		}
	}

	/**
	 * Clean up: cancel timers and optionally delete subscriptions.
	 */
	async cleanup(): Promise<void> {
		if (this.renewalTimer) {
			clearInterval(this.renewalTimer);
			this.renewalTimer = null;
		}

		for (const [mailbox, info] of this.subscriptions) {
			try {
				await this.graphClient.deleteSubscription(info.id);
				this.logger?.info(
					`Deleted subscription for ${mailbox}: ${info.id}`,
				);
			} catch (error) {
				this.logger?.warn(
					`Failed to delete subscription for ${mailbox}`,
					error,
				);
			}
		}

		this.subscriptions.clear();
	}

	private getExpirationDateTime(): string {
		const expiry = new Date();
		expiry.setMinutes(expiry.getMinutes() + this.renewalIntervalMinutes);
		return expiry.toISOString();
	}

	/**
	 * Check if a subscription ID belongs to this manager.
	 */
	hasSubscription(subscriptionId: string): boolean {
		for (const info of this.subscriptions.values()) {
			if (info.id === subscriptionId) return true;
		}
		return false;
	}

	/**
	 * Get all active subscription info for debugging/monitoring.
	 */
	getSubscriptions(): SubscriptionInfo[] {
		return Array.from(this.subscriptions.values());
	}
}
