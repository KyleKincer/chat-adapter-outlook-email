import { ClientSecretCredential } from "@azure/identity";
import { handleGraphError } from "./errors.js";
import type { GraphAttachment, GraphMessage, GraphSubscription } from "./types.js";

const GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0";
const GRAPH_SCOPE = "https://graph.microsoft.com/.default";

/**
 * Thin wrapper around Microsoft Graph REST API.
 * Uses @azure/identity for MSAL client credentials auth.
 * Makes direct fetch() calls rather than using the heavy @microsoft/microsoft-graph-client SDK.
 */
export class GraphClient {
	private credential: ClientSecretCredential;
	private accessToken: string | null = null;
	private tokenExpiry = 0;

	constructor(tenantId: string, clientId: string, clientSecret: string) {
		this.credential = new ClientSecretCredential(
			tenantId,
			clientId,
			clientSecret,
		);
	}

	private async getToken(): Promise<string> {
		// Reuse token if still valid (with 60s buffer)
		if (this.accessToken && Date.now() < this.tokenExpiry - 60_000) {
			return this.accessToken;
		}
		const token = await this.credential.getToken(GRAPH_SCOPE);
		this.accessToken = token.token;
		this.tokenExpiry = token.expiresOnTimestamp;
		return this.accessToken;
	}

	private async request<T>(
		method: string,
		path: string,
		body?: unknown,
	): Promise<T> {
		const token = await this.getToken();
		const url = path.startsWith("http")
			? path
			: `${GRAPH_BASE_URL}${path}`;

		const response = await fetch(url, {
			method,
			headers: {
				Authorization: `Bearer ${token}`,
				"Content-Type": "application/json",
			},
			body: body !== undefined ? JSON.stringify(body) : undefined,
		});

		if (!response.ok) {
			const errorBody = await response.json().catch(() => null);
			const statusCode = response.status;
			const message =
				(errorBody as Record<string, Record<string, string>>)?.error
					?.message || response.statusText;
			const retryAfter = response.headers.get("Retry-After");

			handleGraphError(
				{
					statusCode,
					message,
					...(retryAfter ? { retryAfter: Number(retryAfter) } : {}),
				},
				`${method} ${path}`,
			);
		}

		// No-content responses (204, or 202 with empty body)
		if (response.status === 204 || response.status === 202) {
			return undefined as T;
		}

		const text = await response.text();
		if (!text) {
			return undefined as T;
		}

		return JSON.parse(text) as T;
	}

	// --- Messages ---

	async getMessage(mailbox: string, messageId: string): Promise<GraphMessage> {
		const select = [
			"id",
			"conversationId",
			"subject",
			"body",
			"from",
			"toRecipients",
			"ccRecipients",
			"receivedDateTime",
			"sentDateTime",
			"internetMessageId",
			"internetMessageHeaders",
			"hasAttachments",
			"isRead",
			"isDraft",
			"conversationIndex",
		].join(",");
		return this.request<GraphMessage>(
			"GET",
			`/users/${encodeURIComponent(mailbox)}/messages/${encodeURIComponent(messageId)}?$select=${select}`,
		);
	}

	async getConversationMessages(
		mailbox: string,
		conversationId: string,
		options?: { top?: number; skip?: number; orderBy?: string },
	): Promise<{ value: GraphMessage[]; "@odata.nextLink"?: string }> {
		const top = options?.top ?? 25;
		const orderBy = options?.orderBy ?? "receivedDateTime asc";
		const filter = `conversationId eq '${conversationId}'`;
		const skipParam = options?.skip ? `&$skip=${options.skip}` : "";
		return this.request(
			"GET",
			`/users/${encodeURIComponent(mailbox)}/messages?$filter=${encodeURIComponent(filter)}&$orderby=${encodeURIComponent(orderBy)}&$top=${top}${skipParam}`,
		);
	}

	// --- Reply ---

	async createReply(
		mailbox: string,
		messageId: string,
	): Promise<GraphMessage> {
		return this.request<GraphMessage>(
			"POST",
			`/users/${encodeURIComponent(mailbox)}/messages/${encodeURIComponent(messageId)}/createReply`,
			{},
		);
	}

	async updateMessage(
		mailbox: string,
		messageId: string,
		updates: Record<string, unknown>,
	): Promise<GraphMessage> {
		return this.request<GraphMessage>(
			"PATCH",
			`/users/${encodeURIComponent(mailbox)}/messages/${encodeURIComponent(messageId)}`,
			updates,
		);
	}

	async sendDraft(mailbox: string, draftId: string): Promise<void> {
		return this.request<void>(
			"POST",
			`/users/${encodeURIComponent(mailbox)}/messages/${encodeURIComponent(draftId)}/send`,
		);
	}

	// --- Send new mail ---

	async sendMail(
		mailbox: string,
		message: Record<string, unknown>,
		saveToSentItems = true,
	): Promise<void> {
		return this.request<void>(
			"POST",
			`/users/${encodeURIComponent(mailbox)}/sendMail`,
			{ message, saveToSentItems },
		);
	}

	// --- Subscriptions ---

	async createSubscription(
		subscription: Omit<GraphSubscription, "id">,
	): Promise<GraphSubscription> {
		return this.request<GraphSubscription>(
			"POST",
			"/subscriptions",
			subscription,
		);
	}

	async renewSubscription(
		subscriptionId: string,
		expirationDateTime: string,
	): Promise<GraphSubscription> {
		return this.request<GraphSubscription>(
			"PATCH",
			`/subscriptions/${encodeURIComponent(subscriptionId)}`,
			{ expirationDateTime },
		);
	}

	async deleteSubscription(subscriptionId: string): Promise<void> {
		return this.request<void>(
			"DELETE",
			`/subscriptions/${encodeURIComponent(subscriptionId)}`,
		);
	}

	async listSubscriptions(): Promise<{ value: GraphSubscription[] }> {
		return this.request("GET", "/subscriptions");
	}

	// --- Attachments ---

	async getAttachments(
		mailbox: string,
		messageId: string,
	): Promise<{ value: GraphAttachment[] }> {
		return this.request(
			"GET",
			`/users/${encodeURIComponent(mailbox)}/messages/${encodeURIComponent(messageId)}/attachments`,
		);
	}

	async addAttachment(
		mailbox: string,
		messageId: string,
		attachment: Record<string, unknown>,
	): Promise<GraphAttachment> {
		return this.request<GraphAttachment>(
			"POST",
			`/users/${encodeURIComponent(mailbox)}/messages/${encodeURIComponent(messageId)}/attachments`,
			attachment,
		);
	}
}
