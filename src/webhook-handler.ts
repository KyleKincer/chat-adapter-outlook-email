import type { GraphNotification, GraphNotificationPayload } from "./types.js";

export interface WebhookResult {
	type: "validation" | "notification";
	response: Response;
	notifications?: GraphNotification[];
}

/**
 * Handle a Microsoft Graph webhook request.
 *
 * Two request types:
 * 1. Validation challenge: Graph POSTs with ?validationToken=... query param
 *    during subscription creation. Must echo the token as text/plain within 10s.
 * 2. Notification delivery: Graph POSTs a JSON payload with change notifications.
 *    Must return 202 within 3s (recommended) / 10s (hard deadline).
 */
export async function handleGraphWebhook(
	request: Request,
	clientState: string,
): Promise<WebhookResult> {
	const url = new URL(request.url);

	// Step 1: Validation token challenge (subscription creation)
	const validationToken = url.searchParams.get("validationToken");
	if (validationToken) {
		// URL-decode is handled by searchParams.get() automatically.
		// Echo the token back as text/plain within 10 seconds.
		return {
			type: "validation",
			response: new Response(validationToken, {
				status: 200,
				headers: { "Content-Type": "text/plain" },
			}),
		};
	}

	// Step 2: Parse notification payload
	let body: GraphNotificationPayload;
	try {
		body = (await request.json()) as GraphNotificationPayload;
	} catch {
		return {
			type: "notification",
			response: new Response(null, { status: 400 }),
			notifications: [],
		};
	}

	if (!body.value || !Array.isArray(body.value)) {
		return {
			type: "notification",
			response: new Response(null, { status: 400 }),
			notifications: [],
		};
	}

	// Step 3: Validate clientState on each notification to verify authenticity
	const validNotifications = body.value.filter(
		(notification) => notification.clientState === clientState,
	);

	// Step 4: Return 202 Accepted immediately (processing happens asynchronously)
	return {
		type: "notification",
		response: new Response(null, { status: 202 }),
		notifications: validNotifications,
	};
}

/**
 * Extract the message ID from a Graph notification.
 * Tries resourceData.id first, then parses the resource path.
 * Resource path format: "users/{userId}/messages/{messageId}"
 */
export function extractMessageId(
	notification: GraphNotification,
): string | null {
	if (notification.resourceData?.id) {
		return notification.resourceData.id;
	}

	const match = notification.resource.match(/messages\/([^/]+)$/);
	return match ? (match[1] as string) : null;
}

/**
 * Extract the mailbox identifier from a Graph notification resource path.
 * Resource path format: "users/{mailbox}/messages/{messageId}"
 */
export function extractMailbox(
	notification: GraphNotification,
): string | null {
	const match = notification.resource.match(/^users\/([^/]+)\//);
	return match ? (match[1] as string) : null;
}
