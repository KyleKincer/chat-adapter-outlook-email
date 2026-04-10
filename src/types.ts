import type { Logger } from "chat";

/**
 * Decoded thread identifier for Outlook email threads.
 * Encoded format: `outlook-email:{base64url(mailbox)}:{base64url(conversationId)}`
 */
export interface OutlookThreadId {
  /** The mailbox address (e.g., bot@contoso.com) */
  mailbox: string;
  /** Graph conversation ID (groups related messages in a thread) */
  conversationId: string;
}

/**
 * Raw Microsoft Graph message object, as returned by the Graph API.
 * This is the TRawMessage type parameter for our adapter.
 */
export interface GraphMessage {
  id: string;
  conversationId: string;
  subject: string;
  body: {
    contentType: "text" | "html";
    content: string;
  };
  from: {
    emailAddress: {
      name: string;
      address: string;
    };
  };
  toRecipients: Array<{
    emailAddress: {
      name: string;
      address: string;
    };
  }>;
  ccRecipients?: Array<{
    emailAddress: {
      name: string;
      address: string;
    };
  }>;
  receivedDateTime: string;
  sentDateTime: string;
  internetMessageId: string;
  internetMessageHeaders?: Array<{
    name: string;
    value: string;
  }>;
  hasAttachments: boolean;
  isRead: boolean;
  isDraft: boolean;
  conversationIndex?: string;
}

/**
 * Graph attachment object.
 */
export interface GraphAttachment {
  id: string;
  name: string;
  contentType: string;
  size: number;
  isInline: boolean;
  contentId?: string;
  contentBytes?: string;
  contentLocation?: string;
}

/**
 * Graph change notification payload.
 */
export interface GraphNotificationPayload {
  value: GraphNotification[];
}

export interface GraphNotification {
  subscriptionId: string;
  changeType: "created" | "updated" | "deleted";
  clientState: string;
  resource: string;
  resourceData?: {
    "@odata.type": string;
    "@odata.id": string;
    id: string;
  };
  subscriptionExpirationDateTime: string;
  tenantId: string;
}

/**
 * Graph subscription object.
 */
export interface GraphSubscription {
  id: string;
  resource: string;
  changeType: string;
  notificationUrl: string;
  expirationDateTime: string;
  clientState: string;
}

/**
 * Adapter configuration.
 */
export interface OutlookEmailAdapterConfig {
  /** Azure AD tenant ID. Falls back to OUTLOOK_TENANT_ID env var. */
  tenantId?: string;

  /** Azure AD application (client) ID. Falls back to OUTLOOK_CLIENT_ID env var. */
  clientId?: string;

  /** Azure AD client secret. Falls back to OUTLOOK_CLIENT_SECRET env var. */
  clientSecret?: string;

  /**
   * Target mailbox address(es) to monitor.
   * Falls back to OUTLOOK_MAILBOX env var.
   */
  mailbox?: string | string[];

  /**
   * Public HTTPS URL that receives Graph notifications.
   * Falls back to OUTLOOK_NOTIFICATION_URL env var.
   */
  notificationUrl?: string;

  /**
   * Secret included in subscription for webhook verification.
   * Falls back to OUTLOOK_CLIENT_STATE env var. Max 128 characters.
   */
  clientState?: string;

  /**
   * Bot display name. Falls back to OUTLOOK_BOT_NAME or BOT_USERNAME env var.
   * Defaults to "Bot".
   */
  botName?: string;

  /** Folder to monitor. Defaults to "inbox". */
  folder?: string;

  /**
   * Reply strategy for postMessage on existing threads.
   * - "createReply": Create draft reply, update body, then send (default)
   * - "sendMail": Use sendMail with threading headers
   */
  replyStrategy?: "createReply" | "sendMail";

  /** Save outbound messages to Sent Items. Defaults to true. */
  saveToSentItems?: boolean;

  /**
   * Subscription renewal interval in minutes.
   * Defaults to 4000 (under the 7-day max).
   */
  renewalIntervalMinutes?: number;

  /** Logger instance. */
  logger?: Logger;
}
