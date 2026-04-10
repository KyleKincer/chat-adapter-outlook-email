import { ValidationError } from "@chat-adapter/shared";
import type { OutlookThreadId } from "./types.js";

const PREFIX = "outlook-email";

/**
 * Encode platform-specific thread data into a Chat SDK thread ID string.
 * Format: outlook-email:{base64url(mailbox)}:{base64url(conversationId)}
 */
export function encodeThreadId(platformData: OutlookThreadId): string {
  const encodedMailbox = Buffer.from(platformData.mailbox).toString(
    "base64url"
  );
  const encodedConversationId = Buffer.from(
    platformData.conversationId
  ).toString("base64url");
  return `${PREFIX}:${encodedMailbox}:${encodedConversationId}`;
}

/**
 * Decode a Chat SDK thread ID string back to platform-specific data.
 */
export function decodeThreadId(threadId: string): OutlookThreadId {
  const parts = threadId.split(":");
  if (parts.length !== 3 || parts[0] !== PREFIX) {
    throw new ValidationError(
      PREFIX,
      `Invalid Outlook email thread ID: ${threadId}`
    );
  }
  const mailbox = Buffer.from(parts[1] as string, "base64url").toString(
    "utf-8"
  );
  const conversationId = Buffer.from(
    parts[2] as string,
    "base64url"
  ).toString("utf-8");
  return { mailbox, conversationId };
}

/**
 * Derive the channel ID from a thread ID.
 * The channel represents a mailbox.
 */
export function channelIdFromThreadId(threadId: string): string {
  const { mailbox } = decodeThreadId(threadId);
  return `${PREFIX}:${Buffer.from(mailbox).toString("base64url")}`;
}
