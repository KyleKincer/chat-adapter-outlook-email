# chat-adapter-outlook-email

Microsoft Outlook / M365 email adapter for the [Vercel Chat SDK](https://chat-sdk.dev). Treats email conversations as Chat SDK threads using Microsoft Graph change notifications for inbound and Graph mail APIs for outbound.

Supports both dedicated bot mailboxes and shared mailboxes.

## Install

```bash
pnpm add chat-adapter-outlook-email chat @chat-adapter/shared
```

## Quick Start

```ts
import { Chat } from "chat";
import { createOutlookEmailAdapter } from "chat-adapter-outlook-email";

const adapter = createOutlookEmailAdapter({
  tenantId: process.env.OUTLOOK_TENANT_ID,
  clientId: process.env.OUTLOOK_CLIENT_ID,
  clientSecret: process.env.OUTLOOK_CLIENT_SECRET,
  mailbox: process.env.OUTLOOK_MAILBOX,
  notificationUrl: process.env.OUTLOOK_NOTIFICATION_URL,
  clientState: process.env.OUTLOOK_CLIENT_STATE,
});

const chat = new Chat({
  adapter,
  async onMessage({ message, thread, reply }) {
    // Echo the message back as a threaded reply
    await reply(`You said: ${message.text}`);
  },
});

// Mount the webhook handler on your HTTP server
// POST /api/webhooks/outlook-email → chat.handleWebhook(request)
```

## Configuration

### Environment Variables

| Variable | Required | Description |
|---|---|---|
| `OUTLOOK_TENANT_ID` | Yes | Azure AD / Entra ID tenant ID |
| `OUTLOOK_CLIENT_ID` | Yes | Azure AD application (client) ID |
| `OUTLOOK_CLIENT_SECRET` | Yes | Azure AD client secret |
| `OUTLOOK_MAILBOX` | Yes | Mailbox to monitor (comma-separated for multiple) |
| `OUTLOOK_NOTIFICATION_URL` | Yes | Public HTTPS URL for Graph webhook notifications |
| `OUTLOOK_CLIENT_STATE` | Yes | Secret for webhook verification (max 128 chars) |
| `OUTLOOK_BOT_NAME` | No | Bot display name (defaults to `"Bot"`) |

All configuration can also be passed directly via the `OutlookEmailAdapterConfig` object.

### Azure AD App Registration

1. Register an application in [Azure AD / Entra ID](https://portal.azure.com/#view/Microsoft_AAD_RegisteredApps/ApplicationsListBlade)
2. Add **application permissions**: `Mail.ReadWrite` + `Mail.Send`
3. Grant admin consent for the permissions
4. Create a client secret
5. (Recommended) Scope access to your bot mailbox using [RBAC for Applications in Exchange Online](https://learn.microsoft.com/en-us/graph/auth-limit-mailbox-access) to avoid tenant-wide mailbox access

## Features

### Receive Emails via Webhooks

Inbound emails are received via Microsoft Graph change notifications and delivered as normalized Chat SDK `Message` objects. The adapter handles subscription creation, renewal, and reconciliation automatically.

### Reply to Email Threads

Replies are sent as properly threaded responses using Graph's `createReply` API:

```ts
async onMessage({ message, reply }) {
  await reply("Thanks for your email! We'll get back to you shortly.");
}
```

### Send Emails Proactively

Start new outbound email threads using `openDM`:

```ts
const threadId = await chat.openDM("recipient@example.com");
await chat.postMessage(threadId, "Hello from the bot!");
```

### Markdown and Rich Content

Send formatted messages using markdown:

```ts
await reply({ markdown: "**Important:** Please review the attached document." });
```

Or pass an mdast AST directly:

```ts
await reply({ ast: myMdastRoot });
```

### Attachments

Inbound attachments are mapped to Chat SDK `Attachment` objects with type detection (image, video, audio, file). Outbound files can be attached via the `files` property:

```ts
await reply({
  markdown: "Here's the report you requested.",
  files: [{ data: pdfBuffer, filename: "report.pdf", mimeType: "application/pdf" }],
});
```

### Fetch Message History

Retrieve the full conversation history for an email thread:

```ts
const { messages, nextCursor } = await chat.fetchMessages(threadId, { limit: 50 });
```

### Email Body Normalization

Inbound email bodies are automatically cleaned up:
- Quoted reply chains are stripped (Outlook, Gmail, Apple Mail formats)
- Signatures are removed (`--`, `Sent from my iPhone`, etc.)
- HTML is converted to plain text
- The original HTML is preserved in `message.raw` for handlers that need full fidelity

### Shared Mailbox Support

Shared mailboxes work identically to dedicated mailboxes — just configure the shared mailbox address in `OUTLOOK_MAILBOX`. No additional code or permissions are needed (application permissions access shared mailboxes the same way via Graph API).

## Unsupported Operations

These methods throw `NotImplementedError` as they have no email equivalent:

- `editMessage()` — email cannot be edited after send
- `deleteMessage()` — email cannot be recalled reliably
- `addReaction()` / `removeReaction()` — no email equivalent
- `startTyping()` — no email equivalent

## Loop Prevention

The adapter uses multiple layers to prevent reply loops:

1. **From-address check** — messages from the bot's own mailbox are skipped
2. **Custom header** — outbound messages include `X-Chat-SDK-Bot: true`; inbound messages with this header are skipped
3. **Auto-reply detection** — OOF replies, read receipts, and Exchange-generated messages are filtered via `Auto-Submitted` and related headers

## License

MIT
