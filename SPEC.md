# Outlook Email Adapter for Vercel Chat SDK

## Specification Document

**Package:** `chat-adapter-outlook-email`
**Status:** Implemented (v0.1.0)
**Last Updated:** 2026-04-10

---

## 1. Overview

This adapter integrates Microsoft 365 / Outlook email with the [Vercel Chat SDK](https://chat-sdk.dev) (`chat` npm package). It treats email conversations as Chat SDK threads, uses **Microsoft Graph change notifications** (webhooks) for inbound mail events, and **Graph mail APIs** for outbound replies and new messages.

### Goals

- Receive inbound emails via Microsoft Graph change notifications and surface them as normalized Chat SDK `Message` objects
- Reply to email threads using Graph's `createReply` API
- Start new outbound email threads using Graph's `sendMail` API
- Fetch message history from email threads
- Handle attachments bidirectionally
- Manage Graph subscription lifecycle (creation, renewal, cleanup)
- Prevent reply loops and deduplicate notifications
- Support both dedicated bot mailboxes and shared mailboxes

### Non-Goals

- Reactions, typing indicators, ephemeral messages, modals, streaming (email has no equivalent)
- Perfect HTML round-tripping or Outlook rich-text fidelity
- Delegated (user-context) permissions (application permissions only)
- Calendar, contacts, or other Outlook resources
- Multi-tenant SaaS deployment patterns

---

## 2. Architecture

### Layer Diagram

```
                   Inbound                          Outbound
                     |                                 |
    [Microsoft Graph Webhooks]              [Chat SDK Handler]
                     |                                 |
                     v                                 v
    +----------------------------------+  +---------------------------+
    |        WebhookHandler            |  |     OutlookEmailAdapter   |
    |  - validation token challenge    |  |  - postMessage()          |
    |  - clientState verification      |  |  - openDM()               |
    |  - notification parsing          |  |  - fetchMessages()        |
    +----------------------------------+  +---------------------------+
                     |                                 |
                     v                                 v
    +----------------------------------+  +---------------------------+
    |        MessageParser             |  |     MessageRenderer       |
    |  - Graph message hydration       |  |  - mdast -> HTML          |
    |  - body normalization            |  |  - card rendering         |
    |  - attachment mapping            |  |  - plain text fallback    |
    +----------------------------------+  +---------------------------+
                     |                                 |
                     +--------+    +-------------------+
                              |    |
                              v    v
                     +------------------+
                     |   GraphClient    |
                     |  - auth (MSAL)   |
                     |  - subscriptions |
                     |  - messages      |
                     |  - send/reply    |
                     |  - attachments   |
                     +------------------+
                              |
                     +------------------+
                     | SubscriptionMgr  |
                     |  - create        |
                     |  - renew         |
                     |  - reconcile     |
                     |  - cleanup       |
                     +------------------+
```

### Four Layers

1. **Adapter Layer** - Implements `Adapter<OutlookThreadId, GraphMessage>` from Chat SDK. Entry point for all SDK interactions.

2. **Graph Client Layer** - Thin wrapper around Microsoft Graph REST API using `@azure/identity` for MSAL auth. Handles subscriptions, message CRUD, and mail send/reply.

3. **Webhook Ingestion Layer** - Handles Graph validation challenges, verifies `clientState`, parses notification payloads, and dispatches to the adapter for processing.

4. **Subscription Management Layer** - Creates, renews, and reconciles Graph subscriptions. Runs renewal before expiry. Cleans up on disconnect.

---

## 3. Chat SDK Adapter Interface Contract

Implements `Adapter<TThreadId, TRawMessage>` from the `chat` package:

```typescript
// TThreadId = OutlookThreadId (decoded thread identifier)
// TRawMessage = GraphMessage (raw Microsoft Graph message object)
```

### Functional Methods

```typescript
// Lifecycle
initialize(chat: ChatInstance): Promise<void>
disconnect(): Promise<void>

// Thread ID
encodeThreadId(platformData: OutlookThreadId): string
decodeThreadId(threadId: string): OutlookThreadId
channelIdFromThreadId(threadId: string): string

// Webhook
handleWebhook(request: Request, options?: WebhookOptions): Promise<Response>

// Message parsing
parseMessage(raw: GraphMessage): Message<GraphMessage>

// Messaging
postMessage(threadId: string, message: AdapterPostableMessage): Promise<RawMessage<GraphMessage>>

// History
fetchMessages(threadId: string, options?: FetchOptions): Promise<FetchResult<GraphMessage>>
fetchThread(threadId: string): Promise<ThreadInfo>

// Formatting
renderFormatted(content: FormattedContent): string

// New thread creation
openDM(email: string): Promise<string>  // Returns encoded threadId for a new outbound thread
```

### Methods That Throw NotImplementedError

These are required by the interface but have no email equivalent:

```typescript
editMessage()      // Email cannot be edited after send
deleteMessage()    // Email cannot be recalled reliably
addReaction()      // No email equivalent
removeReaction()   // No email equivalent
startTyping()      // No email equivalent
```

### Readonly Properties

```typescript
readonly name = "outlook-email"
readonly userName: string  // Bot's display name (from config)
readonly botUserId: string // Bot mailbox address
readonly persistMessageHistory = false  // Graph is the source of truth
```

---

## 4. Microsoft Graph API Surface

All API calls target Microsoft Graph v1.0.

### Authentication

Uses **client credentials flow** (application permissions) via `@azure/identity` `ClientSecretCredential`:

```
POST https://login.microsoftonline.com/{tenantId}/oauth2/v2.0/token
scope=https://graph.microsoft.com/.default
```

### Required Application Permissions

| Permission | Reason |
|---|---|
| `Mail.ReadWrite` | Read messages, create reply drafts, send drafts |
| `Mail.Send` | Send new outbound emails via `sendMail` |

**Admin consent required** for application permissions.

Per-operation breakdown of least-privileged permissions:

| Operation | Least Privileged Permission |
|---|---|
| Subscribe to mail notifications | `Mail.ReadBasic` or `Mail.Read` |
| Read message content | `Mail.Read` |
| Create a reply draft | `Mail.ReadWrite` |
| Send a draft | `Mail.ReadWrite` |
| Send new mail | `Mail.Send` |

### API Endpoints Used

#### Subscriptions

```
POST   /subscriptions                          # Create subscription
PATCH  /subscriptions/{id}                     # Renew subscription
DELETE /subscriptions/{id}                     # Delete subscription
GET    /subscriptions                          # List active subscriptions
```

Subscription resource for inbox messages:
```
resource: "users/{mailbox}/messages"
changeType: "created"
notificationUrl: "https://{host}/api/webhooks/outlook-email"
expirationDateTime: <now + 4230 minutes>  (max 10,080 min / ~7 days for basic)
clientState: "{random secret}"
```

**Limits:**
- Max 1,000 active subscriptions per mailbox across all applications
- Basic notification max lifetime: 10,080 minutes (~7 days)
- Rich notification (with resource data) max lifetime: 1,440 minutes (~1 day)

#### Messages

```
GET  /users/{mailbox}/messages/{id}                    # Get single message
GET  /users/{mailbox}/messages?$filter=conversationId eq '{id}'&$orderby=receivedDateTime asc
                                                        # List messages in conversation
```

#### Send / Reply

```
POST /users/{mailbox}/messages/{id}/createReply         # Create reply draft -> returns Message (isDraft: true)
POST /users/{mailbox}/messages/{draftId}/send            # Send the draft
POST /users/{mailbox}/sendMail                           # Send new mail (returns 202 Accepted)
```

#### Attachments

```
GET  /users/{mailbox}/messages/{id}/attachments          # List attachments
GET  /users/{mailbox}/messages/{id}/attachments/{attId}  # Get attachment content
POST /users/{mailbox}/messages/{id}/attachments          # Add attachment to draft
```

### Webhook Notification Payload

```typescript
interface GraphNotificationPayload {
  value: GraphNotification[];
}

interface GraphNotification {
  subscriptionId: string;
  changeType: "created" | "updated" | "deleted";
  clientState: string;
  resource: string;              // e.g., "users/{id}/messages/{messageId}"
  resourceData?: {
    "@odata.type": string;
    "@odata.id": string;
    id: string;
  };
  subscriptionExpirationDateTime: string;
  tenantId: string;
}
```

### Webhook Timing Requirements

Graph enforces strict timing on webhook responses:
- **3 seconds:** Recommended target for returning `200 OK` when processing inline
- **10 seconds:** Hard deadline — no response within 10s triggers delivery failure + retry
- **Slow state:** If >10% of responses exceed 10s in a 10-minute window, new notifications are delayed by 10s
- **Drop state:** If >15% of responses exceed 10s in a 10-minute window, notifications are dropped for up to 10 minutes

The adapter handles this by returning `202 Accepted` immediately and processing notifications asynchronously via `waitUntil`.

### Webhook Validation Challenge

When creating a subscription, Graph validates the endpoint by sending the `validationToken` as a **query parameter** on a POST request (not in the body). The adapter must URL-decode it and echo it back as `text/plain` within 10 seconds.

---

## 5. Shared Mailbox Support

Shared mailboxes are fully supported. From the Graph API's perspective, a shared mailbox is accessed identically to a regular user mailbox — the same `/users/{mailbox}/messages` endpoints apply. No code-level distinction is needed.

### How It Works

- A shared mailbox in Exchange Online has an associated (disabled) user account with its own email address and object ID
- Configure the shared mailbox address in `OUTLOOK_MAILBOX` (or `config.mailbox`) exactly as you would a dedicated bot mailbox
- All Graph operations (read, write, send, subscribe) work identically via application permissions

### Permission Scoping

Application permissions (`Mail.ReadWrite`, `Mail.Send`) are **tenant-wide by default** — granting them gives the app access to every mailbox in the organization. For production deployments, scope access to only the required mailbox(es):

**RBAC for Applications in Exchange Online (recommended):**
1. Create a service principal in Exchange Online
2. Define a management scope targeting the shared mailbox
3. Assign application roles scoped to that management scope
4. Remove the broad Microsoft Entra ID permission consent

```powershell
New-ServicePrincipal -AppId <appId> -ObjectId <objectId> -DisplayName "ChatBot"
New-ManagementScope -Name "BotMailboxOnly" -RecipientRestrictionFilter "Alias -eq 'botmailbox'"
New-ManagementRoleAssignment -App <objectId> -Role "Application Mail.ReadWrite" -CustomResourceScope "BotMailboxOnly"
New-ManagementRoleAssignment -App <objectId> -Role "Application Mail.Send" -CustomResourceScope "BotMailboxOnly"
```

> **Note:** Permission changes can take 30 minutes to 2 hours to propagate.

### Licensing

- Shared mailboxes do **not** require an Exchange Online license for basic usage (under 50 GB)
- A license is required only for mailboxes over 50 GB, in-place archiving, or litigation hold
- Sign-in should remain **blocked** on the shared mailbox's user account — this does not affect API access via application permissions

### Sent Items Behavior

When sending via `/users/{shared-mailbox}/sendMail`, the message is saved to the **shared mailbox's** Sent Items folder (not any individual user's). This is the expected behavior and requires no special configuration.

---

## 6. Thread ID Design

### Decoded Type

```typescript
interface OutlookThreadId {
  /** The mailbox address (e.g., bot@contoso.com) */
  mailbox: string;
  /** Graph conversation ID (groups related messages) */
  conversationId: string;
}
```

### Encoding

Following the Chat SDK convention of `{adapter}:{channel}:{thread}`:

```
outlook-email:{base64url(mailbox)}:{base64url(conversationId)}
```

Using `base64url` encoding to safely handle special characters in both the mailbox address and Graph's conversationId.

### Examples

```
outlook-email:Ym90QGNvbnRvc28uY29t:QUFRa0FEQXdNREF3TURZM...
```

### Channel ID

The channel represents a mailbox:

```typescript
channelIdFromThreadId(threadId: string): string {
  const { mailbox } = this.decodeThreadId(threadId);
  return `outlook-email:${Buffer.from(mailbox).toString("base64url")}`;
}
```

### Why conversationId (not subject or internetMessageId)

- **Subject-based threading is unreliable:** Users change subjects, multiple threads can share subjects
- **Graph's `conversationId`** is the canonical grouping key that Outlook uses internally
- **It survives forwards and reply-all** within the same conversation
- **It's indexed and filterable** via `$filter=conversationId eq '...'` for history fetch

---

## 7. Configuration

```typescript
interface OutlookEmailAdapterConfig {
  /** Azure AD tenant ID. Falls back to OUTLOOK_TENANT_ID env var. */
  tenantId?: string;

  /** Azure AD application (client) ID. Falls back to OUTLOOK_CLIENT_ID env var. */
  clientId?: string;

  /** Azure AD client secret. Falls back to OUTLOOK_CLIENT_SECRET env var. */
  clientSecret?: string;

  /**
   * Target mailbox address(es) to monitor.
   * Falls back to OUTLOOK_MAILBOX env var.
   * If string[], creates a subscription for each mailbox.
   * Supports both dedicated user mailboxes and shared mailboxes.
   */
  mailbox?: string | string[];

  /**
   * Public HTTPS URL that receives Graph notifications.
   * Falls back to OUTLOOK_NOTIFICATION_URL env var.
   * Must be reachable from Microsoft's servers.
   */
  notificationUrl?: string;

  /**
   * Secret string included in subscription and verified on each notification.
   * Falls back to OUTLOOK_CLIENT_STATE env var.
   * Max 128 characters.
   */
  clientState?: string;

  /**
   * Bot display name used in the "From" header and Chat SDK userName.
   * Falls back to OUTLOOK_BOT_NAME or BOT_USERNAME env var.
   * Defaults to "Bot".
   */
  botName?: string;

  /**
   * Folder to monitor. Defaults to "inbox".
   * Can be a well-known folder name or folder ID.
   */
  folder?: string;

  /**
   * Reply strategy for postMessage on existing threads.
   * - "createReply": Create a draft reply, optionally update it, then send (default)
   * - "sendMail": Use sendMail with In-Reply-To/References headers
   */
  replyStrategy?: "createReply" | "sendMail";

  /**
   * Whether to save outbound messages to the Sent Items folder.
   * Defaults to true.
   */
  saveToSentItems?: boolean;

  /**
   * Custom subscription renewal interval in minutes.
   * Defaults to 4000 (just under the 4320-minute / 3-day safe window).
   * Max is 10,080 minutes (~7 days) for basic notifications.
   */
  renewalIntervalMinutes?: number;

  /** Logger instance. Defaults to console-based logger. */
  logger?: Logger;
}
```

### Environment Variables

| Variable | Description |
|---|---|
| `OUTLOOK_TENANT_ID` | Azure AD tenant ID |
| `OUTLOOK_CLIENT_ID` | Azure AD application (client) ID |
| `OUTLOOK_CLIENT_SECRET` | Azure AD client secret |
| `OUTLOOK_MAILBOX` | Mailbox address — dedicated or shared (comma-separated for multiple) |
| `OUTLOOK_NOTIFICATION_URL` | Public HTTPS webhook URL |
| `OUTLOOK_CLIENT_STATE` | Webhook verification secret |
| `OUTLOOK_BOT_NAME` | Bot display name |

---

## 8. Inbound Path

### Step 1: Subscription Creation (on `initialize()`)

When the adapter initializes:

1. Authenticate with MSAL using client credentials
2. For each configured mailbox, create a Graph subscription:
   - `resource`: `users/{mailbox}/messages`
   - `changeType`: `created`
   - `notificationUrl`: configured webhook URL
   - `expirationDateTime`: now + `renewalIntervalMinutes`
   - `clientState`: configured secret
3. Store subscription IDs and expiration times
4. Schedule renewal timer

### Step 2: Webhook Validation Challenge

When Graph creates a subscription, it validates the endpoint:

1. Graph sends `POST {notificationUrl}?validationToken={url-encoded-token}`
2. Adapter detects the `validationToken` query parameter
3. Returns `200 OK` with `Content-Type: text/plain` and the URL-decoded token as the body
4. Must respond within **10 seconds** or subscription creation fails

### Step 3: Notification Receipt

When a new email arrives in the monitored mailbox:

1. Graph sends `POST {notificationUrl}` with JSON body containing notification(s)
2. Adapter validates `clientState` on each notification
3. Returns `202 Accepted` immediately (processing happens via `waitUntil`)
4. For each notification in the batch:
   a. Extract the message resource path to get the message ID
   b. Check deduplication (have we seen this notification/message before?)
   c. Fetch the full message via Graph API
   d. Check for self-echo (is this a message the bot sent?)
   e. Parse into a Chat SDK `Message`
   f. Call `this.chat.processMessage(this, threadId, message, options)`

### Step 4: Message Hydration

Fetch the full message from Graph:

```
GET /users/{mailbox}/messages/{messageId}
?$select=id,conversationId,subject,body,from,toRecipients,ccRecipients,
         receivedDateTime,sentDateTime,internetMessageId,
         internetMessageHeaders,hasAttachments,isRead
```

### Step 5: Body Normalization

Email bodies are messy. The normalization pipeline:

1. **Prefer plain text** if `body.contentType === "text"` (use directly)
2. **HTML fallback**: Strip HTML tags to extract text
3. **Strip quoted content**: Remove `>` prefixed lines, `On {date}, {person} wrote:` blocks, and common quote markers
4. **Strip signatures**: Remove content after common signature delimiters (`--`, `___`, `Sent from my iPhone`, etc.)
5. **Trim whitespace**: Collapse excessive newlines
6. Store original HTML in `raw` for debugging/advanced use

### Step 6: Self-Echo Detection

To prevent the bot from replying to its own messages:

1. Check if `from.emailAddress.address` matches any configured bot mailbox
2. Check for a custom `X-Chat-SDK-Bot` internet message header (set on outbound)
3. If either matches, skip processing

### Step 7: Deduplication

Graph notifications can be duplicated (retries, batching, folder moves). Deduplicate by:

1. Using Chat SDK's built-in deduplication (`setIfNotExists` on `dedupe:{adapter}:{messageId}`)
2. Additionally tracking `internetMessageId` to catch edge cases where Graph assigns different internal IDs

---

## 9. Outbound Path

### Replying to an Existing Thread (`postMessage`)

Default strategy: `createReply`

1. Decode the thread ID to get `mailbox` and `conversationId`
2. Find the latest message in the conversation to reply to:
   ```
   GET /users/{mailbox}/messages?$filter=conversationId eq '{conversationId}'&$orderby=receivedDateTime desc&$top=1
   ```
3. Create a reply draft:
   ```
   POST /users/{mailbox}/messages/{latestMessageId}/createReply
   ```
4. Update the draft body with rendered content:
   ```
   PATCH /users/{mailbox}/messages/{draftId}
   Body: { body: { contentType: "html", content: renderedHtml } }
   ```
5. Add `X-Chat-SDK-Bot: true` header for self-echo detection
6. Attach any files from `extractFiles(message)`
7. Send the draft:
   ```
   POST /users/{mailbox}/messages/{draftId}/send
   ```

### Starting a New Thread (`openDM` + `postMessage`)

1. `openDM(email)` creates a new thread ID with a generated conversation placeholder
2. `postMessage` detects the placeholder and uses `sendMail`:
   ```
   POST /users/{mailbox}/sendMail
   Body: {
     message: {
       subject: "...",
       body: { contentType: "html", content: renderedHtml },
       toRecipients: [{ emailAddress: { address: recipientEmail } }],
       internetMessageHeaders: [
         { name: "X-Chat-SDK-Bot", value: "true" }
       ]
     },
     saveToSentItems: true
   }
   ```
3. Returns `202 Accepted` (no message ID returned). The sent message will eventually appear via subscription notification, at which point the real `conversationId` can be resolved.

---

## 10. Format Conversion

### Inbound: Email -> Chat SDK Message (mdast)

```typescript
class OutlookEmailFormatConverter {
  // Parse email text/HTML into mdast AST
  toAst(emailText: string): Root {
    // Split on double newlines into paragraphs
    // Return mdast Root with paragraph nodes
  }

  // Render mdast AST to email-safe HTML
  fromAst(ast: Root): string {
    // Convert mdast -> hast -> HTML via hast-util-to-html
    const hast = toHast(ast);
    return toHtml(hast);
  }
}
```

### Outbound: Chat SDK Message -> Email HTML

The message renderer handles all `AdapterPostableMessage` variants:

1. **String messages**: Escape HTML, wrap in `<p>` tags
2. **Markdown messages**: Parse to mdast, convert to hast, render to HTML
3. **AST messages**: Convert mdast to hast, render to HTML
4. **Card messages**: Extract fallback text, render as simple HTML
5. **Raw messages**: Escape and wrap in `<p>` tags

### HTML Sanitization for Email

Email HTML has different constraints than web HTML. Keep it conservative:

- Support: `<p>`, `<br>`, `<strong>`, `<em>`, `<a>`, `<ul>`, `<ol>`, `<li>`, `<code>`, `<pre>`, `<blockquote>`, `<h1>`-`<h6>`, `<img>`, `<table>`, `<tr>`, `<td>`, `<th>`
- No `<script>`, `<style>`, `<iframe>`, `<form>`
- Use inline styles only (many email clients strip `<style>` blocks)
- Prefer simple layout over complex CSS

---

## 11. Subscription Lifecycle Management

### Creation

On `initialize()`, for each configured mailbox:

1. Check if there's an existing subscription (via `GET /subscriptions`)
2. If not, create a new subscription
3. Store the subscription ID and expiration

### Renewal

Subscriptions must be renewed before they expire:

1. Schedule a renewal timer at half the configured `renewalIntervalMinutes` for safety margin
2. On renewal:
   ```
   PATCH /subscriptions/{id}
   Body: { expirationDateTime: <now + renewalIntervalMinutes> }
   ```
3. Update stored expiration time
4. If renewal fails (subscription gone), recreate from scratch

### Reconciliation

On startup (in `initialize()`):

1. List all active subscriptions via `GET /subscriptions`
2. For each configured mailbox, check if a valid subscription exists
3. If subscription exists but is near expiry, renew it
4. If subscription is missing, create a new one
5. Clean up any orphaned subscriptions from this application

### Cleanup

On `disconnect()`:

1. Cancel renewal timers
2. Delete all managed subscriptions

### Lifecycle Notifications (future)

Graph can send lifecycle notifications when subscriptions are about to expire or when reauthorization is needed. These arrive at a separate `lifecycleNotificationUrl`. Not implemented yet but should be considered for production hardening.

---

## 12. Deduplication and Loop Prevention

### Deduplication Strategy

**Layer 1: Chat SDK built-in dedup**
The Chat SDK's `handleIncomingMessage` already deduplicates by `{adapter}:{messageId}` using the state adapter's `setIfNotExists`. This handles most cases.

**Layer 2: internetMessageId tracking**
As a secondary check, track `internetMessageId` (RFC 822 Message-ID) to catch cases where Graph assigns different internal IDs to the same email (e.g., folder moves).

### Loop Prevention

Email loops are a critical failure mode. A bot that replies to itself will generate infinite emails.

**Strategy 1: Self-echo detection (primary)**
- Check `from.emailAddress.address` against configured bot mailbox(es)
- Chat SDK also checks `author.isMe` in `handleIncomingMessage`

**Strategy 2: Custom header tagging (defense in depth)**
- Set `X-Chat-SDK-Bot: true` on all outbound messages via `internetMessageHeaders`
- Check for this header on inbound messages
- If present, skip processing

**Strategy 3: Sent message fingerprinting (fallback)**
- Hash `{to}:{subject}:{bodyPrefix}:{timestamp}` for each outbound message
- Store hashes with a TTL
- Check inbound messages against stored hashes
- Catches cases where headers are stripped by intermediate mail servers

### Edge Cases

- **Distribution lists / group emails**: Bot receives its own message via a DL -> header check catches this
- **Auto-replies (OOF)**: Check for `X-Auto-Reply-From` or `Auto-Submitted` headers, skip if present
- **Read receipts**: Filter out non-message notifications (`@odata.type` check)
- **Calendar invites / meeting updates**: Filter by message class if needed

---

## 13. File Structure

```
chat-adapter-outlook-email/
├── src/
│   ├── index.ts                    # Factory function + re-exports
│   ├── adapter.ts                  # OutlookEmailAdapter class
│   ├── types.ts                    # All type definitions
│   ├── thread-id.ts                # Thread ID encode/decode
│   ├── format-converter.ts         # Email <-> mdast conversion
│   ├── webhook-handler.ts          # Graph webhook validation + parsing
│   ├── message-parser.ts           # Graph message -> Chat SDK Message
│   ├── message-renderer.ts         # Chat SDK message -> email HTML
│   ├── graph-client.ts             # Microsoft Graph API wrapper
│   ├── subscription-manager.ts     # Subscription lifecycle management
│   ├── body-normalizer.ts          # Email body cleanup (quotes, signatures)
│   └── errors.ts                   # Error mapping (Graph -> Chat SDK errors)
├── tests/
│   ├── adapter.test.ts
│   ├── thread-id.test.ts
│   ├── format-converter.test.ts
│   ├── webhook-handler.test.ts
│   ├── message-parser.test.ts
│   ├── body-normalizer.test.ts
│   ├── message-renderer.test.ts
│   ├── subscription-manager.test.ts
│   └── fixtures/
│       ├── graph-message.json       # Sample Graph API message response
│       ├── notification.json        # Sample webhook notification payload
│       └── email-bodies/            # Various HTML email body samples
│           ├── plain-text.txt
│           ├── html-simple.html
│           ├── html-with-quotes.html
│           ├── html-with-signature.html
│           └── html-outlook-reply.html
├── package.json
├── tsconfig.json
├── tsup.config.ts
├── vitest.config.ts
├── SPEC.md                          # This file
└── .env.example
```

---

## 14. Package Configuration

### Key Dependency Decisions

| Dependency | Why |
|---|---|
| `@azure/identity` | MSAL client credentials auth. The official Microsoft SDK for Azure AD token acquisition. No need for the full `@microsoft/microsoft-graph-client` — direct `fetch()` calls give better control and a smaller bundle. |
| `@react-email/components` + `@react-email/render` | Available for card rendering to email HTML. |
| `mdast-util-to-hast` + `hast-util-to-html` | mdast AST -> HTML conversion for email bodies. |
| **NOT** `@microsoft/microsoft-graph-client` | Too heavy, abstracts away details needed for subscription management and error handling. |
| **NOT** `@microsoft/teams.*` | That's the Bot Framework SDK. This adapter uses Graph API directly. |

---

## 15. Testing Strategy

### Unit Tests (Vitest)

Mock Graph API responses. Test:

- Thread ID encode/decode roundtripping
- Webhook validation token challenge
- Webhook clientState verification
- Notification payload parsing
- Message hydration and normalization
- Email body cleanup (quoted text, signatures, HTML stripping)
- Format conversion (text -> mdast -> HTML)
- Self-echo detection
- Deduplication logic
- Subscription creation/renewal logic
- Error mapping (Graph HTTP errors -> Chat SDK errors)
- Reply draft creation flow
- Attachment mapping

### Integration Tests

Requires a test M365 tenant with:

- One bot mailbox (application permissions configured)
- One human test mailbox
- Azure AD app registration with `Mail.ReadWrite` + `Mail.Send`
- A publicly reachable webhook URL (ngrok or similar for local dev)

Test scenarios:

1. **Full round-trip:** Send email to bot -> bot receives via webhook -> bot replies -> reply appears in sender's inbox as threaded reply
2. **New thread:** Bot initiates an outbound email -> recipient replies -> bot receives the reply in the correct thread
3. **Attachment round-trip:** Send email with attachment -> bot reads attachment -> bot replies with a file
4. **Subscription lifecycle:** Create subscription -> wait for near-expiry -> verify renewal -> restart adapter -> verify reconciliation
5. **Deduplication:** Simulate duplicate notifications -> verify single processing

### Fixture-Based Tests

Realistic fixtures from actual Graph API responses:

- `fixtures/graph-message.json` - Full message object with all fields
- `fixtures/notification.json` - Webhook notification batch
- `fixtures/email-bodies/` - Various real-world email body formats:
  - Plain text only
  - Simple HTML
  - HTML with Outlook-style quoted reply chains
  - HTML with corporate signatures/disclaimers
  - HTML with inline images
  - Gmail-style quoted text
  - Apple Mail format

---

## 16. Known Risks and Mitigations

### Risk 1: Email Body Normalization Quality

**Risk:** Quoted text stripping and signature removal will never be perfect. Different email clients format replies differently, and corporate disclaimers vary wildly.

**Mitigation:**
- Start conservative (strip only obvious patterns)
- Keep original HTML in `raw` for handlers that need full fidelity
- Use heuristic-based stripping with known patterns from major clients (Outlook, Gmail, Apple Mail)
- Make normalization configurable (handlers can opt into raw bodies)

### Risk 2: Conversation Identity Edge Cases

**Risk:** Graph's `conversationId` is usually reliable but can break with forwarded messages, changed subjects (some clients), or cross-tenant scenarios.

**Mitigation:**
- Use `conversationId` as primary key (it's the best option)
- Fall back to `internetMessageId` / `In-Reply-To` header chain when `conversationId` is missing
- Log warnings for edge cases to build understanding over time

### Risk 3: Webhook Reliability / Graph Throttling

**Risk:** Slow webhook processing leads to dropped notifications. Missed notifications mean missed emails.

**Mitigation:**
- Return `202 Accepted` immediately, always (use `waitUntil` for async processing)
- Keep webhook handler minimal (parse + enqueue)
- Monitor for slow/drop state indicators
- Consider periodic polling as a fallback for missed notifications (future)

### Risk 4: Subscription Expiration During Downtime

**Risk:** If the adapter is offline for >7 days, subscriptions expire and inbound stops.

**Mitigation:**
- Reconciliation on startup (detect and recreate expired subscriptions)
- Health check endpoint that monitors subscription state
- Alert when subscription renewal fails

### Risk 5: Reply Loops

**Risk:** Bot replies to itself, causing infinite email chains.

**Mitigation:** Three-layer defense (see Section 12):
1. `from` address check against bot mailbox(es)
2. `X-Chat-SDK-Bot` custom header
3. Sent message fingerprint matching

### Risk 6: Rate Limiting

**Risk:** Graph API has rate limits that vary by tenant and endpoint.

**Mitigation:**
- Respect `Retry-After` headers
- Implement exponential backoff in `graph-client.ts`
- Batch notification processing to minimize API calls
- Use `$select` to minimize response sizes

### Risk 7: Application Permission Scope

**Risk:** Application permissions are tenant-wide by default, granting access to all mailboxes.

**Mitigation:**
- Use RBAC for Applications in Exchange Online to scope access to only the bot/shared mailbox (see Section 5)
- Document the scoping requirement clearly for operators
- Consider supporting `ClientCertificateCredential` as an alternative to client secrets for higher-security deployments

---

## 17. Reference Implementations

The following existing adapters informed the design of this adapter:

### Resend Email Adapter (`@resend/chat-sdk-adapter`)
- **Repo:** https://github.com/resend/resend-chat-sdk
- **Patterns adopted:** Thread resolver, webhook handler separation, message parser/renderer split, format converter using mdast->hast->HTML, `openDM` for new threads, `NotImplementedError` for unsupported methods

### Teams Adapter (`@chat-adapter/teams`)
- **Repo:** https://github.com/vercel/chat (monorepo `packages/adapter-teams/`)
- **Patterns adopted:** Thread ID base64url encoding, Graph API client module, error mapping from HTTP status codes to shared error classes, config with env var fallbacks

### Chat SDK Core (`chat`)
- **Repo:** https://github.com/vercel/chat (monorepo `packages/chat/`)
- **Key reference:** `Adapter` interface, `Message` class, `BaseFormatConverter`, `processMessage` flow with built-in deduplication, `WebhookOptions` with `waitUntil`

---

## 18. Design Decisions

Decisions made during implementation:

1. **Subscription storage:** Subscriptions are tracked in-memory by the `SubscriptionManager`. On startup, existing subscriptions are reconciled via `GET /subscriptions`. This avoids an external store dependency; the tradeoff is that subscriptions are recreated on restart if not found.

2. **Multi-mailbox routing:** One adapter instance manages all configured mailboxes, with one subscription per mailbox and internal routing by mailbox address extracted from the notification resource path.

3. **New thread subject line:** Currently defaults to `"New message"` for outbound threads initiated via `openDM`. A future enhancement could derive the subject from message content or accept it as a parameter.

4. **Fetch-after-notify (not rich notifications):** Inbound uses basic notifications (no resource data) followed by a `GET /messages/{id}` call to hydrate the full message. Rich notifications with encrypted resource data could reduce API calls but add certificate management complexity.

5. **HTML email template:** Outbound emails use minimal HTML (just the rendered content). No branding template or footer is applied. A future enhancement could add an optional template hook.

6. **Card rendering:** Cards are rendered as plain text fallback rather than React Email HTML. The `@react-email/components` dependency is available for a future rich card rendering implementation.
