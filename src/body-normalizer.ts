/**
 * Email body normalization utilities.
 * Strips quoted reply chains, signatures, and cleans up HTML
 * to extract the meaningful new content from an email.
 */

// "On {date}, {person} wrote:" patterns (Gmail, Apple Mail)
const ON_WROTE_RE = /^\s*On .+wrote:\s*$/m;

// Outlook-style separator: "--- Original Message ---" or "-----Original Message-----"
const OUTLOOK_SEPARATOR_RE = /^-{2,}\s*Original Message\s*-{2,}/im;

// Outlook "From:" block that starts a quoted section
const OUTLOOK_FROM_BLOCK_RE =
	/^From:\s+.+\nSent:\s+.+\nTo:\s+.+\n(?:Cc:\s+.+\n)?Subject:\s+.+$/im;

// ">" prefix quoted lines (Gmail, plain text clients)
const QUOTED_LINE_RE = /^>+\s?.*$/gm;

// Common signature delimiters
const SIGNATURE_PATTERNS: RegExp[] = [
	/^--\s*$/m, // Standard "-- " delimiter
	/^_{3,}\s*$/m, // Underscores
	/^Sent from my iPhone/im,
	/^Sent from my iPad/im,
	/^Sent from Mail for Windows/im,
	/^Sent from Outlook/im,
	/^Get Outlook for /im,
	/^Sent from my Galaxy/im,
	/^Sent from Samsung/im,
	/^Sent from Yahoo Mail/im,
];

/**
 * Normalize an email body by stripping quoted replies and signatures.
 * Works on plain text content (call stripHtml first for HTML emails).
 */
export function normalizeEmailBody(text: string): string {
	let result = text;

	result = stripQuotedContent(result);
	result = stripSignature(result);

	// Collapse excessive whitespace
	result = result.replace(/\n{3,}/g, "\n\n").trim();

	return result;
}

/**
 * Strip quoted reply content from various email clients.
 */
function stripQuotedContent(text: string): string {
	let result = text;

	// Remove Outlook "--- Original Message ---" and everything after
	const outlookMatch = result.match(OUTLOOK_SEPARATOR_RE);
	if (outlookMatch?.index !== undefined) {
		result = result.substring(0, outlookMatch.index).trim();
	}

	// Remove Outlook From:/Sent:/To:/Subject: block and everything after
	const fromBlockMatch = result.match(OUTLOOK_FROM_BLOCK_RE);
	if (fromBlockMatch?.index !== undefined) {
		result = result.substring(0, fromBlockMatch.index).trim();
	}

	// Remove "On {date}, {person} wrote:" and everything after
	const onWroteMatch = result.match(ON_WROTE_RE);
	if (onWroteMatch?.index !== undefined) {
		result = result.substring(0, onWroteMatch.index).trim();
	}

	// Remove > quoted lines
	result = result.replace(QUOTED_LINE_RE, "").trim();

	return result;
}

/**
 * Strip email signatures by finding the earliest signature delimiter.
 */
function stripSignature(text: string): string {
	let earliestIndex = text.length;

	for (const pattern of SIGNATURE_PATTERNS) {
		const match = text.match(pattern);
		if (match?.index !== undefined && match.index < earliestIndex) {
			earliestIndex = match.index;
		}
	}

	if (earliestIndex < text.length) {
		return text.substring(0, earliestIndex).trim();
	}

	return text;
}

/**
 * Strip HTML tags and decode entities to extract plain text.
 * Converts block-level elements to appropriate whitespace.
 */
export function stripHtml(html: string): string {
	return (
		html
			// Convert block elements to newlines before stripping tags
			.replace(/<br\s*\/?>/gi, "\n")
			.replace(/<\/p>/gi, "\n\n")
			.replace(/<\/div>/gi, "\n")
			.replace(/<\/li>/gi, "\n")
			.replace(/<\/tr>/gi, "\n")
			.replace(/<\/h[1-6]>/gi, "\n\n")
			// Strip all remaining tags
			.replace(/<[^>]+>/g, "")
			// Decode common HTML entities
			.replace(/&nbsp;/gi, " ")
			.replace(/&amp;/gi, "&")
			.replace(/&lt;/gi, "<")
			.replace(/&gt;/gi, ">")
			.replace(/&quot;/gi, '"')
			.replace(/&#39;/gi, "'")
			.replace(/&#x27;/gi, "'")
			.replace(/&hellip;/gi, "\u2026")
			.replace(/&mdash;/gi, "\u2014")
			.replace(/&ndash;/gi, "\u2013")
			// Clean up whitespace
			.replace(/\n{3,}/g, "\n\n")
			.trim()
	);
}
