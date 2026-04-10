import type { AdapterPostableMessage } from "chat";
import { parseMarkdown } from "chat";
import type { Root } from "mdast";
import { OutlookEmailFormatConverter } from "./format-converter.js";

interface FileUpload {
	data: Buffer | Blob | ArrayBuffer;
	filename: string;
	mimeType?: string;
}

interface RenderedMessage {
	html: string;
	text: string;
	files: FileUpload[];
}

const formatConverter = new OutlookEmailFormatConverter();

/**
 * Render a Chat SDK AdapterPostableMessage into email-ready HTML and plain text.
 * Handles all message variants: string, raw, markdown, AST, and card.
 */
export function renderMessage(message: AdapterPostableMessage): RenderedMessage {
	// Plain string
	if (typeof message === "string") {
		return {
			html: textToHtml(message),
			text: message,
			files: [],
		};
	}

	// CardElement (has 'type' and 'children' properties directly, no wrapper)
	if (isCardElement(message)) {
		const text = cardToFallbackText(message);
		return { html: textToHtml(text), text, files: [] };
	}

	const files = extractFiles(message as unknown as Record<string, unknown>);

	// PostableRaw - raw string content
	if ("raw" in message && typeof message.raw === "string") {
		return {
			html: textToHtml(message.raw),
			text: message.raw,
			files,
		};
	}

	// PostableMarkdown - markdown string
	if ("markdown" in message && typeof message.markdown === "string") {
		const ast = parseMarkdown(message.markdown);
		return {
			html: formatConverter.fromAst(ast),
			text: message.markdown,
			files,
		};
	}

	// PostableAst - mdast AST
	if ("ast" in message && message.ast) {
		const ast = message.ast as Root;
		return {
			html: formatConverter.fromAst(ast),
			text: astToPlainText(ast),
			files,
		};
	}

	// PostableCard - card with optional fallback
	if ("card" in message && message.card) {
		const text =
			("fallbackText" in message && typeof message.fallbackText === "string"
				? message.fallbackText
				: null) || cardToFallbackText(message.card);
		return { html: textToHtml(text), text, files };
	}

	return { html: "", text: "", files: [] };
}

/**
 * Convert plain text to simple email-safe HTML.
 * Escapes HTML entities and wraps paragraphs in <p> tags.
 */
function textToHtml(text: string): string {
	const escaped = text
		.replace(/&/g, "&amp;")
		.replace(/</g, "&lt;")
		.replace(/>/g, "&gt;");

	const paragraphs = escaped.split(/\n\n+/);
	return paragraphs
		.map((p) => `<p>${p.replace(/\n/g, "<br>")}</p>`)
		.join("");
}

function extractFiles(message: Record<string, unknown>): FileUpload[] {
	if ("files" in message && Array.isArray(message.files)) {
		return message.files as FileUpload[];
	}
	return [];
}

/**
 * Extract plain text from an mdast AST by walking text nodes.
 */
function astToPlainText(ast: Root): string {
	const parts: string[] = [];
	for (const node of ast.children) {
		parts.push(nodeToText(node));
	}
	return parts.join("\n\n");
}

interface AstNode {
	type: string;
	value?: string;
	children?: AstNode[];
}

function nodeToText(node: AstNode): string {
	if (node.value !== undefined) {
		return node.value;
	}
	if (node.children) {
		return node.children.map(nodeToText).join("");
	}
	return "";
}

/**
 * Check if a value looks like a CardElement (has type + children, but isn't
 * a wrapped PostableCard).
 */
function isCardElement(
	value: unknown,
): value is { type: string; children: unknown[] } {
	return (
		typeof value === "object" &&
		value !== null &&
		"type" in value &&
		"children" in value &&
		!("card" in value) &&
		!("raw" in value) &&
		!("markdown" in value) &&
		!("ast" in value)
	);
}

/**
 * Convert a CardElement to plain text for email rendering.
 * Recursively extracts text content from card children.
 */
function cardToFallbackText(card: {
	type: string;
	children?: unknown[];
}): string {
	if (!card.children || !Array.isArray(card.children)) {
		return "";
	}
	return card.children
		.map((child) => {
			if (typeof child === "string") return child;
			if (child && typeof child === "object") {
				if ("text" in child && typeof (child as { text: unknown }).text === "string") {
					return (child as { text: string }).text;
				}
				if ("children" in child) {
					return cardToFallbackText(
						child as { type: string; children?: unknown[] },
					);
				}
			}
			return "";
		})
		.filter(Boolean)
		.join("\n");
}
