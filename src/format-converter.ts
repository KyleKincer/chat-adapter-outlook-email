import { toHtml } from "hast-util-to-html";
import type { Root, RootContent } from "mdast";
import { toHast } from "mdast-util-to-hast";

const PARAGRAPH_SPLIT_RE = /\n\n+/;

/**
 * Format converter for Outlook email.
 * Converts between email text/HTML and mdast AST.
 *
 * Follows the same pattern as the Resend adapter's format converter,
 * using mdast -> hast -> HTML for outbound rendering.
 */
export class OutlookEmailFormatConverter {
  /**
   * Convert mdast AST to email-safe HTML.
   */
  fromAst(ast: Root): string {
    const hast = toHast(ast);
    if (!hast) {
      return "";
    }
    return toHtml(hast);
  }

  /**
   * Convert plain text into mdast AST.
   * Splits on double newlines into paragraphs.
   */
  toAst(text: string): Root {
    if (!text || text.trim() === "") {
      return { type: "root", children: [] };
    }

    const paragraphs = text.split(PARAGRAPH_SPLIT_RE);
    const children: RootContent[] = paragraphs
      .filter((p) => p.trim() !== "")
      .map((p) => ({
        type: "paragraph" as const,
        children: [{ type: "text" as const, value: p.trim() }],
      }));

    return { type: "root", children };
  }

  /**
   * Extract plain text from email content.
   */
  extractPlainText(text: string): string {
    return text;
  }
}
