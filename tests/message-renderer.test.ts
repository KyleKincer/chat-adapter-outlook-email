import { describe, expect, it } from "vitest";
import { renderMessage } from "../src/message-renderer.js";
import type { Root } from "mdast";

describe("renderMessage", () => {
	describe("string messages", () => {
		it("renders a plain string as HTML", () => {
			const result = renderMessage("Hello world");
			expect(result.html).toContain("<p>Hello world</p>");
			expect(result.text).toBe("Hello world");
			expect(result.files).toEqual([]);
		});

		it("escapes HTML in string messages", () => {
			const result = renderMessage("<script>alert('xss')</script>");
			expect(result.html).not.toContain("<script>");
			expect(result.html).toContain("&lt;script&gt;");
		});

		it("converts newlines to paragraphs and breaks", () => {
			const result = renderMessage("First\n\nSecond");
			expect(result.html).toContain("<p>First</p>");
			expect(result.html).toContain("<p>Second</p>");
		});

		it("converts single newlines to <br>", () => {
			const result = renderMessage("Line 1\nLine 2");
			expect(result.html).toContain("Line 1<br>Line 2");
		});
	});

	describe("raw messages", () => {
		it("renders PostableRaw content", () => {
			const result = renderMessage({ raw: "Raw content here" });
			expect(result.html).toContain("Raw content here");
			expect(result.text).toBe("Raw content here");
		});

		it("extracts files from PostableRaw", () => {
			const fileData = Buffer.from("file contents");
			const result = renderMessage({
				raw: "Message with file",
				files: [
					{
						data: fileData,
						filename: "test.txt",
						mimeType: "text/plain",
					},
				],
			});
			expect(result.files).toHaveLength(1);
			expect(result.files[0]!.filename).toBe("test.txt");
		});
	});

	describe("markdown messages", () => {
		it("renders PostableMarkdown to HTML", () => {
			const result = renderMessage({ markdown: "**bold** text" });
			expect(result.html).toContain("<strong>");
			expect(result.html).toContain("bold");
			expect(result.text).toBe("**bold** text");
		});

		it("renders markdown lists", () => {
			const result = renderMessage({
				markdown: "- Item 1\n- Item 2",
			});
			expect(result.html).toContain("<ul>");
			expect(result.html).toContain("<li>");
		});
	});

	describe("AST messages", () => {
		it("renders PostableAst to HTML", () => {
			const ast: Root = {
				type: "root",
				children: [
					{
						type: "paragraph",
						children: [{ type: "text", value: "AST content" }],
					},
				],
			};
			const result = renderMessage({ ast });
			expect(result.html).toContain("AST content");
			expect(result.text).toBe("AST content");
		});

		it("extracts text from nested AST nodes", () => {
			const ast: Root = {
				type: "root",
				children: [
					{
						type: "paragraph",
						children: [
							{ type: "text", value: "Hello " },
							{
								type: "strong",
								children: [
									{ type: "text", value: "world" },
								],
							},
						],
					},
				],
			};
			const result = renderMessage({ ast });
			expect(result.text).toBe("Hello world");
		});
	});

	describe("card messages", () => {
		it("renders PostableCard with fallback text", () => {
			const result = renderMessage({
				card: {
					type: "card",
					children: [
						{ type: "text", text: "Card content" },
					],
				},
				fallbackText: "Fallback text",
			});
			expect(result.text).toBe("Fallback text");
			expect(result.html).toContain("Fallback text");
		});

		it("generates fallback from card children when no fallbackText", () => {
			const result = renderMessage({
				card: {
					type: "card",
					children: [
						{ type: "text", text: "Hello from card" },
					],
				},
			});
			expect(result.text).toBe("Hello from card");
		});

		it("handles CardElement directly", () => {
			const result = renderMessage({
				type: "card",
				children: [
					{ type: "text", text: "Direct card" },
				],
			} as unknown as string);
			expect(result.text).toBe("Direct card");
		});
	});

	describe("empty/unknown messages", () => {
		it("handles empty object gracefully", () => {
			const result = renderMessage({} as unknown as string);
			expect(result.html).toBe("");
			expect(result.text).toBe("");
		});
	});
});
