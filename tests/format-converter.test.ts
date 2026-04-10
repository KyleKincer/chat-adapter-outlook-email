import { describe, expect, it } from "vitest";
import { OutlookEmailFormatConverter } from "../src/format-converter.js";
import type { Root } from "mdast";

const converter = new OutlookEmailFormatConverter();

describe("OutlookEmailFormatConverter", () => {
	describe("fromAst", () => {
		it("converts a simple paragraph to HTML", () => {
			const ast: Root = {
				type: "root",
				children: [
					{
						type: "paragraph",
						children: [{ type: "text", value: "Hello world" }],
					},
				],
			};
			const html = converter.fromAst(ast);
			expect(html).toContain("Hello world");
			expect(html).toContain("<p>");
		});

		it("converts multiple paragraphs", () => {
			const ast: Root = {
				type: "root",
				children: [
					{
						type: "paragraph",
						children: [{ type: "text", value: "First" }],
					},
					{
						type: "paragraph",
						children: [{ type: "text", value: "Second" }],
					},
				],
			};
			const html = converter.fromAst(ast);
			expect(html).toContain("First");
			expect(html).toContain("Second");
		});

		it("returns empty string for empty AST", () => {
			const ast: Root = { type: "root", children: [] };
			const html = converter.fromAst(ast);
			expect(html).toBe("");
		});

		it("converts strong emphasis", () => {
			const ast: Root = {
				type: "root",
				children: [
					{
						type: "paragraph",
						children: [
							{
								type: "strong",
								children: [
									{ type: "text", value: "bold text" },
								],
							},
						],
					},
				],
			};
			const html = converter.fromAst(ast);
			expect(html).toContain("<strong>");
			expect(html).toContain("bold text");
		});

		it("converts emphasis", () => {
			const ast: Root = {
				type: "root",
				children: [
					{
						type: "paragraph",
						children: [
							{
								type: "emphasis",
								children: [
									{ type: "text", value: "italic text" },
								],
							},
						],
					},
				],
			};
			const html = converter.fromAst(ast);
			expect(html).toContain("<em>");
			expect(html).toContain("italic text");
		});

		it("converts a list", () => {
			const ast: Root = {
				type: "root",
				children: [
					{
						type: "list",
						ordered: false,
						spread: false,
						children: [
							{
								type: "listItem",
								spread: false,
								children: [
									{
										type: "paragraph",
										children: [
											{
												type: "text",
												value: "Item 1",
											},
										],
									},
								],
							},
							{
								type: "listItem",
								spread: false,
								children: [
									{
										type: "paragraph",
										children: [
											{
												type: "text",
												value: "Item 2",
											},
										],
									},
								],
							},
						],
					},
				],
			};
			const html = converter.fromAst(ast);
			expect(html).toContain("<ul>");
			expect(html).toContain("<li>");
			expect(html).toContain("Item 1");
			expect(html).toContain("Item 2");
		});
	});

	describe("toAst", () => {
		it("converts plain text to mdast AST", () => {
			const ast = converter.toAst("Hello world");
			expect(ast.type).toBe("root");
			expect(ast.children).toHaveLength(1);
			expect(ast.children[0]!.type).toBe("paragraph");
		});

		it("splits double newlines into paragraphs", () => {
			const ast = converter.toAst("First paragraph\n\nSecond paragraph");
			expect(ast.children).toHaveLength(2);
		});

		it("returns empty root for empty string", () => {
			const ast = converter.toAst("");
			expect(ast.type).toBe("root");
			expect(ast.children).toHaveLength(0);
		});

		it("returns empty root for whitespace-only string", () => {
			const ast = converter.toAst("   \n\n   ");
			expect(ast.type).toBe("root");
			expect(ast.children).toHaveLength(0);
		});

		it("trims paragraph content", () => {
			const ast = converter.toAst("  Hello  \n\n  World  ");
			expect(ast.children).toHaveLength(2);
			const first = ast.children[0] as {
				children: Array<{ value: string }>;
			};
			expect(first.children[0]!.value).toBe("Hello");
		});
	});

	describe("extractPlainText", () => {
		it("returns text as-is", () => {
			expect(converter.extractPlainText("Hello world")).toBe(
				"Hello world",
			);
		});
	});
});
