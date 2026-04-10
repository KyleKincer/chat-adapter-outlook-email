import { describe, expect, it } from "vitest";
import { normalizeEmailBody, stripHtml } from "../src/body-normalizer.js";

describe("stripHtml", () => {
	it("strips simple HTML tags", () => {
		expect(stripHtml("<p>Hello</p>")).toBe("Hello");
	});

	it("converts <br> to newlines", () => {
		expect(stripHtml("Hello<br>World")).toBe("Hello\nWorld");
		expect(stripHtml("Hello<br/>World")).toBe("Hello\nWorld");
		expect(stripHtml("Hello<br />World")).toBe("Hello\nWorld");
	});

	it("converts </p> to double newlines", () => {
		expect(stripHtml("<p>First</p><p>Second</p>")).toBe("First\n\nSecond");
	});

	it("converts </div> to newlines", () => {
		expect(stripHtml("<div>First</div><div>Second</div>")).toBe(
			"First\nSecond",
		);
	});

	it("decodes common HTML entities", () => {
		expect(stripHtml("&amp; &lt; &gt; &quot; &#39;")).toBe(
			'& < > " \'',
		);
	});

	it("decodes &nbsp; to space", () => {
		expect(stripHtml("Hello&nbsp;World")).toBe("Hello World");
	});

	it("collapses excessive newlines", () => {
		expect(stripHtml("<p>A</p><p></p><p></p><p>B</p>")).toBe("A\n\nB");
	});

	it("handles a realistic HTML email body", () => {
		const html =
			"<html><body><p>Hello,</p><p>This is a test.</p></body></html>";
		const result = stripHtml(html);
		expect(result).toContain("Hello,");
		expect(result).toContain("This is a test.");
	});

	it("trims whitespace", () => {
		expect(stripHtml("  <p>Hello</p>  ")).toBe("Hello");
	});
});

describe("normalizeEmailBody", () => {
	it("returns simple text unchanged", () => {
		expect(normalizeEmailBody("Hello, how are you?")).toBe(
			"Hello, how are you?",
		);
	});

	it("strips Gmail-style 'On ... wrote:' quoted content", () => {
		const body = `Sure, I'll look into it.

On Thu, Apr 10, 2026 at 12:00 PM Alice Smith <alice@contoso.com> wrote:
> Can you check the logs?
> Thanks.`;
		const result = normalizeEmailBody(body);
		expect(result).toBe("Sure, I'll look into it.");
	});

	it("strips > quoted lines", () => {
		const body = `Got it, thanks!

> Previous message content
> More quoted content`;
		const result = normalizeEmailBody(body);
		expect(result).toBe("Got it, thanks!");
	});

	it("strips Outlook Original Message separator", () => {
		const body = `I agree with your proposal.

-----Original Message-----
From: Alice Smith
Sent: Thursday, April 10, 2026 12:00 PM
To: Bob
Subject: Proposal

Please review the attached proposal.`;
		const result = normalizeEmailBody(body);
		expect(result).toBe("I agree with your proposal.");
	});

	it("strips Outlook From/Sent/To/Subject blocks", () => {
		const body = `Sounds good.

From: Alice Smith <alice@contoso.com>
Sent: Thursday, April 10, 2026 12:00 PM
To: Bot <bot@contoso.com>
Subject: Follow up

Original message here.`;
		const result = normalizeEmailBody(body);
		expect(result).toBe("Sounds good.");
	});

	it("strips standard signature delimiter", () => {
		const body = `Please see the attached report.

--
Alice Smith
Senior Engineer
Contoso Inc.`;
		const result = normalizeEmailBody(body);
		expect(result).toBe("Please see the attached report.");
	});

	it("strips 'Sent from my iPhone' signature", () => {
		const body = `Ok sounds good

Sent from my iPhone`;
		const result = normalizeEmailBody(body);
		expect(result).toBe("Ok sounds good");
	});

	it("strips 'Sent from Outlook' signature", () => {
		const body = `Meeting confirmed for 3pm.

Sent from Outlook for iOS`;
		const result = normalizeEmailBody(body);
		expect(result).toBe("Meeting confirmed for 3pm.");
	});

	it("handles underscore signature delimiter", () => {
		const body = `Check the document.

___
John Doe
VP Engineering`;
		const result = normalizeEmailBody(body);
		expect(result).toBe("Check the document.");
	});

	it("collapses excessive newlines", () => {
		const body = "First line\n\n\n\n\nSecond line";
		const result = normalizeEmailBody(body);
		expect(result).toBe("First line\n\nSecond line");
	});

	it("handles empty string", () => {
		expect(normalizeEmailBody("")).toBe("");
	});

	it("handles text with no quotes or signatures", () => {
		const body = "This is a clean email with no quotes or signatures.";
		expect(normalizeEmailBody(body)).toBe(body);
	});
});
