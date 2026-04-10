import { describe, expect, it } from "vitest";
import {
  channelIdFromThreadId,
  decodeThreadId,
  encodeThreadId,
} from "../src/thread-id.js";

describe("thread-id", () => {
  const sampleData = {
    mailbox: "bot@contoso.com",
    conversationId: "AAQkADAwMDAwMDY3LTk...",
  };

  describe("encodeThreadId", () => {
    it("encodes with outlook-email prefix", () => {
      const encoded = encodeThreadId(sampleData);
      expect(encoded).toMatch(/^outlook-email:/);
    });

    it("produces three colon-separated parts", () => {
      const encoded = encodeThreadId(sampleData);
      const parts = encoded.split(":");
      expect(parts).toHaveLength(3);
      expect(parts[0]).toBe("outlook-email");
    });
  });

  describe("decodeThreadId", () => {
    it("roundtrips encode/decode", () => {
      const encoded = encodeThreadId(sampleData);
      const decoded = decodeThreadId(encoded);
      expect(decoded).toEqual(sampleData);
    });

    it("throws on invalid prefix", () => {
      expect(() => decodeThreadId("slack:abc:def")).toThrow(
        "Invalid Outlook email thread ID"
      );
    });

    it("throws on wrong number of parts", () => {
      expect(() => decodeThreadId("outlook-email:abc")).toThrow(
        "Invalid Outlook email thread ID"
      );
    });
  });

  describe("channelIdFromThreadId", () => {
    it("returns channel based on mailbox only", () => {
      const threadA = encodeThreadId({
        mailbox: "bot@contoso.com",
        conversationId: "conv-1",
      });
      const threadB = encodeThreadId({
        mailbox: "bot@contoso.com",
        conversationId: "conv-2",
      });
      expect(channelIdFromThreadId(threadA)).toBe(
        channelIdFromThreadId(threadB)
      );
    });

    it("differs for different mailboxes", () => {
      const threadA = encodeThreadId({
        mailbox: "bot@contoso.com",
        conversationId: "conv-1",
      });
      const threadB = encodeThreadId({
        mailbox: "other@contoso.com",
        conversationId: "conv-1",
      });
      expect(channelIdFromThreadId(threadA)).not.toBe(
        channelIdFromThreadId(threadB)
      );
    });
  });
});
