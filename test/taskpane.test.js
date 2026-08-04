"use strict";

const test = require("node:test");
const assert = require("node:assert/strict");
const {
  checkLookalikeDomaIn,
  extractDomain,
  extractMailboxAddress,
  reconcileAiVerdict,
  runRuleEngine,
} = require("../taskpane.js");

function email(overrides = {}) {
  return {
    subject: "Weekly project update",
    senderName: "Example Sender",
    senderEmail: "sender@example.com",
    replyTo: null,
    bodyText: "Here is the normal weekly update.",
    links: [],
    attachments: [],
    ...overrides,
  };
}

test("extracts a mailbox address and domain from a display-name Reply-To", () => {
  assert.equal(extractMailboxAddress("Support Team <reply@example.com>"), "reply@example.com");
  assert.equal(extractDomain("Support Team <reply@example.com>"), "example.com");
});

test("does not flag matching sender and Reply-To domains", () => {
  const result = runRuleEngine(email({ replyTo: "Support Team <reply@example.com>" }));
  assert.equal(result.flags.some(flag => flag.startsWith("Reply-To")), false);
});

test("checks urgency language in the subject", () => {
  const result = runRuleEngine(email({ subject: "URGENT: immediate action required" }));
  assert.equal(result.flags.includes("Urgency language detected"), true);
  assert.equal(result.level, "warning");
});

test("does not use broad edit distance for very short trusted domains", () => {
  assert.equal(checkLookalikeDomaIn("y.com"), null);
  assert.equal(checkLookalikeDomaIn("up.com"), null);
  assert.equal(checkLookalikeDomaIn("alerts.microsoft.com"), null);
  assert.equal(checkLookalikeDomaIn("paypa1.com"), "paypal.com");
  assert.equal(checkLookalikeDomaIn("mail.paypa1.com"), "paypal.com");
});

test("AI caution raises a rule-safe result", () => {
  const rules = runRuleEngine(email());
  const result = reconcileAiVerdict(rules, "CAUTION");
  assert.equal(result.score, 25);
  assert.equal(result.level, "warning");
});

test("AI safe cannot lower a dangerous rule result", () => {
  const rules = runRuleEngine(email({ bodyText: "Urgent: send your password and credit card now." }));
  const result = reconcileAiVerdict(rules, "SAFE");
  assert.equal(result.score, rules.score);
  assert.equal(result.level, rules.level);
});
