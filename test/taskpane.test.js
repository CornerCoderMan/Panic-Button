"use strict";

const test = require("node:test");
const assert = require("node:assert/strict");
const {
  SUPPORT_ADDRESS,
  checkLookalikeDomaIn,
  extractDomain,
  extractMailboxAddress,
  runRuleEngine,
  buildReportSubject,
  buildReportHtml,
  isHighImpact,
  attachmentName,
  esc,
} = require("../taskpane.js");

function email(overrides = {}) {
  return {
    itemId: "AAMkAExampleItemId==",
    subject: "Weekly project update",
    senderName: "Example Sender",
    senderEmail: "sender@example.com",
    toRecipients: ["me@mattnj.com"],
    replyTo: null,
    bodyText: "Here is the normal weekly update.",
    links: [],
    attachments: [],
    receivedTime: null,
    fullHeaders: "Received: from mail.example.com\r\nFrom: sender@example.com\r\nSubject: Weekly project update",
    ...overrides,
  };
}

function answers(overrides = {}) {
  return {
    clicked: "No",
    credentials: "No",
    attachment: "No",
    replied: "No",
    when: "Earlier today",
    notes: "",
    ...overrides,
  };
}

// ── Utility parsing (unchanged behaviour) ─────────────────────────────────────

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

// ── Report subject ────────────────────────────────────────────────────────────

test("subject is neutral when no high-impact action was taken", () => {
  const subject = buildReportSubject(email(), answers());
  assert.equal(subject, "Suspicious email report: Weekly project update");
});

test("subject is flagged [ACTION NEEDED] when the user clicked or entered credentials", () => {
  assert.match(buildReportSubject(email(), answers({ clicked: "Yes" })), /^\[ACTION NEEDED\] /);
  assert.match(buildReportSubject(email(), answers({ credentials: "Yes" })), /^\[ACTION NEEDED\] /);
  assert.match(buildReportSubject(email(), answers({ replied: "Yes" })), /^\[ACTION NEEDED\] /);
});

test("isHighImpact only trips on Yes answers", () => {
  assert.equal(isHighImpact(answers()), false);
  assert.equal(isHighImpact(answers({ credentials: "Not sure" })), false);
  assert.equal(isHighImpact(answers({ attachment: "Yes" })), false); // opening an attachment alone isn't credential exposure
  assert.equal(isHighImpact(answers({ clicked: "Yes" })), true);
});

// ── Report body ───────────────────────────────────────────────────────────────

test("report body includes the reporter's answers, headers, and support context", () => {
  const rules = runRuleEngine(email());
  const html = buildReportHtml(email(), answers({ clicked: "Yes", notes: "I typed my password" }), rules);
  assert.match(html, /Suspicious Email Report/);
  assert.match(html, /Full internet headers/);
  assert.match(html, /Received: from mail\.example\.com/);
  assert.match(html, /I typed my password/);
  assert.match(html, /Possible account exposure/); // high-impact banner present
  assert.match(html, /original message is attached/i);
});

test("report body warns when the original message could not be attached", () => {
  const rules = runRuleEngine(email({ itemId: null }));
  const html = buildReportHtml(email({ itemId: null }), answers(), rules);
  assert.match(html, /could not be attached automatically/i);
});

test("report escapes HTML in email content to prevent injection", () => {
  const evil = email({ subject: "<script>alert(1)</script>", senderName: "<b>x</b>" });
  const html = buildReportHtml(evil, answers(), runRuleEngine(evil));
  assert.equal(html.includes("<script>alert(1)</script>"), false);
  assert.match(html, /&lt;script&gt;/);
});

test("esc encodes the five significant HTML characters", () => {
  assert.equal(esc(`<>&"'`), "&lt;&gt;&amp;&quot;&#39;");
});

// ── Config ────────────────────────────────────────────────────────────────────

test("reports are addressed to the support mailbox", () => {
  assert.equal(SUPPORT_ADDRESS, "support@mattnj.com");
});

test("attachment name is filesystem-safe and prefixed", () => {
  assert.equal(attachmentName('Invoice: Q3/2026 *final*?'), "Reported - Invoice Q3 2026 final");
  assert.equal(attachmentName(""), "Reported message");
});
