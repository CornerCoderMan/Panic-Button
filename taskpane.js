/* ============================================================
   Report This! — Task Pane Logic
   ============================================================

   Post-click incident reporter for Outlook.

   Flow:
     1. User opens a suspicious message and clicks "Report This!"
     2. Task pane asks a few triage questions (clicked? entered
        credentials? opened attachment? replied? when?)
     3. Office.js reads the message (sender, subject, links,
        attachments, body excerpt) and the FULL internet headers
     4. A lightweight rule engine adds "automated flags" for context
     5. A pre-filled report opens as a new message to the support
        team, with the ORIGINAL message forwarded as an attachment
     6. The user reviews and clicks Send — nothing leaves the
        mailbox automatically

   Runs entirely client-side. Requires Mailbox requirement set 1.8
   (for getAllInternetHeadersAsync). Permission: ReadItem.
   ============================================================ */

"use strict";

// ── Configuration ─────────────────────────────────────────────────────────────

const SUPPORT_ADDRESS = "support@mattnj.com";

// ── Constants (rule engine) ───────────────────────────────────────────────────

// Known URL shortener domains (non-exhaustive — extend as needed)
const URL_SHORTENERS = new Set([
  "bit.ly","tinyurl.com","t.co","goo.gl","ow.ly","buff.ly","rebrand.ly",
  "short.link","is.gd","v.gd","cutt.ly","tiny.cc","shorturl.at","rb.gy",
  "clck.ru","0rz.tw","1url.com","2tu.us","4sq.com","u.to"
]);

// High-risk attachment extensions
const DANGEROUS_EXTENSIONS = new Set([
  "exe","bat","cmd","vbs","vbe","js","jse","wsf","wsh","msc","scr",
  "ps1","ps2","psc1","psc2","reg","lnk","pif","application","gadget",
  "hta","cpl","msp","msi","jar","com","inf","sys","dll"
]);

// Urgency / social-engineering keywords (scored by severity)
const URGENCY_PATTERNS = [
  { re: /\b(act now|immediate action required|urgent|urgently)\b/i, weight: 15, label: "Urgency language detected" },
  { re: /\b(verify your account|confirm your (identity|details|account))\b/i, weight: 20, label: "Account verification request" },
  { re: /\b(password (expired|reset|change)|update your credentials)\b/i, weight: 20, label: "Credential reset request" },
  { re: /\b(you (have won|are a winner)|congratulations.*prize|lottery)\b/i, weight: 25, label: "Prize / lottery language" },
  { re: /\b(click here to (claim|verify|confirm|unlock|restore))\b/i, weight: 20, label: "Generic 'click here' instruction" },
  { re: /\b(suspended|compromised|unauthorized access|unusual (sign-in|activity))\b/i, weight: 20, label: "Account threat language" },
  { re: /\b(invoice attached|payment (overdue|required|pending))\b/i, weight: 10, label: "Financial urgency language" },
  { re: /\b(dear (customer|user|member|account holder))\b/i, weight: 10, label: "Generic salutation (not personalized)" },
  { re: /\b(your (account|access) (will be|has been) (locked|disabled|suspended))\b/i, weight: 20, label: "Account lockout threat" },
];

// ── State ─────────────────────────────────────────────────────────────────────

let officeReady = false;

// ── Office initialization ─────────────────────────────────────────────────────

if (typeof Office !== "undefined") {
  Office.onReady(info => {
    if (info.host === Office.HostType.Outlook) {
      officeReady = true;
      initUI();
    } else {
      showError("This add-in only works in Outlook.");
    }
  });
}

// ── UI initialization ─────────────────────────────────────────────────────────

function initUI() {
  const supportLabel = el("support-addr");
  if (supportLabel) supportLabel.textContent = SUPPORT_ADDRESS;

  el("report-form").addEventListener("submit", event => {
    event.preventDefault();
    submitReport();
  });
  el("retry-btn").addEventListener("click", resetToForm);
  el("another-btn").addEventListener("click", resetToForm);

  setFooterStatus("Ready");
}

// ── Main flow ─────────────────────────────────────────────────────────────────

async function submitReport() {
  showLoading("Reading this email…");

  try {
    const answers = collectAnswers();
    const emailData = await readEmail();

    setLoadingMessage("Running automated checks…");
    const ruleResult = runRuleEngine(emailData);

    setLoadingMessage("Preparing report…");
    const subject = buildReportSubject(emailData, answers);
    const htmlBody = buildReportHtml(emailData, answers, ruleResult);

    const attachedOriginal = openReportDraft(emailData, subject, htmlBody);
    showDone(attachedOriginal);

  } catch (err) {
    console.error("Report failed:", err);
    showError(err.message || "An unexpected error occurred while building the report.");
  }
}

// Collect the triage answers from the form.
function collectAnswers() {
  const radio = name => {
    const checked = document.querySelector(`input[name="${name}"]:checked`);
    return checked ? checked.value : "Not sure";
  };
  return {
    clicked:     radio("clicked"),
    credentials: radio("credentials"),
    attachment:  radio("attachment"),
    replied:     radio("replied"),
    when:        el("when-select") ? el("when-select").value : "Not sure",
    notes:       el("notes") ? el("notes").value.trim() : "",
  };
}

// True if the user took a high-impact action (drives urgency in the report).
function isHighImpact(answers) {
  return answers.clicked === "Yes" || answers.credentials === "Yes" || answers.replied === "Yes";
}

// ── Email reader ──────────────────────────────────────────────────────────────

function readEmail() {
  return new Promise((resolve, reject) => {
    const item = Office.context.mailbox.item;
    if (!item) { reject(new Error("No email is currently open. Open a message, then click Report This!.")); return; }

    const emailData = {
      itemId:       item.itemId || null,
      subject:      item.subject || "",
      senderName:   item.sender?.displayName || item.from?.displayName || "",
      senderEmail:  item.sender?.emailAddress || item.from?.emailAddress || "",
      toRecipients: (item.to || []).map(r => r.emailAddress),
      replyTo:      null,
      bodyText:     "",
      links:        [],
      attachments:  (item.attachments || []).map(a => ({
        name: a.name || "",
        type: a.attachmentType,
        contentType: a.contentType || ""
      })),
      receivedTime: item.dateTimeCreated || null,
      fullHeaders:  "",
    };

    // Full internet headers (Mailbox 1.8+). Also used to derive Reply-To.
    if (item.getAllInternetHeadersAsync) {
      item.getAllInternetHeadersAsync(headersResult => {
        if (headersResult.status === Office.AsyncResultStatus.Succeeded && headersResult.value) {
          emailData.fullHeaders = headersResult.value;
          const unfoldedHeaders = headersResult.value.replace(/\r?\n[\t ]+/g, " ");
          const replyToMatch = unfoldedHeaders.match(/^Reply-To:\s*(.+)$/im);
          if (replyToMatch) {
            emailData.replyTo = extractMailboxAddress(replyToMatch[1]) || replyToMatch[1].trim();
          }
        } else {
          emailData.fullHeaders = "(Full internet headers were not available on this Outlook client or version.)";
        }
        getBody(item, emailData, resolve, reject);
      });
    } else {
      emailData.fullHeaders = "(Full internet headers are not supported by this Outlook client — requires Mailbox 1.8+.)";
      getBody(item, emailData, resolve, reject);
    }
  });
}

function getBody(item, emailData, resolve, reject) {
  item.body.getAsync(Office.CoercionType.Text, { asyncContext: null }, result => {
    if (result.status !== Office.AsyncResultStatus.Succeeded) {
      reject(new Error("Could not read the email body."));
      return;
    }
    emailData.bodyText = result.value || "";

    item.body.getAsync(Office.CoercionType.Html, {}, htmlResult => {
      if (htmlResult.status === Office.AsyncResultStatus.Succeeded && htmlResult.value) {
        emailData.links = extractLinksFromHtml(htmlResult.value);
      }
      resolve(emailData);
    });
  });
}

// ── Report draft ──────────────────────────────────────────────────────────────

// Opens a pre-filled new message to support. Returns true if the original
// message was attached, false if it had to be skipped (no item id available).
function openReportDraft(emailData, subject, htmlBody) {
  const parameters = {
    toRecipients: [SUPPORT_ADDRESS],
    subject,
    htmlBody,
  };

  let attachedOriginal = false;
  if (emailData.itemId) {
    parameters.attachments = [{
      type: "item",
      itemId: emailData.itemId,
      name: attachmentName(emailData.subject),
    }];
    attachedOriginal = true;
  }

  Office.context.mailbox.displayNewMessageForm(parameters);
  return attachedOriginal;
}

function attachmentName(subject) {
  const clean = String(subject || "").replace(/[\\/:*?"<>|]+/g, " ").replace(/\s+/g, " ").trim().slice(0, 80);
  return clean ? `Reported - ${clean}` : "Reported message";
}

// ── Report content builders (pure — unit tested) ──────────────────────────────

function buildReportSubject(emailData, answers) {
  const base = String(emailData.subject || "(no subject)").replace(/[\r\n]+/g, " ").trim();
  const prefix = isHighImpact(answers) ? "[ACTION NEEDED] " : "";
  return `${prefix}Suspicious email report: ${base}`.slice(0, 255);
}

function buildReportHtml(emailData, answers, ruleResult) {
  const rows = [
    ["Reported by", currentUserEmail()],
    ["Clicked a link?", answers.clicked],
    ["Entered credentials / personal info?", answers.credentials],
    ["Opened / downloaded an attachment?", answers.attachment],
    ["Replied or sent information back?", answers.replied],
    ["When did it happen?", answers.when],
  ];

  const impactBanner = isHighImpact(answers)
    ? `<p style="margin:0 0 14px;padding:10px 12px;background:#fde7e9;border:1px solid #f4abab;border-radius:6px;color:#a4262c;">
         <strong>⚠ Possible account exposure.</strong> The reporter indicated they clicked a link, entered
         information, or replied. Treat this as time-sensitive.
       </p>`
    : "";

  const answersTable = rows.map(([k, v]) =>
    `<tr><td style="padding:3px 10px 3px 0;color:#605e5c;white-space:nowrap;vertical-align:top;">${esc(k)}</td>` +
    `<td style="padding:3px 0;"><strong>${esc(v || "—")}</strong></td></tr>`
  ).join("");

  const notesBlock = answers.notes
    ? `<h3 style="margin:16px 0 4px;font-size:13px;">Notes from the reporter</h3>
       <div style="white-space:pre-wrap;padding:8px 10px;background:#faf9f8;border:1px solid #edebe9;border-radius:6px;">${esc(answers.notes)}</div>`
    : "";

  const msgRows = [
    ["From (display name)", emailData.senderName],
    ["From (address)", emailData.senderEmail],
    ["Reply-To", emailData.replyTo || "(same as sender or not set)"],
    ["To", (emailData.toRecipients || []).join(", ")],
    ["Subject", emailData.subject],
    ["Received", formatDate(emailData.receivedTime)],
  ].map(([k, v]) =>
    `<tr><td style="padding:3px 10px 3px 0;color:#605e5c;white-space:nowrap;vertical-align:top;">${esc(k)}</td>` +
    `<td style="padding:3px 0;">${esc(v || "—")}</td></tr>`
  ).join("");

  const linksBlock = (emailData.links && emailData.links.length)
    ? `<h3 style="margin:16px 0 4px;font-size:13px;">Links in the message (${emailData.links.length})</h3>
       <ul style="margin:0;padding-left:18px;">` +
      emailData.links.slice(0, 40).map(l =>
        `<li style="margin-bottom:3px;">${esc(l.display || "(no text)")} &rarr; <code>${esc(l.href)}</code></li>`
      ).join("") +
      (emailData.links.length > 40 ? `<li>… and ${emailData.links.length - 40} more</li>` : "") +
      `</ul>`
    : `<p style="margin:16px 0 0;color:#605e5c;">No links found in the message body.</p>`;

  const attBlock = (emailData.attachments && emailData.attachments.length)
    ? `<h3 style="margin:16px 0 4px;font-size:13px;">Attachments (${emailData.attachments.length})</h3>
       <ul style="margin:0;padding-left:18px;">` +
      emailData.attachments.map(a => `<li>${esc(a.name)}${a.contentType ? ` <span style="color:#605e5c;">(${esc(a.contentType)})</span>` : ""}</li>`).join("") +
      `</ul>`
    : `<p style="margin:16px 0 0;color:#605e5c;">No attachments on the message.</p>`;

  const flagsBlock = (ruleResult && ruleResult.flags && ruleResult.flags.length)
    ? `<h3 style="margin:16px 0 4px;font-size:13px;">Automated flags (score ${ruleResult.score}/100 — ${esc(ruleResult.verdict)})</h3>
       <ul style="margin:0;padding-left:18px;">` +
      ruleResult.flags.map(f => `<li style="margin-bottom:3px;color:#a4262c;">${esc(f)}</li>`).join("") +
      `</ul>`
    : `<p style="margin:16px 0 0;color:#605e5c;">Automated checks (score ${ruleResult ? ruleResult.score : 0}/100) found no specific red flags. This does not mean the message is safe.</p>`;

  const attachmentNote = emailData.itemId
    ? `<p style="margin:0 0 14px;color:#605e5c;">The original message is attached to this report for full analysis.</p>`
    : `<p style="margin:0 0 14px;padding:8px 10px;background:#fff4ce;border:1px solid #f0d86e;border-radius:6px;color:#6d4b00;">
         The original message could not be attached automatically on this client. Please attach or forward it manually.
       </p>`;

  return `<div style="font-family:'Segoe UI',system-ui,sans-serif;font-size:13px;color:#323130;line-height:1.5;">
  <h2 style="margin:0 0 10px;font-size:16px;">🚩 Suspicious Email Report</h2>
  ${impactBanner}
  ${attachmentNote}
  <h3 style="margin:0 0 4px;font-size:13px;">What the reporter said</h3>
  <table style="border-collapse:collapse;font-size:13px;">${answersTable}</table>
  ${notesBlock}
  <h3 style="margin:16px 0 4px;font-size:13px;">Message details</h3>
  <table style="border-collapse:collapse;font-size:13px;">${msgRows}</table>
  ${linksBlock}
  ${attBlock}
  ${flagsBlock}
  <h3 style="margin:16px 0 4px;font-size:13px;">Full internet headers</h3>
  <pre style="white-space:pre-wrap;word-break:break-word;background:#faf9f8;border:1px solid #edebe9;border-radius:6px;padding:10px;font-size:11px;overflow-x:auto;">${esc(emailData.fullHeaders)}</pre>
  <p style="margin:14px 0 0;color:#a19f9d;font-size:11px;">Generated by Report This! · Review before sending.</p>
</div>`;
}

function currentUserEmail() {
  try {
    return Office.context.mailbox.userProfile?.emailAddress || "";
  } catch {
    return "";
  }
}

function formatDate(value) {
  if (!value) return "";
  try {
    const d = value instanceof Date ? value : new Date(value);
    return isNaN(d.getTime()) ? String(value) : d.toLocaleString();
  } catch {
    return String(value);
  }
}

// ── Rule engine ───────────────────────────────────────────────────────────────

function runRuleEngine(email) {
  const flags   = [];  // Red flag strings
  const safeOk  = [];  // Safe indicator strings
  let   score   = 0;   // 0–100, higher = riskier

  const senderEmail = String(email.senderEmail || "").toLowerCase();
  const senderDomain = extractDomain(senderEmail);

  if (email.senderName) {
    const nameEmailMatch = email.senderName.match(/[a-zA-Z0-9._%+\-]+@([a-zA-Z0-9.\-]+\.[a-zA-Z]{2,})/);
    if (nameEmailMatch) {
      const nameDomain = nameEmailMatch[1].toLowerCase();
      if (nameDomain && senderDomain && nameDomain !== senderDomain) {
        flags.push(`Sender name contains "${nameEmailMatch[0]}" but actual email is ${senderEmail} — possible spoofing`);
        score += 35;
      }
    }
  }

  const lookalikes = checkLookalikeDomaIn(senderDomain);
  if (lookalikes) {
    flags.push(`Sender domain "${senderDomain}" resembles trusted domain "${lookalikes}" — possible typosquatting`);
    score += 40;
  }

  const freeEmailHosts = ["gmail.com","yahoo.com","hotmail.com","outlook.com","aol.com","protonmail.com","icloud.com"];
  const senderNameLower = String(email.senderName || "").toLowerCase();
  const looksLikeBusiness = (
    senderNameLower.includes("support") ||
    senderNameLower.includes("service") ||
    senderNameLower.includes("security") ||
    senderNameLower.includes("help") ||
    senderNameLower.includes("billing") ||
    senderNameLower.includes("admin") ||
    senderNameLower.includes("noreply") ||
    senderNameLower.includes("no-reply")
  );
  if (looksLikeBusiness && freeEmailHosts.includes(senderDomain)) {
    flags.push(`"${email.senderName}" is using a personal email provider (${senderDomain}) — businesses usually use their own domain`);
    score += 25;
  }

  if (email.replyTo) {
    const replyDomain = extractDomain(String(email.replyTo).toLowerCase());
    if (replyDomain && senderDomain && replyDomain !== senderDomain) {
      flags.push(`Reply-To (${email.replyTo}) is a different domain than the sender (${senderDomain})`);
      score += 20;
    }
  }

  const subject = String(email.subject || "");

  if (/[!?]{2,}/.test(subject)) {
    flags.push("Subject line uses multiple exclamation/question marks");
    score += 8;
  }
  if (subject.replace(/[^A-Z]/g, "").length > 0.4 * subject.replace(/[^a-zA-Z]/g, "").length && subject.length > 8) {
    flags.push("Subject line is written in excessive CAPITALS");
    score += 10;
  }

  const body = String(email.bodyText || "");
  const messageText = `${subject}\n${body}`;

  for (const p of URGENCY_PATTERNS) {
    if (p.re.test(messageText)) {
      flags.push(p.label);
      score += p.weight;
    }
  }

  if (/\b(social security|ssn|date of birth|credit card|bank account|routing number|mother['']?s maiden|password)\b/i.test(body)) {
    flags.push("Email asks for sensitive personal or financial information");
    score += 30;
  }

  if (/call\s+(us|now|immediately)\s+(at|on)?\s*[\+\d\s\-\(\)]{7,}/i.test(body)) {
    flags.push("Email pressures you to call a phone number immediately");
    score += 15;
  }

  let ipLinkCount = 0, shortenerCount = 0, mismatchCount = 0, httpCount = 0;

  for (const link of (email.links || [])) {
    const { href, display } = link;

    if (/^https?:\/\/\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}/i.test(href)) ipLinkCount++;

    const hrefDomain = extractDomain(String(href).toLowerCase());
    if (URL_SHORTENERS.has(hrefDomain)) shortenerCount++;

    if (/https?:\/\//i.test(display)) {
      const dispDomain = extractDomain(String(display).toLowerCase());
      if (dispDomain && hrefDomain && dispDomain !== hrefDomain) mismatchCount++;
    }

    if (/^http:\/\//i.test(href)) httpCount++;
  }

  if (ipLinkCount > 0)    { flags.push(`${ipLinkCount} link(s) go to raw IP addresses instead of domain names`); score += ipLinkCount * 20; }
  if (shortenerCount > 0) { flags.push(`${shortenerCount} link(s) use URL shorteners — destination hidden`); score += shortenerCount * 15; }
  if (mismatchCount > 0)  { flags.push(`${mismatchCount} link(s) show a different URL in text than the actual destination`); score += mismatchCount * 25; }
  if (httpCount > 0 && (email.links || []).length > 0) { flags.push(`${httpCount} link(s) use unencrypted HTTP instead of HTTPS`); score += httpCount * 5; }

  for (const att of (email.attachments || [])) {
    const ext = String(att.name || "").split(".").pop().toLowerCase();
    if (DANGEROUS_EXTENSIONS.has(ext)) {
      flags.push(`Attachment "${att.name}" has a potentially dangerous file type (.${ext})`);
      score += 40;
    }
    const parts = String(att.name || "").split(".");
    if (parts.length >= 3) {
      const outerExt = parts[parts.length - 1].toLowerCase();
      const innerExt = parts[parts.length - 2].toLowerCase();
      if (DANGEROUS_EXTENSIONS.has(outerExt) && ["pdf","doc","docx","xls","xlsx","jpg","png"].includes(innerExt)) {
        flags.push(`Attachment "${att.name}" uses a double extension to disguise its type`);
        score += 50;
      }
    }
  }

  if (flags.length === 0) safeOk.push("No spoofing or mismatched sender information detected");
  if ((email.links || []).length > 0 && ipLinkCount === 0 && shortenerCount === 0 && mismatchCount === 0 && httpCount === 0)
    safeOk.push("No obvious URL hiding or unencrypted links detected");
  if ((email.attachments || []).length === 0) safeOk.push("No attachments");
  if (!URGENCY_PATTERNS.some(p => p.re.test(messageText))) safeOk.push("No urgency or social-engineering language found");

  score = Math.min(100, score);

  let verdict, level;
  if (score >= 60)                         { verdict = "Likely Unsafe"; level = "danger"; }
  else if (score >= 25 || flags.length > 0){ verdict = "Exercise Caution"; level = "warning"; }
  else                                     { verdict = "No Obvious Warning Signs"; level = "safe"; }

  return { score, verdict, level, flags, safeOk };
}

// ── Lookalike domain detection ────────────────────────────────────────────────

function checkLookalikeDomaIn(domain) {
  if (!domain) return null;
  const trusted = [
    "paypal.com","amazon.com","microsoft.com","google.com","apple.com",
    "facebook.com","instagram.com","netflix.com","chase.com","wellsfargo.com",
    "bankofamerica.com","citibank.com","irs.gov","usps.com","fedex.com","ups.com",
    "dropbox.com","linkedin.com","twitter.com","x.com","docusign.com",
  ];

  for (const t of trusted) {
    if (domain === t || domain.endsWith(`.${t}`)) return null;
    const tParts = t.split(".");
    const dParts = domain.split(".");
    const tBase = tParts.length > 1 ? tParts[tParts.length - 2] : t;
    const dBase = dParts.length > 1 ? dParts[dParts.length - 2] : domain;

    const threshold = tBase.length >= 8 ? 2 : tBase.length >= 5 ? 1 : 0;
    if (threshold > 0 && levenshtein(dBase, tBase) <= threshold && domain !== t) {
      return t;
    }
  }
  return null;
}

function levenshtein(a, b) {
  const m = a.length, n = b.length;
  const dp = Array.from({length: m+1}, (_, i) => Array.from({length: n+1}, (_, j) => j === 0 ? i : 0));
  for (let j = 0; j <= n; j++) dp[0][j] = j;
  for (let i = 1; i <= m; i++) {
    for (let j = 1; j <= n; j++) {
      dp[i][j] = a[i-1] === b[j-1] ? dp[i-1][j-1] : 1 + Math.min(dp[i-1][j], dp[i][j-1], dp[i-1][j-1]);
    }
  }
  return dp[m][n];
}

// ── Link extraction ───────────────────────────────────────────────────────────

function extractLinksFromHtml(html) {
  if (typeof DOMParser !== "undefined") {
    const doc = new DOMParser().parseFromString(html, "text/html");
    return Array.from(doc.querySelectorAll("a[href]")).map(anchor => ({
      href: anchor.getAttribute("href").trim(),
      display: anchor.textContent.replace(/\s+/g, " ").trim(),
    }));
  }

  const links = [];
  const hrefRe = /<a\s[^>]*href=["']([^"']+)["'][^>]*>([\s\S]*?)<\/a>/gi;
  let m;
  while ((m = hrefRe.exec(html)) !== null) {
    links.push({ href: m[1].trim(), display: stripTags(m[2]).trim() });
  }
  return links;
}

function stripTags(html) {
  return html.replace(/<[^>]+>/g, "").replace(/\s+/g, " ");
}

// ── UI state helpers ──────────────────────────────────────────────────────────

function showLoading(msg) {
  hideAll();
  el("loading-state").classList.remove("hidden");
  setLoadingMessage(msg);
}

function setLoadingMessage(msg) {
  el("loading-message").textContent = msg;
}

function showDone(attachedOriginal) {
  hideAll();
  el("done-state").classList.remove("hidden");
  const fallback = el("done-fallback");
  if (!attachedOriginal) {
    fallback.textContent = "Note: the original message couldn't be attached automatically on this Outlook client. Please attach or forward it manually before sending.";
    fallback.classList.remove("hidden");
  } else {
    fallback.classList.add("hidden");
  }
  setFooterStatus(`Report drafted: ${new Date().toLocaleTimeString()}`);
}

function showError(msg) {
  hideAll();
  el("error-state").classList.remove("hidden");
  el("error-message").textContent = msg;
}

function resetToForm() {
  hideAll();
  el("form-state").classList.remove("hidden");
  setFooterStatus("Ready");
}

function hideAll() {
  ["form-state","loading-state","done-state","error-state"].forEach(id => {
    const node = el(id);
    if (node) node.classList.add("hidden");
  });
}

function setFooterStatus(msg) {
  const node = el("footer-status");
  if (node) node.textContent = msg;
}

// ── Utilities ─────────────────────────────────────────────────────────────────

function el(id) { return document.getElementById(id); }

function esc(value) {
  return String(value == null ? "" : value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function extractDomain(str) {
  const value = String(str || "").trim().toLowerCase();
  const mailbox = extractMailboxAddress(value);
  if (mailbox) return mailbox.split("@").pop();

  try {
    const url = /^https?:\/\//i.test(value) ? new URL(value) : new URL("https://" + value);
    return url.hostname.replace(/^www\./, "").replace(/\.$/, "");
  } catch {
    return value.replace(/[<>]/g, "");
  }
}

function extractMailboxAddress(value) {
  const text = String(value || "").trim();
  const angleAddress = text.match(/<\s*([^<>\s]+@[^<>\s]+)\s*>/);
  if (angleAddress) return angleAddress[1].toLowerCase();

  const plainAddress = text.match(/\b[A-Z0-9._%+\-]+@[A-Z0-9.\-]+\.[A-Z]{2,}\b/i);
  return plainAddress ? plainAddress[0].toLowerCase() : null;
}

// ── Exports for tests (Node) ──────────────────────────────────────────────────

if (typeof module !== "undefined" && module.exports) {
  module.exports = {
    SUPPORT_ADDRESS,
    checkLookalikeDomaIn,
    extractDomain,
    extractMailboxAddress,
    extractLinksFromHtml,
    levenshtein,
    runRuleEngine,
    buildReportSubject,
    buildReportHtml,
    isHighImpact,
    attachmentName,
    esc,
  };
}
