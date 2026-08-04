/* ============================================================
   Email Safety Checker — Task Pane Logic
   ============================================================

   Architecture:
     1. Office.js reads the current email (sender, subject, body, attachments, headers)
     2. Rule engine runs synchronous checks and produces a risk score + flags
     3. Sends a limited excerpt to the server-side AI proxy when configured
     4. UI renders verdict, score, red flags, and suggested actions

   Provider credentials remain on the server and are never exposed to the pane.
   ============================================================ */

"use strict";

// ── Constants ────────────────────────────────────────────────────────────────

const AI_PROXY_URL = "/api/analyze";
const AI_ENABLED = false; // GitHub Pages pilot: no server-side proxy is available.

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

// Patterns that reduce suspicion
const SAFE_INDICATORS = [
  { re: /^[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}$/, type: "sender", label: "Sender email is well-formed" },
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
  // Primary analyze button
  el("analyze-btn").addEventListener("click", runAnalysis);
  el("reanalyze-btn").addEventListener("click", resetToIdle);
  el("retry-btn").addEventListener("click", runAnalysis);

  setFooterStatus("Ready");
}

// ── Main analysis flow ────────────────────────────────────────────────────────

async function runAnalysis() {
  showLoading("Reading email…");

  try {
    const emailData = await readEmail();
    setLoadingMessage("Running safety checks…");

    const ruleResult = runRuleEngine(emailData);

    let aiResult = null;
    let aiWarning = null;

    if (AI_ENABLED) {
      setLoadingMessage("Checking optional AI analysis…");
      try {
        aiResult = await callAiProxy(emailData, ruleResult);
      } catch (aiErr) {
        console.warn("AI analysis unavailable:", aiErr.message);
        aiWarning = `${aiErr.message} Results below use rule-based checks only.`;
      }
    }

    const finalResult = reconcileAiVerdict(ruleResult, aiResult?.verdict);
    showResults(finalResult, aiResult?.text || null, aiWarning);

  } catch (err) {
    console.error("Analysis failed:", err);
    showError(err.message || "An unexpected error occurred.");
  }
}

// ── Email reader ──────────────────────────────────────────────────────────────

function readEmail() {
  return new Promise((resolve, reject) => {
    const item = Office.context.mailbox.item;
    if (!item) { reject(new Error("No email is currently open.")); return; }

    const emailData = {
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
      receivedTime: item.dateTimeCreated,
    };

    // Try to get reply-to from internet headers
    if (item.getAllInternetHeadersAsync) {
      item.getAllInternetHeadersAsync(headersResult => {
        if (headersResult.status === Office.AsyncResultStatus.Succeeded) {
          const unfoldedHeaders = headersResult.value.replace(/\r?\n[\t ]+/g, " ");
          const replyToMatch = unfoldedHeaders.match(/^Reply-To:\s*(.+)$/im);
          if (replyToMatch) {
            emailData.replyTo = extractMailboxAddress(replyToMatch[1]) || replyToMatch[1].trim();
          }
        }
        // Now get body
        getBody(item, emailData, resolve, reject);
      });
    } else {
      getBody(item, emailData, resolve, reject);
    }
  });
}

function getBody(item, emailData, resolve, reject) {
  item.body.getAsync(Office.CoercionType.Text, { asyncContext: null }, result => {
    if (result.status !== Office.AsyncResultStatus.Succeeded) {
      reject(new Error("Could not read email body."));
      return;
    }
    emailData.bodyText = result.value || "";

    // Extract links from HTML body for deeper analysis
    item.body.getAsync(Office.CoercionType.Html, {}, htmlResult => {
      if (htmlResult.status === Office.AsyncResultStatus.Succeeded && htmlResult.value) {
        emailData.links = extractLinksFromHtml(htmlResult.value);
      }
      resolve(emailData);
    });
  });
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
  // Match <a href="...">display text</a>
  const hrefRe = /<a\s[^>]*href=["']([^"']+)["'][^>]*>([\s\S]*?)<\/a>/gi;
  let m;
  while ((m = hrefRe.exec(html)) !== null) {
    const href    = m[1].trim();
    const display = stripTags(m[2]).trim();
    links.push({ href, display });
  }
  return links;
}

function stripTags(html) {
  return html.replace(/<[^>]+>/g, "").replace(/\s+/g, " ");
}

// ── Rule engine ───────────────────────────────────────────────────────────────

function runRuleEngine(email) {
  const flags   = [];  // Red flag strings
  const safeOk  = [];  // Safe indicator strings
  let   score   = 0;   // 0–100, higher = riskier

  // ── 1. Sender analysis ────────────────────────────────────────────────────

  const senderEmail = email.senderEmail.toLowerCase();
  const senderDomain = extractDomain(senderEmail);

  // Display name contains a different domain than the actual email
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

  // Lookalike domains (e.g. paypa1.com, micosoft.com)
  const lookalikes = checkLookalikeDomaIn(senderDomain);
  if (lookalikes) {
    flags.push(`Sender domain "${senderDomain}" resembles trusted domain "${lookalikes}" — possible typosquatting`);
    score += 40;
  }

  // Free email hosting for a "business" sender
  const freeEmailHosts = ["gmail.com","yahoo.com","hotmail.com","outlook.com","aol.com","protonmail.com","icloud.com"];
  const senderNameLower = email.senderName.toLowerCase();
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

  // Reply-to differs from sender
  if (email.replyTo) {
    const replyDomain = extractDomain(email.replyTo.toLowerCase());
    if (replyDomain && senderDomain && replyDomain !== senderDomain) {
      flags.push(`Reply-To (${email.replyTo}) is a different domain than the sender (${senderDomain})`);
      score += 20;
    }
  }

  // ── 2. Subject analysis ───────────────────────────────────────────────────

  const subject = email.subject;

  // Excessive punctuation / caps
  if (/[!?]{2,}/.test(subject)) {
    flags.push("Subject line uses multiple exclamation/question marks");
    score += 8;
  }
  if (subject.replace(/[^A-Z]/g, "").length > 0.4 * subject.replace(/[^a-zA-Z]/g, "").length && subject.length > 8) {
    flags.push("Subject line is written in excessive CAPITALS");
    score += 10;
  }

  // ── 3. Body analysis ─────────────────────────────────────────────────────

  const body = email.bodyText;
  const messageText = `${subject}\n${body}`;
  const bodyLower = body.toLowerCase();

  // Urgency patterns
  for (const p of URGENCY_PATTERNS) {
    if (p.re.test(messageText)) {
      flags.push(p.label);
      score += p.weight;
    }
  }

  // Asks for sensitive info
  if (/\b(social security|ssn|date of birth|credit card|bank account|routing number|mother['']?s maiden|password)\b/i.test(body)) {
    flags.push("Email asks for sensitive personal or financial information");
    score += 30;
  }

  // Contains phone number with pressured CTA
  if (/call\s+(us|now|immediately)\s+(at|on)?\s*[\+\d\s\-\(\)]{7,}/i.test(body)) {
    flags.push("Email pressures you to call a phone number immediately");
    score += 15;
  }

  // ── 4. Link analysis ─────────────────────────────────────────────────────

  let ipLinkCount      = 0;
  let shortenerCount   = 0;
  let mismatchCount    = 0;
  let httpCount        = 0;

  for (const link of email.links) {
    const { href, display } = link;

    // IP-based URL
    if (/^https?:\/\/\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}/i.test(href)) {
      ipLinkCount++;
    }

    // URL shortener
    const hrefDomain = extractDomain(href.toLowerCase());
    if (URL_SHORTENERS.has(hrefDomain)) {
      shortenerCount++;
    }

    // Mismatched display text vs href (display looks like a URL but differs)
    if (/https?:\/\//i.test(display)) {
      const dispDomain = extractDomain(display.toLowerCase());
      if (dispDomain && hrefDomain && dispDomain !== hrefDomain) {
        mismatchCount++;
      }
    }

    // Plain HTTP (not HTTPS)
    if (/^http:\/\//i.test(href)) {
      httpCount++;
    }
  }

  if (ipLinkCount > 0) {
    flags.push(`${ipLinkCount} link(s) go to raw IP addresses instead of domain names`);
    score += ipLinkCount * 20;
  }

  if (shortenerCount > 0) {
    flags.push(`${shortenerCount} link(s) use URL shorteners — destination hidden`);
    score += shortenerCount * 15;
  }

  if (mismatchCount > 0) {
    flags.push(`${mismatchCount} link(s) show a different URL in text than the actual destination`);
    score += mismatchCount * 25;
  }

  if (httpCount > 0 && email.links.length > 0) {
    flags.push(`${httpCount} link(s) use unencrypted HTTP instead of HTTPS`);
    score += httpCount * 5;
  }

  // ── 5. Attachment analysis ────────────────────────────────────────────────

  for (const att of email.attachments) {
    const ext = att.name.split(".").pop().toLowerCase();
    if (DANGEROUS_EXTENSIONS.has(ext)) {
      flags.push(`Attachment "${att.name}" has a potentially dangerous file type (.${ext})`);
      score += 40;
    }

    // Double extension (e.g., invoice.pdf.exe)
    const parts = att.name.split(".");
    if (parts.length >= 3) {
      const outerExt = parts[parts.length - 1].toLowerCase();
      const innerExt = parts[parts.length - 2].toLowerCase();
      if (DANGEROUS_EXTENSIONS.has(outerExt) && ["pdf","doc","docx","xls","xlsx","jpg","png"].includes(innerExt)) {
        flags.push(`Attachment "${att.name}" uses a double extension to disguise its type`);
        score += 50;
      }
    }
  }

  // ── 6. Safe indicators ───────────────────────────────────────────────────

  if (flags.length === 0) {
    safeOk.push("No spoofing or mismatched sender information detected");
  }
  if (email.links.length > 0 && ipLinkCount === 0 && shortenerCount === 0 && mismatchCount === 0 && httpCount === 0) {
    safeOk.push("No obvious URL hiding or unencrypted links detected");
  }
  if (email.attachments.length === 0) {
    safeOk.push("No attachments");
  }
  if (!URGENCY_PATTERNS.some(p => p.re.test(messageText))) {
    safeOk.push("No urgency or social-engineering language found");
  }

  // ── Score capping + verdict ───────────────────────────────────────────────

  score = Math.min(100, score);

  let verdict, level;
  if (score >= 60) {
    verdict = "Likely Unsafe";
    level   = "danger";
  } else if (score >= 25 || flags.length > 0) {
    verdict = "Exercise Caution";
    level   = "warning";
  } else {
    verdict = "No Obvious Warning Signs";
    level   = "safe";
  }

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

// ── Server-side AI analysis ───────────────────────────────────────────────────

async function callAiProxy(email, ruleResult) {
  const response = await fetch(AI_PROXY_URL, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      email: {
        senderName: email.senderName,
        senderEmail: email.senderEmail,
        subject: email.subject,
        replyTo: email.replyTo,
        bodySnippet: email.bodyText.slice(0, 1200).replace(/\s+/g, " ").trim(),
        links: email.links.slice(0, 10),
        attachmentNames: email.attachments.slice(0, 20).map(a => a.name),
      },
      ruleResult: {
        score: ruleResult.score,
        verdict: ruleResult.verdict,
        flags: ruleResult.flags.slice(0, 20),
      },
    }),
  });

  if (!response.ok) {
    const err = await response.json().catch(() => ({}));
    throw new Error(err.error || `AI service returned error ${response.status}.`);
  }

  return response.json();
}

function reconcileAiVerdict(ruleResult, aiVerdict) {
  const normalized = String(aiVerdict || "").toUpperCase();
  const result = { ...ruleResult, flags: [...ruleResult.flags], safeOk: [...ruleResult.safeOk] };

  // AI can raise risk, but never override concrete rule findings to lower it.
  if (normalized === "UNSAFE" && result.score < 60) {
    result.score = 60;
    result.verdict = "Likely Unsafe";
    result.level = "danger";
    result.flags.push("AI analysis identified additional high-risk context");
  } else if (normalized === "CAUTION" && result.score < 25) {
    result.score = 25;
    result.verdict = "Exercise Caution";
    result.level = "warning";
    result.flags.push("AI analysis identified context that warrants caution");
  }

  return result;
}

// ── UI rendering ──────────────────────────────────────────────────────────────

function showResults(ruleResult, aiText, aiWarning) {
  hideAll();
  el("results-state").classList.remove("hidden");

  const { score, verdict, level, flags, safeOk } = ruleResult;

  // Verdict banner
  const banner = el("verdict-banner");
  banner.className = level;
  el("verdict-icon").textContent = level === "safe" ? "✅" : level === "warning" ? "⚠️" : "🚨";
  const vLabel = el("verdict-label");
  vLabel.textContent = verdict;
  vLabel.className = level;
  el("verdict-summary").textContent =
    level === "safe"    ? "Automated checks found no obvious warning signs, but cannot guarantee this email is safe." :
    level === "warning" ? "Some suspicious patterns detected — proceed carefully." :
                          "Multiple red flags detected — do not click links or reply.";

  // Score bar
  el("score-value").textContent = `${score} / 100`;
  const bar = el("score-bar");
  bar.style.width = `${score}%`;
  bar.className = `score-bar ${level}`;

  // Flags
  const flagsSection = el("flags-section");
  if (flags.length > 0) {
    const list = el("flags-list");
    list.innerHTML = "";
    flags.forEach(f => {
      const li = document.createElement("li");
      li.textContent = f;
      list.appendChild(li);
    });
    flagsSection.classList.remove("hidden");
  } else {
    flagsSection.classList.add("hidden");
  }

  // Safe indicators
  const safeSection = el("safe-section");
  if (safeOk.length > 0) {
    const list = el("safe-list");
    list.innerHTML = "";
    safeOk.forEach(s => {
      const li = document.createElement("li");
      li.textContent = s;
      list.appendChild(li);
    });
    safeSection.classList.remove("hidden");
  }

  // Suggested actions
  const actions = buildSuggestedActions(level, flags, ruleResult);
  const actList = el("actions-list");
  actList.innerHTML = "";
  actions.forEach(a => {
    const card = document.createElement("div");
    card.className = "action-card";
    card.innerHTML = `
      <div class="action-icon">${a.icon}</div>
      <div>
        <div class="action-title">${a.title}</div>
        <div class="action-desc">${a.desc}</div>
      </div>`;
    actList.appendChild(card);
  });

  // AI analysis
  const aiSection = el("ai-section");
  if (aiText) {
    el("ai-text").textContent = aiText;
    aiSection.classList.remove("hidden");
  } else {
    aiSection.classList.add("hidden");
  }

  const warning = el("ai-warning");
  if (aiWarning) {
    warning.textContent = aiWarning;
    warning.classList.remove("hidden");
  } else {
    warning.classList.add("hidden");
  }

  setFooterStatus(`Last checked: ${new Date().toLocaleTimeString()}`);
}

function buildSuggestedActions(level, flags, result) {
  const actions = [];

  if (level === "danger") {
    actions.push({
      icon: "🚫",
      title: "Do not click any links",
      desc: "Links in this email may redirect you to fake login pages or download malware. Hover before you click — if the URL looks suspicious, don't."
    });
    actions.push({
      icon: "🗑️",
      title: "Delete this email",
      desc: "Unless you are certain this email is legitimate (call the sender via a known number to verify), delete it immediately."
    });
    actions.push({
      icon: "📣",
      title: "Report as phishing",
      desc: "In Outlook: right-click → Report → Phishing. This helps protect others in your organization."
    });
    if (result.flags.some(f => f.includes("attachment"))) {
      actions.push({
        icon: "📎",
        title: "Do not open attachments",
        desc: "Attached files could contain malware. Even if it looks like a PDF, it may be disguised."
      });
    }
  } else if (level === "warning") {
    actions.push({
      icon: "🔍",
      title: "Verify the sender independently",
      desc: "If this email appears to be from a company or colleague, confirm by calling them directly using a number from their official website — not from this email."
    });
    actions.push({
      icon: "🖱️",
      title: "Hover over links before clicking",
      desc: "Check that the link destination matches the displayed text. If anything seems off, don't click."
    });
    if (result.flags.some(f => f.includes("credentials") || f.includes("verify"))) {
      actions.push({
        icon: "🔐",
        title: "Never enter credentials via email links",
        desc: "Go directly to the website by typing the URL into your browser — don't use any link in this email to log in."
      });
    }
  } else {
    actions.push({
      icon: "✅",
      title: "Continue with normal caution",
      desc: "Automated checks found no obvious warning signs. Independently verify unexpected requests before replying, opening attachments, or following links."
    });
    actions.push({
      icon: "💡",
      title: "Tip: Stay vigilant",
      desc: "Even safe-looking emails can be sophisticated phishing attempts. Never share passwords or sensitive info via email."
    });
  }

  return actions;
}

// ── Loading / error helpers ───────────────────────────────────────────────────

function showLoading(msg) {
  hideAll();
  el("loading-state").classList.remove("hidden");
  setLoadingMessage(msg);
}

function setLoadingMessage(msg) {
  el("loading-message").textContent = msg;
}

function showError(msg) {
  hideAll();
  el("error-state").classList.remove("hidden");
  el("error-message").textContent = msg;
}

function resetToIdle() {
  hideAll();
  el("idle-state").classList.remove("hidden");
}

function hideAll() {
  ["idle-state","loading-state","results-state","error-state"].forEach(id => {
    el(id).classList.add("hidden");
  });
}

function setFooterStatus(msg) {
  el("footer-status").textContent = msg;
}

// ── Utilities ─────────────────────────────────────────────────────────────────

function el(id) { return document.getElementById(id); }

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

if (typeof module !== "undefined" && module.exports) {
  module.exports = {
    checkLookalikeDomaIn,
    extractDomain,
    extractMailboxAddress,
    extractLinksFromHtml,
    levenshtein,
    reconcileAiVerdict,
    runRuleEngine,
  };
}
