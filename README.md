# Report This! — Outlook Add-in

An Outlook add-in for **post-click incident reporting**. When someone receives a suspicious email — or realizes they've already clicked a link, entered credentials, or opened an attachment — they open the message, click **Report This!**, answer a few quick questions, and the add-in builds a pre-filled report to the IT support team with the original message and its full headers attached.

This is deliberately **not** a "should I trust this email?" checker. That pre-click question is handled by the separate *Is This Legit* add-in. Report This! answers the other half: *"Something happened — capture the evidence and get it to IT."*

Nothing is sent automatically. The add-in opens a draft addressed to the support mailbox; the user reviews it and clicks **Send**.

## What it does

1. Reads the open message (sender, subject, recipients, links, attachments, body excerpt) and its **full internet headers**.
2. Asks the reporter a short set of triage questions: did you click a link, enter credentials, open an attachment, or reply — and when did it happen.
3. Runs a lightweight rule engine to add **automated flags** (spoofing, lookalike domains, urgency language, risky links/attachments) as extra context for IT.
4. Opens a pre-filled report to **support@mattnj.com** with:
   - the reporter's answers and any notes,
   - full message details and internet headers in the body,
   - the automated flags, and
   - the **original message forwarded as an attachment** for full analysis.
5. Marks the subject with `[ACTION NEEDED]` when the reporter clicked, entered information, or replied — so support can triage exposure quickly.

## Configuration

- **Support mailbox:** set in `taskpane.js` via the `SUPPORT_ADDRESS` constant (currently `support@mattnj.com`).
- **Permission:** `ReadItem`. Both APIs the add-in uses (`getAllInternetHeadersAsync` and `displayNewMessageForm`) require only read-item, so no mailbox-wide consent is needed.
- **Requirement set:** Mailbox **1.8** (needed for `getAllInternetHeadersAsync`). On older clients that don't support full-header reading, the report is still produced with a note in place of the headers.

## Supported Outlook clients

Office web add-in platform, `MessageReadCommandSurface`, Office.js. Designed for Outlook on the web, new Outlook for Windows, and classic Outlook builds that support Mailbox requirement set 1.8. The command appears on the message toolbar, in an overflow menu, or under **Apps**, depending on the client and toolbar customization.

## Fidelity notes

- The original message is attached by its Outlook item id. On the rare client where the item id isn't available, the add-in still sends the full report and tells the user to attach the message manually.
- The rule engine is a heuristic aid for the person triaging the report — it does not decide whether the message is malicious.
- No message content is sent to any outside service. The report stays inside your Microsoft 365 tenant.

## Project structure

```text
manifest.xml       Outlook add-in manifest (Report This!)
taskpane.html      Task pane UI (intake form + confirmation)
taskpane.css       Task pane styles
taskpane.js        Outlook integration, report builder, rule engine
commands.html      Outlook function-file page
server.js          Static development host (no backend API)
support.html       Manifest support page
assets/            Required Outlook icons
test/              Node regression tests
```

## Local development

Requirements: Node.js 20 or later and a Microsoft 365 mailbox supported by Outlook add-ins.

1. Install dependencies.

   ```powershell
   npm install
   ```

2. Start the HTTPS development server.

   ```powershell
   npm start
   ```

3. Sideload `manifest.xml` using Outlook's custom add-in flow or your Microsoft 365 admin deployment process.

4. Open a message and select **Report This!**. If the command isn't directly visible, look under **Apps** or customize the message toolbar.

The development manifest points to `https://localhost:3000`. The first start installs Microsoft's trusted local development certificate.

## Verification

Run all regression tests, JavaScript syntax checks, and Microsoft's manifest validator:

```powershell
npm run check
```

## Production deployment

1. Host the static files on an HTTPS origin.
2. Replace the GitHub Pages URLs in `manifest.xml` with that origin (the pilot is hosted from the `Panic-Button` GitHub Pages repo; rename or re-point as you prefer).
3. Confirm the `SUPPORT_ADDRESS` in `taskpane.js` matches the mailbox that should receive reports.
4. Deploy centrally via the Microsoft 365 admin center (Integrated Apps) so the button appears for all users.
5. Run `npm run check` against the final manifest.

## Relationship to the other add-ins

- **Is This Legit** — *pre-click.* Analyzes an email and advises whether to trust it before you act.
- **Report This!** (this project) — *post-click.* Captures the evidence and reports a suspicious or already-clicked email to IT.

They are complementary and do not overlap: one prevents, the other responds.
