# Email Safety Checker — Outlook Add-in

An Outlook web add-in that checks an open message for common phishing, spoofing, and social-engineering warning signs. The current GitHub Pages pilot runs the local rule engine only; message content is not sent to an AI service.

The result is guidance, not a guarantee that a message is safe. The add-in does not scan attachment contents, resolve redirect chains, consult threat-intelligence feeds, or independently validate SPF, DKIM, and DMARC.

## Supported Outlook clients

The project uses the Office web add-in platform, `MessageReadCommandSurface`, and Office.js. It is designed for Outlook on the web, new Outlook for Windows, and supported classic Outlook versions. The command can appear directly on the message toolbar, in an overflow menu, or under **Apps**, depending on the client and the user's toolbar customization.

## Checks performed

- Display-name and sender-domain mismatch
- Lookalike domains for a limited trusted-domain list
- Business-like senders using free email providers
- Reply-To domain mismatch
- Urgency, credential, prize, and threat language in the subject and body
- Requests for sensitive information
- IP-address links, URL shorteners, misleading link text, and plain HTTP
- Dangerous attachment names and double extensions
- Optional AI review of a limited excerpt, links, filenames, and rule findings

## Security design

- The Claude API key is read only by `server.js` from `ANTHROPIC_API_KEY`.
- No provider key is stored in Outlook, roaming settings, browser storage, or task-pane JavaScript.
- The browser sends at most 1,200 body characters, 10 links, and 20 attachment filenames to the local proxy.
- The server applies request-size and per-address rate limits.
- If AI is unavailable, the pane visibly says that the result uses rules only.
- AI may raise a risk classification, but it cannot erase concrete rule findings.

For a public production deployment, protect `/api/analyze` with your organization's authentication and edge rate limiting. Origin checks alone are not sufficient protection for an API that can spend money.

## Project structure

```text
manifest.xml       Outlook add-in manifest
taskpane.html      Task pane UI
taskpane.css       Task pane styles
taskpane.js        Outlook integration and rule engine
commands.html      Outlook function-file page
server.js          Static development host and server-side AI proxy
support.html       Manifest support page
assets/            Required Outlook icons
test/              Node regression tests
```

## GitHub Pages pilot

The pilot is hosted at `https://cornercoderman.github.io/Panic-Button/`. The checked-in manifest points to that site and is ready for a rules-only Microsoft 365 test after the current project files are published to the repository's Pages branch.

GitHub Pages serves static files and does not run `server.js`. AI analysis therefore remains disabled in `taskpane.js` for this pilot.

## Local development

Requirements: Node.js 20 or later and a Microsoft 365 mailbox supported by Outlook add-ins.

1. Install dependencies.

   ```powershell
   npm install
   ```

2. Optionally enable AI analysis for this terminal session. Do not paste the key into source files.

   ```powershell
   $env:ANTHROPIC_API_KEY = "your-key-from-your-password-manager"
   ```

3. Start the HTTPS development server.

   ```powershell
   npm start
   ```

4. Sideload `manifest.xml` using Outlook's custom add-in flow or your Microsoft 365 admin deployment process.

5. Open a message and select **Check Safety**. If the command isn't directly visible, look under **Apps** or customize the message toolbar.

The development manifest points to `https://localhost:3000`. The first start installs Microsoft's trusted local development certificate.

## Verification

Run all regression tests, JavaScript syntax checks, and Microsoft's manifest validator:

```powershell
npm run check
```

## Production deployment with optional AI

Before centralized deployment or Marketplace submission:

1. Host the static files and proxy on an HTTPS origin.
2. Replace the GitHub Pages URLs in `manifest.xml` with that origin.
3. Change `ProviderName` and the support-page content to your organization.
4. Set `ANTHROPIC_API_KEY` in the hosting platform's secret manager, if AI is enabled.
5. Add organization authentication and production rate limiting to `/api/analyze`.
6. Review disclosure, retention, and approved-provider requirements before sending mailbox excerpts to an AI service.
7. Run `npm run check` against the final manifest.

If AI isn't approved, leave `ANTHROPIC_API_KEY` unset. The add-in will continue using its rule-based checks and will tell the user that AI analysis isn't configured.
