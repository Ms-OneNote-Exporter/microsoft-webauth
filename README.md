# microsoft-webauth

A tool to authenticate against microsoft online (live or professionnal tenants)
We do not rely on GraphAPI.

e.g:
  - https://onenote.cloud.microsoft/notebooks
  - https://outlook.live.com/mail/
  - and by extension https://login.microsoft.com

Microsoft web authentication via Playwright — extracted from [MSOneNote Exporter](https://github.com/enoola/Microsoft-OneNote-Exporter).
I did extract it realising it might be useful out of the box for other projects.
And it will make it easier to maintain as far as I see it.

This is a standalone CLI tool for authenticating with Microsoft accounts using Playwright. It handles:
- NO GraphAPI
- Automated login with email/password
- Manual login in browser
- MFA/2FA support (OTC codes, number matching)
- Session persistence
- The interstitial screens Microsoft injects mid-login (see below)

## Interstitial screens

Microsoft interrupts an otherwise successful login with full-page forms that
take over the navigation. If they are not answered, the login silently hangs and
ends in a timeout. These are handled automatically:

| Screen | Action taken |
|--------|--------------|
| "We're updating our terms" (`account.live.com/tou/accrue`) | Next — accepts the updated Services Agreement |
| "Is your security info still accurate?" (`account.live.com/interrupt/…`, `/proofs/remind`) | Looks good! — keeps existing recovery methods |
| Passkey / security key prompt (`…/consumers/fido/create`) | Cancel |
| "Stay signed in?" | Yes, with "don't show again" ticked |
| Microsoft consent pages (`consent.microsoft.com`) | Accept / Continue |

Each screen only accepts a fixed set of button labels, so nothing else on the
page can be pressed by accident. In particular the tool never chooses "Update
now" or "I don't have any of these" on the security-info screen, since both would
change or delete the account's recovery methods.

Note that accepting the Services Agreement is a real change to the account, and
is done on your behalf. If you would rather see it, run without `--email` and
`--password` and sign in manually.

## Why this project ?

While this let you authenticate this is a part of a bigger purpose,
primary aim is to offer people a simple way to get out of Microsoft OneNote, because you regardless of what ms documentation states
=> https://learn.microsoft.com/en-us/answers/questions/2276682/onenote-api-fails-with-large-sharepoint-document-l

in essence you want to search for microsoft-onenote-list-notebook, microsoft-onenote-exporter

## Available on npmjs

You can find this package here: https://www.npmjs.com/package/@msout/microsoft-webauth

## Installation

```bash
npm install -g @msout/microsoft-webauth
```

Or locally:

```bash
npm install @msout/microsoft-webauth
```

## Usage

### Login (Automated)

```bash
microsoft-webauth login --email your@email.com --password yourpassword
```

With custom auth file path:

```bash
microsoft-webauth login --email your@email.com --password yourpassword --auth-file /path/to/authfile.json
```

### Login (Manual/Interactive)

```bash
microsoft-webauth login
```

This will open a browser window. Log in manually, then press Enter when you see the notebooks list.

With custom auth file path:

```bash
microsoft-webauth login --auth-file /path/to/authfile.json
```

### Check Authentication Status

```bash
microsoft-webauth check
```

With custom auth file path:

```bash
microsoft-webauth check --auth-file /path/to/authfile.json
```

### Logout

```bash
microsoft-webauth logout
```

With custom auth file path:

```bash
microsoft-webauth logout --auth-file /path/to/authfile.json
```

## Exit codes

`login` and `check` exit `0` only when they actually succeeded, so a shell script
or a CI step can branch on them:

```bash
microsoft-webauth login --email you@example.com --password ... \
  && microsoft-webauth export   # only runs if the login worked
```

| Command | Exit 0 | Exit 1 |
|---------|--------|--------|
| `login` | the session authenticated **and** the auth file was written and read back | the login failed, or it reached the app but the auth file is missing, unparseable or not a Playwright storage state |
| `check` | the saved session was confirmed live | no auth file, the session has expired, or it could not be verified |
| `logout` | always | — |

Two details worth knowing:

- A login that reaches the authenticated app but leaves nothing usable on disk
  exits `1`. The auth file is written and then read back, because a write that
  returned without throwing is not evidence that a file exists.
- `check` exits `1` when it cannot confirm a session — including when the check
  itself failed on a network error. The auth file is still left in place in that
  case: the check failing is not evidence that the session is bad. This is why
  `check` waits for Microsoft to either open the app or redirect to a login page,
  rather than deciding after a fixed pause, and why a session that stays on the
  unauthenticated page is reported as expired instead of working.

## Options

| Option | Description |
|--------|-------------|
| `--email <email>` | Microsoft account email (for automated login) |
| `--password <password>` | Microsoft account password (for automated login) |
| `--notheadless` | Run in visible browser mode (disable headless) |
| `--dodump` | Dump HTML content to files for debugging |
| `--screenshot` | With `--dodump`, also save a PNG screenshot of each dumped page |
| `--auth-file <path>` | Path to auth file (default: ~/.microsoft-webauth/auth-file.json) |

## Debug dumps and screenshots

```bash
microsoft-webauth login --email you@example.com --password ... --dodump --screenshot
```

With `--dodump`, every screen state captured during an automated login is
written to `src/logs/dumps/<YYYY-MM-DD_HHhMM>/` as HTML — after the email step,
after the password, on each interstitial screen, and on the failure paths.
Adding `--screenshot` saves a full-page PNG of the same screen under the same
basename, so `debug_after_email.html` is accompanied by
`debug_after_email.png`.

Interstitial screens are captured wherever they turn up, including the ones that
arrive *late* — Microsoft often serves them 20–60 s into a login, behind a "Stay
signed in?" prompt. Those are written as `debug_late_blocking_screen_N.html`,
kept separate from the `debug_blocking_screen_N.html` of the earlier pass so the
two cannot overwrite each other.

The HTML says *which* screen this was; the screenshot says what it *looked* like,
which is what answers the questions a dump usually gets asked — was something
covering the button, was there a banner or an overlay, did the page render at all.
A capture that cannot be taken (the page navigated away mid-capture) is reported
as a warning and never fails the login.

Two things to know before sharing a dump directory:

- HTML dumps have credential fields redacted; screenshots cannot be, being
  bitmaps. That costs nothing in credential terms — a password field renders as
  dots — but a screenshot does show the number-match MFA code, exactly as the
  HTML dump and the terminal log already do.
- Both land in the same gitignored directory. Treat them as you would the dumps:
  worth attaching to a bug report, not worth pasting into a public channel
  unedited.

## Output

Authentication state is saved to the specified auth file path. By default, it uses `~/.microsoft-webauth/auth-file.json`. A metadata file `{auth-file-prefix}-meta.json` stores:
- Email used for login
- Login timestamp

When files already exist, they are automatically backed up with `.old` suffix. If `.old` files already exist, a warning is displayed before overwriting.

## Testing

```bash
npm test
```

Run with coverage:

```bash
npm run test:coverage
```

## Project Structure

```
microsoft-webauth-playwright/
├── src/
│   ├── auth.js          # Authentication logic
│   ├── config.js        # Configuration (paths, URLs)
│   └── utils/
│       ├── logger.js    # Logging utilities
│       └── retry.js     # Retry helpers
├── test/                # Jest tests
├── package.json
└── README.md
```

## License

MIT — see [LICENSE](LICENSE), and read [NOTICE.md](NOTICE.md).

