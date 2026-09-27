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

### Login (OAuth2 + PKCE) — experimental

> **Status: experiment, not a replacement.** It is a separate subcommand so the
> working password and manual flows above are never touched.

```bash
microsoft-webauth login-pkce --client-id <your-public-client-id>
```

Opens Microsoft's own sign-in page in your **system** browser. Your password is
typed on Microsoft and is never seen by this tool. A local listener on
`http://localhost:8400/callback` receives the authorization code, which is
exchanged for tokens and written to `auth-file-token.json` with mode `600`.

```bash
# remote/headless box: print the URL instead of launching a browser
microsoft-webauth login-pkce --client-id <id> --no-open-browser

# or read the client id from the environment
export MSOUT_CLIENT_ID=<your-public-client-id>
microsoft-webauth login-pkce
```

| Option | Description |
|--------|-------------|
| `--client-id <id>` | Public client id (or set `MSOUT_CLIENT_ID`) |
| `--tenant <t>` | `common` (default), `organizations`, `consumers`, or a tenant id |
| `--scopes <list>` | Default: `openid profile email User.Read Notes.Read offline_access` |
| `--redirect-uri <uri>` | Default: `http://localhost:8400/callback` (must be registered) |
| `--token-file <path>` | Default: alongside the auth file, `auth-file-token.json` |
| `--login-hint <email>` | Pre-fills the account picker |
| `--timeout <seconds>` | Default: `300` |
| `--no-open-browser` | Print the authorize URL instead of launching a browser |

#### App registration (one-time setup)

1. <https://portal.azure.com> -> **Microsoft Entra ID** -> **App registrations** -> **New registration**
2. Name: anything, e.g. `webauth-pkce-test`
3. **Supported account types:** *Accounts in any organizational directory and personal Microsoft accounts* (this is what makes tenant `common` work)
4. **Redirect URI** platform: **Mobile and desktop applications**, value `http://localhost`
   (Microsoft ignores the port for this platform; `8400` is just where this tool listens)
5. Create, then copy the **Application (client) ID**

No client secret is needed: that is the point of PKCE for a public client.

#### What this does not do

This flow produces **OAuth2/Graph tokens**. It does **not** produce a Playwright
`storageState`, which is what `microsoft-onenote-list-notebooks` and
`microsoft-onenote-export-notebook` require, because those drive the OneNote web
SPA and authenticate with session cookies (`ESTSCS`/`MSPOR`), not bearer tokens.

So after this login you will have a token file, but you will still need
`microsoft-webauth login` to run the export tools. Whether a token can be
exchanged into a usable web session is the open question this experiment exists
to answer.

Treat `*-token.json` as a password: the refresh token can mint access tokens
until it is revoked. The tool writes it `600` and warns if the mode drifts.

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

## Options

| Option | Description |
|--------|-------------|
| `--email <email>` | Microsoft account email (for automated login) |
| `--password <password>` | Microsoft account password (for automated login) |
| `--notheadless` | Run in visible browser mode (disable headless) |
| `--dodump` | Dump HTML content to files for debugging |
| `--auth-file <path>` | Path to auth file (default: ~/.microsoft-webauth/auth-file.json) |

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

MIT — see [LICENSE](LICENSE).

