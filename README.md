# microsoft-webauth

A tool to authenticate against microsoft online (live or professionnal tenants)
e.g: 
  - https://onenote.cloud.microsoft/notebooks
  - https://outlook.live.com/mail/
  - and by extension https://login.microsoft.com

Microsoft web authentication via Playwright — extracted from [MSOneNote Exporter](https://github.com/msout/Microsoft-OneNote-Exporter).

This is a standalone CLI tool for authenticating with Microsoft accounts using Playwright. It handles:
- Automated login with email/password
- Manual login in browser
- MFA/2FA support (OTC codes, number matching)
- Session persistence
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

### Login (Manual/Interactive)

```bash
microsoft-webauth login
```

This will open a browser window. Log in manually, then press Enter when you see the notebooks list.

### Check Authentication Status

```bash
microsoft-webauth check
```

### Logout

```bash
microsoft-webauth logout
```

## Options

| Option | Description |
|--------|-------------|
| `--email <email>` | Microsoft account email (for automated login) |
| `--password <password>` | Microsoft account password (for automated login) |
| `--notheadless` | Run in visible browser mode (disable headless) |
| `--dodump` | Dump HTML content to files for debugging |

## Output

Authentication state is saved to `auth.json` in the project directory. A metadata file `auth-meta.json` stores:
- Email used for login
- Login timestamp

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

ISC — same as MSOneNote Exporter.
