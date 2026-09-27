# OAuth2 + PKCE experiment — findings

**Branch:** `try/oauth2-pkce-flow`
**Status:** complete. Result is **negative for the hosted-service use case**, and positive for the local CLI.

## Question

Can a Microsoft OAuth2 token replace the Playwright cookie session?

This matters because the planned hosted service must authenticate a user whose
Microsoft password the service must never handle. If a token can stand in for a
web session cookie, the service can run fully unattended. If not, the design is
forced onto "the user supplies their own `auth.json`".

## What worked

`microsoft-webauth login-pkce` performs a real authorization-code + PKCE flow
against a single-tenant (`consumers`) app registration. Verified end to end:

- consent granted, `User.Read` + `Notes.Read` granted
- code received, exchanged, tokens written `mode 600`
- refresh token issued (417 chars), access token valid 3599s
- `GET https://graph.microsoft.com/v1.0/me` → **HTTP 200**, correct account

The password was entered on Microsoft's own page in the system browser. No
Playwright, no credential handling, no MFA interception. For a user running the
CLI locally, this is strictly better than the `--email/--password` mode.

## What did not work

**A Graph token cannot bootstrap a OneNote web session.** Evidence, in order of
decisiveness:

1. **Fresh Playwright context, no cookies, navigating to
   `onenote.cloud.microsoft/notebooks`** lands on the public marketing page with
   a *Sign in* button. Only 3 cookies are set — `msal.cache.encryption`,
   `HP.SID`, `HP.AR` — all anonymous. No `ESTSCS`, no `MSPOR`.

2. **No first-party API accepts the bearer token.** Probing plausible OneNote
   web endpoints returned `200 text/html` — but so did a deliberately
   nonexistent route (`/apis/v1.0/this-route-does-not-exist-xyz`), byte for
   byte identical. These are single-page-app catch-all shells, not APIs. A
   status code alone would have been a false positive here.

3. **The Graph OneNote API is unavailable for this account type.**
   `GET /v1.0/me/notebooks` → **HTTP 404 `UnknownError`**, while `/me` → 200.
   This independently confirms the project premise that Graph is a dead end for
   OneNote export, beyond the documented 50-page ceiling.

The access token is opaque (1464 chars, not a JWT), so its audience cannot be
inspected client-side. That was confirmed empirically rather than assumed.

## Why the two mechanisms cannot be bridged

The OneNote web SPA authenticates with `ESTSCS`/`MSPOR` session cookies minted
by its own handshake. An OAuth code redemption yields a Graph bearer token. They
are issued by different subsystems and there is no supported conversion between
them.

The one variant that *would* produce a usable `storageState` is running the
PKCE flow **inside** Playwright, so the automated browser performs the sign-in
itself and its cookies are captured. That works — but it requires a visible
browser on the server, i.e. streaming it to the user, i.e. exactly the approach
already rejected on trust and abuse-detection grounds. It is not a way out.

Incidentally, this is why the `some-prompts.md` `#newSessionLink` interstitial
matters: it is the OAuth authorize endpoint appearing *because* web cookies alone
were insufficient. The two mechanisms are entangled, as suspected.

## Consequences for the hosted service

| | |
|---|---|
| Browser streaming (noVNC) | **Still ruled out.** PKCE does not remove the need for it. |
| Accepting Microsoft passwords in our web UI | **Ruled out**, and now also pointless: it would not have made the service work unattended. |
| **User supplies their own `auth.json`** | **The design.** No password, no consent screen, no browser streaming, and the user already runs the CLI to produce the file. |
| Professional tenants | Unchanged. The Playwright flow needs no app registration and no consent, so the verified-publisher and admin-consent walls never apply. |

The unverified-publisher banner on the consent screen is a separate, independent
reason to avoid a shared client id: every user would be asked to grant
"onenote-exporter" read access to all their notebooks, from a publisher nobody
can verify. That is a poor trust signal for a tool whose entire pitch is *get
your data out of Microsoft safely*.

## Keep or drop `login-pkce`?

Worth keeping, for the local CLI only:

- the password never reaches the tool, which is a real improvement
- no MFA interception
- no 300s interstitial-scripting against Microsoft's login pages, so far more
  robust against Microsoft's UI changes than the Playwright flow

It should not become the hosted service's auth mechanism.

## Reproducing

```bash
# 1. one-time app registration: see README "App registration (one-time setup)"
node src/index.js login-pkce --client-id <id>

# 2. token is a Graph token
node -e 'const t=require(process.env.HOME+"/.microsoft-webauth/auth-file-token.json"); \
  fetch("https://graph.microsoft.com/v1.0/me",{headers:{Authorization:"Bearer "+t.access_token}})\
  .then(r=>r.text()).then(console.log)'

# 3. but it does not open a OneNote web session: load onenote.cloud.microsoft
#    in a fresh Playwright context and you get the sign-in page, not notebooks.
```
