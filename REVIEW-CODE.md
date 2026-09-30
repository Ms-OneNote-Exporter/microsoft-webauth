# Code Review Plan — follow-ups to `dc04e9f`

Review of the work-account / `/copilotnotebooks` fix. The diagnosis in `dc04e9f` is
sound and the "each fix verified independently load-bearing" discipline is the right
standard. This document covers the issues that fix left open.

**Baseline:** `main` @ `2f2ce75`, `src/auth.js`, `src/config.js`, 5 test suites, 38/38
passing, working tree clean.

**Decisions taken (answered during review):**

| Decision | Choice |
| --- | --- |
| `ONENOTE_APP_PATH` scoping | Test `url.pathname` only; no host constraint |
| Verification gate | Add a test-only CI workflow that runs on PRs |
| Email-step wait | Shorten 15 s → 5 s |
| Version bumps | One `0.1.5` bump after all branches merge |
| Delivery | Plan first, then one PR per branch |

---

## Findings

| # | Severity | Finding | Branch |
| --- | --- | --- | --- |
| 1 | **Blocker** | `reachPasswordScreen` still burns the full `stateTimeout` when a visible username field sits next to the password box | 1 |
| 2 | **High** | `ONENOTE_APP_PATH` runs against the whole URL string, not the pathname | 2 |
| 3 | Medium | The `\/notebooks\b` branch has no test pinning it | 3 |
| 4 | Medium | `waitForAuthSuccessProbe` duplicates the Outlook branch of production logic | 4 |
| 5 | Low | `Promise.any` degrades the timeout error to an opaque `AggregateError` | 5 |
| 6 | Low | Email-step `Promise.race` wastes 15 s on non-password next screens | 6 |
| 7 | Low | `ONENOTE_URL` still points at the pre-rebrand `/notebooks` | 7 |
| 8 | Low | Whole regression suite can silently skip when chromium is absent | 8 |
| 9 | Trivial | Orphaned JSDoc; `aria-hidden` compared case-sensitively | 9 |
| 10 | Infra | No test workflow runs on pull requests | 10 |

---

## Build order

Ordered so each branch starts from a settled base and nothing needs rework.

```
10 (CI gate)
 └─> 2 (pathname) ──> 4 (dedupe) ──> 3 (pin tests)
 └─> 1 (stateTimeout)        ↑ shares work-account.test.js with 3, different it() blocks
 └─> 5, 6, 7, 8, 9          (independent, any order)
```

Branches 2 → 4 are sequential on purpose: 4 collapses the duplicated logic onto the
constant that 2 finalises, so doing it the other way round means touching the same
lines twice. Branch 3 goes last among the auth work so its fixtures pin the final
shape rather than an intermediate one.

---

### Branch 1 — `fix/accept-password-field-before-state-timeout`

**Finding 1 (Blocker).** The reorder inside the loop in `dc04e9f` is correct but sits
one layer below the gate that actually blocks. The initial read is:

```js
// src/auth.js:828
let state = await waitForSignInState(page, stateTimeout);
```

`waitForSignInState` defaults to `accept = isActionableSignInState`, which returns
`false` whenever `state.emailField` is true (`src/auth.js:673`) — regardless of a
visible password box. On the "Use a different account" layout where `loginfmt` is
genuinely on screen next to `passwd`, that wait can never be accepted and polls to the
deadline. The loop's new early return at `src/auth.js:840` only fires *after* the full
`stateTimeout` has already elapsed. Net effect: the exact 15 s burn described in the
commit message, deferred by one iteration.

The suite already shows it — two tests, same 3000 ms `stateTimeout`:

```
✓ recognises the password box even though the email field lingers   (183 ms)
✓ reaches the password box when a real username field is on screen (3305 ms)
```

**Change**

1. `src/auth.js:828` — pass an accept predicate that also admits a visible password
   box, mirroring the shape already used for the post-click wait at
   `src/auth.js:886-887`:

   ```js
   let state = await waitForSignInState(page, stateTimeout,
       s => s.passwordField || isActionableSignInState(s));
   ```

2. Leave `src/auth.js:840-847` as `dc04e9f` left it. The loop-level check is still
   required — it is what catches a password box on a state the *initial* read already
   returned, and it keeps the return contract unchanged.

**Safety check to reason through before merging:** on the ordinary email step
(`loginfmt` visible, no password box) the new predicate is `false || false`, so the
wait behaves exactly as today and still returns the last readable state. No regression.

**Tests** — `test/work-account.test.js:163-175` currently passes for the wrong reason
and hides the stall. Add a wall-clock bound so the regression is pinned:

- Assert the visible-username case completes well inside `stateTimeout` (e.g.
  `Date.now() - t0 < stateTimeout`), matching the 183 ms of the aria-hidden case rather
  than the current 3305 ms.
- Keep the existing structural assertions (`emailField === true`, `reached === true`,
  `reason === 'password_field'`) — they are the correct behavioural spec.
- Add a negative case: email step only, no password box, must still return
  `reached: false, reason: 'unreadable'`. This is the case the loosened predicate could
  plausibly break, so it needs pinning.

**Verification** — reintroduce the old single-argument call and confirm the timing
assertion fails. `npm test` → 38 + new cases green.

**Risk** — Low. One call site, additive `||`, negative case covers the failure mode.

---

### Branch 2 — `fix/scope-auth-url-check-to-pathname`

**Finding 2 (High).** The predicate runs against the full URL string, so `/notebooks`
anywhere — including a query parameter, and on any host — satisfies the auth check.
Verified against the current regex at `src/auth.js:132`:

```
true   https://onenote.cloud.microsoft/copilotnotebooks              <- intended
false  https://onenote.cloud.microsoft/en-us                         <- intended
true   https://onenote.cloud.microsoft/en-us?next=/copilotnotebooks  <- should be false
true   https://login.microsoftonline.com/common/oauth2/v2.0/authorize?returnUrl=/notebooks
true   https://evil.example/notebooks
```

The marketing-page guard only holds because those particular URLs happen not to carry
a notebooks path in the query. This is the same failure mode the existing comment warns
about ("a bare hostname check would, and did, causing premature auth saving"), one
level removed. It gates saving live auth state, so the bar is a real invariant rather
than a heuristic.

**Change**

- Both call sites — `src/auth.js:156` (probe) and `src/auth.js:183` (production) — test
  the parsed pathname instead of the serialised URL:

  ```js
  page.waitForURL(url => ONENOTE_APP_PATH.test(url.pathname), { timeout: ... })
  ```

- Update the `ONENOTE_APP_PATH` doc comment to state the invariant precisely: the
  authenticated app is identified by *path*, and the reason a host allowlist was
  rejected (OneNote notebooks are SharePoint-backed; some tenants may not be served
  from `onenote.cloud.microsoft`).
- `waitForURL` hands over a parsed `URL`, so no extra parsing cost.

**Tests** — extend `test/work-account.test.js`:

- Serve the marketing fixture at `https://onenote.cloud.microsoft/en-us?next=/copilotnotebooks`
  → must resolve `false`. This is the case that fails before the fix.
- Serve the marketing fixture at `https://login.microsoftonline.com/...?returnUrl=/notebooks`
  → must resolve `false`.
- Keep the existing `en-us` and bare-login-page negative cases as-is.

**Verification** — revert to `url.toString()` and confirm both new cases fail.
`npm test` green.

**Risk** — Low, and strictly narrower than today. The only behaviour that changes is
matching that was never a real authenticated-app URL.

---

### Branch 3 — `test/pin-both-onenote-app-paths`

**Finding 3 (Medium).** The same gap the `dc04e9f` message describes finding for the
rebrand path still exists for the old path. `test/work-account.test.js:194` serves the
`copilotNotebooks` fixture (`:73`), which contains `<div>All Notebooks</div>` — a
signed-in marker. Confirmed:

```
with /notebooks branch REMOVED, test 194 still passes via: text="All Notebooks"
```

`test:177` has the same weakness. The `bareCopilotShell` case (`:184`) pins
`/copilotnotebooks` only, so **the `\/notebooks\b` branch can be deleted today and the
suite stays green.**

**Change**

- Add an old-path case served the `bareCopilotShell` fixture (`:106`) — no signed-in
  text, so only the URL branch can satisfy it. Mirrors the existing rebrand case exactly.
- Add the same bare-shell treatment for `/copilotnotebooks` so both branches are
  pinned symmetrically and neither relies on a text marker.
- Leave `test:177` and `test:194` in place. They document real rendered-app
  behaviour, which is worth keeping; the bare-shell cases are what carry the URL
  invariant.

**Verification** — delete `\/notebooks\b` from the regex and confirm the suite fails
(this is the check `dc04e9f` ran for the rebrand branch, applied to the old branch).
`npm test` green.

**Risk** — None. Test-only.

---

### Branch 4 — `refactor/dedupe-auth-success-detection`

**Finding 4 (Medium).** The OneNote branch is properly shared — `ONENOTE_APP_PATH` and
`ONENOTE_SIGNED_IN_MARKERS` mean production and test cannot drift there. The Outlook
branch is a verbatim copy-paste: `src/auth.js:151-153` duplicates `src/auth.js:172-176`.
A selector fix landed in production will not be exercised by any test, and the suite
keeps reporting green.

**Change**

- Extract the Outlook attempts into a named constant next to the OneNote ones, so both
  branches are shared data rather than two implementations.
- Make `waitForAuthSuccess` (`src/auth.js:165`) the thin wrapper it should be: delegate
  to the probe with the 60000 ms production timeout, then log. The probe becomes the
  single implementation and the exported test surface (`src/auth.js:1417`) stops being
  a parallel code path.
- Extract the 60000 ms literal to a named constant while here.

**Must not change** — the set of selectors, their timeouts, or the `Promise.any` shape.
This branch is a pure refactor; behaviour differences belong in branch 2, which lands
first.

**Verification** — `npm test` green with no test edits, and the Outlook selectors
appear exactly once. Grep confirms a single definition.

**Risk** — Low. Pure refactor behind a green suite, ordered after branch 2 so the
refactor lands on the final constant.

---

### Branch 5 — `fix/report-auth-success-timeout-clearly`

**Finding 5 (Low).** When every attempt fails, `Promise.any` throws an
`AggregateError`, and the handler at `src/auth.js:1290-1304` rethrows it verbatim. The
user-facing report degrades to "All promises were rejected" while the genuinely useful
signal — the "Still on X — heading" line at `src/auth.js:1295` — comes from a separate
`readScreenState` call. `dc04e9f` widened the aggregate from 3 entries to 4.

**Change**

- Wrap the `Promise.any` in `waitForAuthSuccess` so total failure throws a
  `TimeoutError` naming the target app and the last observed URL, instead of an
  `AggregateError` carrying four opaque `TimeoutError`s.
- Add a debug-level line naming which markers were still unfulfilled at timeout, so a
  future marker rename is diagnosable without a dump.

**Tests** — none strictly required (the existing `resolves.toBe(false)` probe cases
cover the detection side). A unit test asserting the thrown error's type and message
would be reasonable if the probe is made to surface errors; otherwise verify manually.

**Risk** — Low. Error path only, no change to when auth is considered successful.

---

### Branch 6 — `fix/shorten-email-step-wait`

**Finding 6 (Low).** The race at `src/auth.js:1085-1090` waits for either the password
box or `loginfmt` going hidden. Neither fires when the next screen is a method list,
KMSI, or an approval prompt, so those work-account paths still pay the full 15 s. The
`waitForSignInState` poll inside the `reachPasswordScreen` call that follows already
covers all of those screens, which makes this wait largely redundant.

**Change** — reduce both branches from 15000 ms to 5000 ms. Keeps a settle window in
front of the `#usernameError` check at `src/auth.js:1095`, removes ~10 s of dead wait on
the non-password-screen paths.

**Explicitly not doing** — removing the wait entirely. It is redundant in theory, but
this is the live login flow and a settle buffer in front of the error check is cheap
insurance. 5 s is the agreed compromise.

**Verification** — `npm test` green. Confirm by hand against a work account that still
lands on KMSI directly, since no test exercises this timeout.

**Risk** — Low, but this is the one branch that touches the real login path with no test
coverage. Worth a manual run before merge.

---

### Branch 7 — `fix/onenote-entry-url-copilotnotebooks`

**Finding 7 (Low).** `src/config.js:40` still points at the pre-rebrand path:

```js
const ONENOTE_URL = 'https://onenote.cloud.microsoft/notebooks';
```

`dc04e9f` made *success detection* accept both spellings but did not revisit the *entry
point*. Whether this needs changing depends on a fact I could not verify offline:
whether `/notebooks` still redirects to `/copilotnotebooks` or now 404s.

**First step — establish the fact.** Check the captured dumps under
`src/logs/dumps/2026-09-28_00h54` for the final settled URL, then confirm the redirect
behaviour. Note the tests in `work-account.test.js` stub all navigation, so no existing
test can answer this.

**Then, one of:**

- **Redirect confirmed working** → leave `ONENOTE_URL` alone. Add a comment recording
  that both spellings are live and that the entry point is intentionally the stable
  alias. No behaviour change.
- **Redirect broken / 404** → switch to `/copilotnotebooks` and update the three tests
  that pass `https://onenote.cloud.microsoft/notebooks` as `targetUrl`
  (`:180`, `:190`, `:198`, `:207`, `:214`) so they exercise the new default. Those
  `targetUrl` arguments are only used for the Outlook/OneNote branch decision, so the
  change is safe, but the fixtures should be revisited for consistency.

**Risk** — Low either way. The success check accepts both paths, so this cannot cause a
false negative; the only downside of leaving it is a redirect hop on every login.

---

### Branch 8 — `test/fail-loudly-without-chromium`

**Finding 8 (Low).** `test/work-account.test.js:26-38` returns `describe.skip` with only
a `console.warn` when chromium is absent. The entire work-account regression suite then
disappears from the report while the run still passes.

**Change** — replace the silent skip with something visible in the summary: either fail
the run, or surface it as a hard warning plus a non-zero marker. Given branch 10 adds
`npx playwright install chromium` to CI, failing outright is now safe and is the
stronger guarantee.

Keep the existing `try/catch` around `chromium.executablePath()` — that is still needed,
since it throws when the path cannot be resolved at all, and is a different failure from
"resolved but not downloaded".

**Verification** — run the suite with `PLAYWRIGHT_BROWSERS_PATH` pointed at an empty
directory and confirm the intended loud behaviour rather than a green run.

**Risk** — None functionally; it makes an existing silent failure mode visible.

---

### Branch 9 — `chore/auth-small-cleanups`

**Finding 9 (Trivial).** Two zero-risk cleanups, grouped to avoid two PRs of a few lines
each.

1. **Orphaned JSDoc** — `src/auth.js:114-118` was `waitForAuthSuccess`'s doc comment
   until `dc04e9f` inserted the `ONENOTE_APP_PATH` block directly beneath it. It now
   reads as documentation for the regex, and `waitForAuthSuccess` has lost its own.
   Move it back onto the function.
2. **`aria-hidden` case sensitivity** — `src/auth.js:627` compares
   `=== 'true'`. ARIA token values are ASCII case-insensitive, so `"True"` would still
   read as visible. Use
   `(el.getAttribute('aria-hidden') || '').toLowerCase() === 'true'`.

**Verification** — `npm test` green. Optionally add a fixture with `aria-hidden="True"`
to the work-account suite; worth it only if the case is cheap.

**Risk** — None.

---

### Branch 10 — `ci/add-pr-test-workflow`

**Finding 10 (Infra).** `.github/workflows/` contains only `npm-publish.yml`, triggered
on `v*` tags. Nothing runs `npm test` on a pull request, so the branch-per-fix plan
would have no automated gate — and the silent skip in branch 8 would go unnoticed.

**Change** — add `.github/workflows/test.yml`:

- Trigger: `pull_request` on `main`, plus `push` to `main`.
- Steps mirroring the proven parts of `npm-publish.yml` for consistency: `actions/checkout`,
  `actions/setup-node` with `package-manager-cache: false` (the comment in
  `npm-publish.yml` explains the caching choice — follow it), then
  `npx playwright install --with-deps chromium`, then `npm test`.

Two things to copy deliberately from `npm-publish.yml` rather than reinvent:

- No dependency caching, per the existing in-file rationale.
- The tag/version consistency check stays in the publish workflow; it does not belong
  here.

**Verify before merging** — confirm the browser download step is actually required. The
full `playwright` package has a postinstall that downloads browsers, so `npm ci` may
already cover it; if so, keep the explicit `playwright install` anyway for idempotence
and add `--with-deps` for the Linux system libraries. A first run on a real PR is the
only way to confirm, and that is the point of landing this branch first.

**Risk** — None to the package. Adds ~1 min of CI. If it turns out flaky, it blocks
merges, so land it early and alone.

---

## Release

After all ten branches merge:

- Single `chore(release): 0.1.5` bump on `main` — `package.json` version, matching the
  existing pattern.
- Tag `v0.1.5`, which triggers `npm-publish.yml` and runs `npm test` as a release gate.
  That run is also the first time the new PR workflow's browser install is exercised in
  anger, so watch it.

No per-branch version bumps.

---

## Non-goals

Explicitly out of scope, so they do not creep into a branch:

- **Reworking `readSignInState` visibility.** The `aria-hidden` rule is the right
  minimal fix. `Element.checkVisibility()` or `offsetParent` would be more general, but
  the current approach is understood and tested. Not worth the churn.
- **Multi-region / multi-tenant OneNote hosts.** Branch 2 deliberately uses no host
  allowlist for this reason. Revisit only with evidence of a real host.
- **The `Promise.any` dangling-wait pattern in production.** The losing
  `waitForSelector` calls keep running after the race resolves. They are handled
  correctly — `Promise.any` attaches handlers to all, so there is no unhandled
  rejection — and the probe tests confirm it. Cosmetic resource use only; not worth a
  branch now.
- **Outlook branch test coverage.** Branch 4 deduplicates the Outlook logic, but writing
  Outlook fixtures is a genuinely separate piece of work. Worth a follow-up if Outlook
  is a supported target.
- **Existing local branches** (`fix/auth-pro`, `fix/auth-url-and-gathering`,
  `try/oauth2-pkce-flow`). Untouched by this plan; check whether any overlap with
  branches 1–4 before starting, to avoid duplicated effort.
