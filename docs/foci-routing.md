# Throttling and FOCI routing

The plugin shares a scheduler across Graph clients, tools, polling, binary transfers,
and Microsoft internal service calls. `Retry-After` (seconds or HTTP date),
`x-ms-retry-after-ms`, and `retry-after-ms` are minimum delays. The longest valid
hint wins. Missing hints use exponential backoff with jitter. A single subrequest
gets at most four service attempts and one client switch.

Settings:

- **Maximum throttle wait (seconds)** defaults to 30, accepts integers from 0 to
  300, and applies to new logical actions. Zero returns immediately when waiting
  is necessary. Expiration never leaves a user write queued for later execution.
- **Adaptive FOCI routing** defaults to enabled. Disabling it preserves current
  preferred clients and coordinated retries. Scope-qualified authorization
  selection for Teams operations remains available.

An alternate needs an account-matching Graph token, a verified client/audience,
and an explicit acceptable delegated scope set for the operation. Client family
membership and public permission catalogs are candidate hints, not authorization.
Candidate redemption is silent, serialized, and paced to at most one per second.
Failed candidates never initiate consent/MFA or reset a healthy primary session.

Application-scoped throttles can select an eligible client whose admission time
is earlier. Tenant/user/resource limits remain shared. An unknown-scope Graph
read allows one alternate probe per operation/resource per five minutes. Success
temporarily favors the alternate until the original cooldown expires; correlated
throttling applies a shared cooldown. Unknown-scope writes never probe.

Outlook has four simultaneous requests per app/mailbox, with ordinary background
work limited to three. Batches reserve all four slots; only unsuccessful items
are retried. Chat listing is paced per user and message listing per conversation.
Background pagination retains its cursor when it is deferred.

OneDrive/SharePoint, preauthorized upload URLs, and undocumented Teams/OWA APIs
use backoff without throughput rotation. Ambiguous writes are not blindly
replayed. Uploads inspect session progress after an ambiguous chunk failure and
advance only when the service confirms progress; unresolved chunks require an
explicit retry. Partial bulk operations report actual successful item counts.

## Verification

Run `npm test`, `npm run typecheck`, and `npm run build`. Tests simulate throttles
using fake time and verify concurrency, token qualification, routing, cancellation,
partial batches, upload reconciliation, and settings. The normal build writes
`dist/`; it does not install or reload the running plugin.

For read-only live validation using an existing, single-account MSAL family-token
cache:

```sh
node scripts/check-foci.mjs /absolute/path/to/token_cache.json
```

The script uses the implementation's token acquisition and scope checks, issues
only small reads (including the read-only presence POST), and reports endpoint
confirmation separately from token-qualified write capabilities. It never tests
a write, persists credentials, or prints account identifiers/tokens. A revoked
session or throttling failure stops further candidate checks.

Logs include workload, client alias, throttle scope/reason, delay, attempt, and
request correlation ID. They do not include authorization headers or request
content. Candidate count is never evidence of extra throughput.

The implementation-time live check on 2026-10-19 reached Microsoft but the existing
local MSAL cache returned `invalid_grant`. Other candidates were skipped; live
qualification and throughput improvement remain unverified until a fresh cache
is supplied through the normal sign-in workflow.
