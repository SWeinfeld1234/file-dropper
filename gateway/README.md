# Refuah Attach Gateway (Azure Function)

Single-purpose HTTP gateway that lets Outlook (including classic **Outlook 2019**)
attach a patient file from the firewalled `refuahfiles` Azure Files share, without
ever opening the storage account to the public internet.

See the full design in `../docs/azure-gateway-spec.md` and the build steps in
`../docs/azure-gateway-implementation-steps.md`.

## What it does

1. Salesforce mints a short-lived, **AES-encrypted + HMAC-signed** token scoped to
   one file (and only after checking the user's access — see the Apex
   `generateGatewayUrl`).
2. The add-in hands Exchange a URL: `…/api/attach?token=<enc.mac>&code=<funcKey>&fn=<file>`.
3. Exchange fetches it **server-side**. This Function MAC-verifies, decrypts,
   checks expiry + one-time nonce, then streams the one named file from
   `refuahfiles` via **Managed Identity over a Private Endpoint**.

The storage account public access stays **Disabled**. The Function is the only
thing Exchange can reach; the Function reaches storage only over the private path.

## Infra (provisioned by Positive MSP — do not re-create)

| Item | Value |
|---|---|
| Function App | `func-attach-gateway-766` (Flex Consumption, Node 20) |
| Base URL | `https://func-attach-gateway-766.azurewebsites.net/api/attach` |
| Storage account | `refuahfiles` |
| File share | `records` |
| Key Vault secret | `kv-attachgw-595` → `gateway-hmac-secret` |
| RBAC | Function identity = read-only on the single `records` share |

App settings (`STORAGE_ACCOUNT`, `MAX_FILE_MB`, `ENFORCE_ONE_TIME`,
`GATEWAY_HMAC_SECRET` as a Key Vault reference, `APPLICATIONINSIGHTS_CONNECTION_STRING`)
are already set. The **function key** and **shared secret** are delivered over a
secure channel — never commit them.

## Security punch-list items addressed in code

- **#1 — no PHI in logs.** The catch block logs only `e.code` / `e.statusCode`,
  never `e.message` (the SDK message can embed the request URL = patient path).
  The happy-path log records nonce + bytes only — never `path` or `fn`.
- **#5 — zero-downtime secret rotation.** `verifyToken` accepts the current
  secret *and* an optional `GATEWAY_HMAC_SECRET_PREV`, so tokens minted with the
  old secret keep working during a rotation overlap window. Remove
  `GATEWAY_HMAC_SECRET_PREV` once the overlap (≥ token TTL) has passed.
- Encrypt-then-MAC, MAC verified before decrypt; `Content-Disposition` filename
  stripped of quotes **and** CR/LF (header-injection safe).

## Deploy

```bash
cd gateway
npm install
func azure functionapp publish func-attach-gateway-766
```

## Smoke test

Mint a token (Salesforce anonymous Apex `generateGatewayUrl`, or a small local
script using the shared secret) then:

```bash
curl -v "https://func-attach-gateway-766.azurewebsites.net/api/attach?token=<token>&code=<funcKey>" -o out.bin
```

Expect: `200` + `Content-Disposition: attachment` for a valid token; `401`
(tampered), `410` (expired / already used), `404` (missing file). Confirm
Application Insights shows **no path/filename** in any log line.
