# Azure File Attachment Gateway — Specification & Technical Build-Out

**Project:** Refuah Helpline — Outlook "Attach Patient File from Salesforce" add-in
**Component:** `refuah-attach-gateway` (Azure Function)
**Owner:** Logicfold (Shloime Weinfeld)
**Status:** Design — client (Refuah) approved in principle 2026-06-08; pending Positive MSP security sign-off
**Audience:** Refuah / Positive MSP security & Azure administrators, Logicfold engineering

---

## 1. Problem statement

Staff need to attach patient files (stored in Azure Files, account `refuahfiles`) directly into an Outlook email from within Salesforce. The Salesforce LWC + Outlook add-in already work end-to-end on **Outlook on the Web / New Outlook** using a client-side base64 path. They **fail on classic Outlook 2019**, which the majority of staff use.

### 1.1 Why Outlook 2019 fails today

Classic Outlook 2019 runs add-ins in an **IE11 / Trident webview (requirement set 1.6)**. It does **not** support `addFileAttachmentFromBase64Async` (requirement set 1.8). The only attach API available is:

```js
Office.context.mailbox.item.addFileAttachmentAsync(url, name, options, cb)  // requirement set 1.1
```

With this API, **Exchange Online's mail servers fetch the URL server-side** — not the user's PC, not Salesforce. The fetch therefore originates from Microsoft 365's network, carrying **no Salesforce session and no user credentials**.

The `refuahfiles` storage account is (correctly) firewalled: **public network access is denied; only Salesforce is permitted**. When Exchange tries to fetch the SAS URL, the storage firewall blocks it. Result: the attachment silently fails or arrives empty.

### 1.2 Why we can't simply "open the firewall to Outlook"

- Exchange Online's outbound fetch comes from a **large, shared, changing range of Microsoft IPs** — allow-listing them is impractical and would effectively expose the account.
- **Azure Storage firewall IP rules are ignored for requests that originate inside Azure.** Exchange Online runs inside Azure, so IP allow-listing the storage account would not even take effect reliably.
- Loosening the storage firewall weakens the PHI security posture we deliberately established.

We need a **narrow, audited, single-purpose broker** that Exchange *can* reach, which itself reaches storage over a trusted private path.

---

## 2. Solution overview

Introduce a single-purpose **Azure Function** (`refuah-attach-gateway`) that acts as a controlled gateway:

```
                         (1) Salesforce mints a short-lived,
                             HMAC-signed token for ONE file
                                       │
                                       ▼
┌───────────────┐   drag SAS-style    ┌──────────────────┐
│ Salesforce    │   gateway URL as    │ Outlook add-in   │
│ LWC (browser) │ ─ text/plain ─────▶ │ taskpane.js      │
└───────────────┘                     └────────┬─────────┘
                                               │ addFileAttachmentAsync(gatewayUrl)
                                               ▼
                                      ┌──────────────────────┐
                                      │ Exchange Online       │  (2) server-side GET
                                      │ (fetches the URL)     │ ───────────────┐
                                      └──────────────────────┘                 │
                                                                               ▼
                                              PUBLIC HTTPS (the only ingress)  ┌──────────────────────┐
                                              ───────────────────────────────▶│ Azure Function        │
                                                                               │ refuah-attach-gateway │
                                                                               │  • validate HMAC      │
                                                                               │  • check expiry/nonce │
                                                                               │  • stream ONE file    │
                                                                               └───────────┬──────────┘
                                                                                           │ (3) Managed Identity,
                                                                                           │     over Private Endpoint
                                                                                           ▼
                                                                               ┌──────────────────────┐
                                                                               │ refuahfiles           │
                                                                               │ (Azure Files)         │
                                                                               │ PUBLIC ACCESS: DENIED │
                                                                               └──────────────────────┘
```

**In one line:** we are *not* exposing the file storage. We issue **disposable, expiring, single-file passes** — no more sensitive than the email the file is being attached to — while the underlying storage stays fully firewalled and private.

---

## 3. Security design (the part the security team cares about)

| Concern | How it's handled |
|---|---|
| **Storage stays private** | `refuahfiles` public network access remains **Disabled**. The Function reaches it only via a **Private Endpoint** on the Microsoft backbone. The public internet is still denied. |
| **No standing credential is exposed** | The only thing that leaves the building is a **single-use, time-boxed, encrypted-and-signed token** in a URL. It contains no readable patient data, no account key, no Salesforce session. |
| **Token reveals no PHI** | The token payload is **AES-256 encrypted** (not merely signed), so the file path — which contains the patient name + DOB — is **not decodable** by anyone holding the URL. This matters because the URL lands in Exchange mail logs and (potentially) telemetry. See §5. |
| **Minting is authorization-gated** | Salesforce will only mint a token for a file the **running user is allowed to see**. `generateGatewayUrl` takes a Case Id, enforces record-level sharing via SOQL in a `with sharing` context, and confirms the requested path sits **within that Case's patient folder** before signing. A signed token is therefore proof of *authorized* access, not just a valid signature. See §8.2. |
| **Token is scoped to exactly one file** | The token's encrypted payload names one file path. The Function will stream *that file and nothing else*. Path traversal (`..`, absolute paths) is rejected. |
| **Token expires fast** | Default lifetime **5 minutes** (configurable). After expiry the token is dead. |
| **Token is (optionally) one-time** | A used token's nonce is recorded in Azure Table Storage; replays are rejected (with a short grace window — see §6.4). |
| **Function → storage uses no key** | The Function authenticates to Azure Files with its **system-assigned Managed Identity** + Azure RBAC (`Storage File Data Privileged Reader`). No account key is stored in the Function. |
| **Logs carry no PHI** | Every request is written to Application Insights for audit, but the log line records only the **nonce, byte count, caller IP, and outcome** — **never the file path** (the path contains patient identifiers). See §6.2. |
| **Stays in HIPAA scope** | Azure Functions, Azure Storage, Key Vault, Private Endpoints, and Application Insights are all covered by the **same Microsoft Business Associate Agreement (BAA)** already in force for the storage account. |
| **Defense in depth on ingress** | In addition to the HMAC token, the Function endpoint requires a Function-level access key, and CORS is locked to no browser origins (server-to-server only). |

### 3.1 What the security team is being asked to accept

1. One new **Azure Function App** in the existing `refuahfiles` subscription.
2. One new **Private Endpoint** on the storage account (`file` sub-resource) — this is the *standard* Microsoft-recommended way to grant private access and is generally viewed favorably in a security review.
3. One **RBAC role assignment** granting the Function's managed identity read-only file access, **scoped to the single share that holds patient files** (not the whole storage account — see §4 note and implementation step A5).
4. One **shared secret** (used for both AES encryption and HMAC) created in Key Vault and shared with Salesforce as a protected setting, **with a documented rotation procedure** (see §13).

There is **no** change that makes the storage account publicly reachable.

> **Two findings must land before production cutover** (per Positive MSP security review): the token must be **encrypted, not just signed** (§5), and minting must be **authorization-gated in Salesforce** (§8.2). Both are reflected throughout this document.

---

## 4. Component inventory

| # | Resource | Type | Purpose |
|---|---|---|---|
| 1 | `refuah-attach-gateway` | Function App (Flex Consumption or Elastic Premium) | The gateway code |
| 2 | `func-vnet` / delegated subnet | Virtual Network + subnet | Regional VNet integration for the Function's outbound traffic |
| 3 | `pe-refuahfiles-file` | Private Endpoint (sub-resource `file`) | Private path from the VNet to `refuahfiles` |
| 4 | `privatelink.file.core.windows.net` | Private DNS Zone | Resolves `refuahfiles.file.core.windows.net` to the private IP |
| 5 | System-assigned Managed Identity on (1) | Identity | Keyless auth to Azure Files |
| 6 | Role assignment: **Storage File Data Privileged Reader** | RBAC | Lets the identity read file data. **Scope to the single share** (`.../fileServices/default/shares/<share>`), not the account, so the gateway can never read any other share on `refuahfiles`. |
| 7 | `kv-refuah-gateway` (or existing KV) | Key Vault | Holds the HMAC shared secret |
| 8 | `attach-gateway-ai` | Application Insights | Audit & diagnostics |
| 9 | `gatewaytokens` table | Azure Table Storage | One-time-use nonce ledger (optional but recommended) |

> **Plan choice:** Use **Flex Consumption** (supports VNet integration, scales to zero, lowest cost) if available in the region; otherwise **Elastic Premium EP1**. The classic Consumption (Y1) plan does **not** support VNet integration and is therefore unsuitable.

---

## 5. Token design

The token is what Salesforce mints and the Function validates. It is a compact, URL-safe, **encrypted-then-MAC'd** structure.

> **Why encrypted, not just signed (security finding #1).** An HMAC-signed-but-plaintext token has its payload base64-encoded, not encrypted — anyone holding the URL can base64-decode the first segment and read the file path, which contains the **patient name + DOB**. The URL is fetched by Exchange (lands in mail logs) and could surface in telemetry. We therefore **encrypt the payload (AES-256-CBC)** so the path is opaque, and **keep an HMAC over the ciphertext** (encrypt-then-MAC) for integrity/authenticity. Tampering fails the MAC; eavesdropping reveals nothing.

### 5.1 Format

```
ct    = AES-256-CBC( minified-JSON-payload )      // 16-byte random IV prepended to ciphertext
seg1  = base64url( IV || ct )
token = seg1 + "." + base64url( HMAC_SHA256( seg1, macKey ) )
```

Two keys are derived from the single Key Vault secret so encryption and authentication keys are independent:
`encKey = SHA256(secret || "enc")`, `macKey = SHA256(secret || "mac")`. (Or store two distinct secrets; one secret + domain separation is simpler to rotate.)

`payload` is minified JSON (now confidential — never visible in the URL):

```json
{
  "v":  1,                          // schema version
  "p":  "<share-relative file path>",  // contains patient identifiers — THIS IS WHY WE ENCRYPT
  "sh": "<share name>",
  "n":  "b7c1e0d2f9a84e3b",         // nonce (16 random hex chars)
  "exp": 1781035200,                 // expiry, epoch seconds (UTC)
  "fn": "cbc.pdf"                    // suggested download filename — BARE filename only, no path/identifiers
}
```

> `fn` is the only field that becomes externally visible (as the email's attachment name via `Content-Disposition`). It **must** be the bare filename (`cbc.pdf`), never the patient-bearing folder path. Salesforce sets it with `substringAfterLast('/')`.

### 5.2 Validation rules (Function side)

1. Split on `.`; recompute HMAC over **seg1** (the base64url ciphertext segment) with `macKey`; reject on mismatch (constant-time compare). **Verify MAC before attempting decryption.**
2. Decrypt seg1 with `encKey` (first 16 bytes = IV); parse JSON; reject if `v != 1`.
3. Reject if `exp` is in the past (allow ≤ 60 s clock skew).
4. Normalize `p`: reject if it contains `..`, backslashes, a leading `/`, a drive letter, or a scheme. Must resolve to a path *inside* the configured share.
5. (Optional) Look up `n` in the nonce table; reject if already consumed outside the grace window (§6.4).
6. Only then fetch and stream the file.

> Authorization (does the *user* have rights to this file) is enforced at **mint time in Salesforce** (§8.2), not here — the Function cannot see Salesforce sharing. The Function's job is to prove the token is authentic, unexpired, unreplayed, and path-safe.

### 5.3 Request URL the add-in attaches

```
https://refuah-attach-gateway.azurewebsites.net/api/attach?token=<token>&code=<functionKey>
```

- `token` — the signed pass above.
- `code` — the Function-level access key (defense in depth; not secret enough to rely on alone, hence the HMAC).

Exchange fetches this URL server-side and receives the file with
`Content-Disposition: attachment; filename="cbc.pdf"`.

---

## 6. Function behavior

### 6.1 Trigger & route

- HTTP trigger, **GET** (Exchange issues GET; may also issue HEAD — support both).
- Route: `/api/attach`
- `authLevel: function` (requires `code` query/header).

### 6.2 Happy path

1. Read `token` from query string.
2. Validate per §5.2.
3. Build the file client for `payload.sh` + `payload.p` using Managed Identity.
4. `download()` the file as a stream.
5. Return `200` with the stream as body and headers:
   - `Content-Type`: from the file's properties (fallback `application/octet-stream`)
   - `Content-Length`: file size
   - `Content-Disposition: attachment; filename="<fn>"`
   - `Cache-Control: no-store`
6. Mark nonce consumed (if one-time enforcement is on).
7. Log `{ tokenNonce, status, callerIp, bytes }` to App Insights. **Do not log `path` or `fn`** — they carry patient identifiers. The nonce is sufficient to correlate a log entry back to the Salesforce mint event (which can log the path inside Salesforce's own PHI-scoped audit if needed).

### 6.3 Error handling

| Condition | HTTP | Body |
|---|---|---|
| Missing/garbled token | 400 | `invalid token` |
| Bad signature | 401 | `unauthorized` |
| Expired | 410 | `link expired` |
| Replayed nonce (past grace) | 410 | `link already used` |
| Path validation fails | 400 | `invalid path` |
| File not found in share | 404 | `not found` |
| Storage/identity error | 502 | `upstream error` |

Keep error bodies generic; put detail only in App Insights.

### 6.4 The "one-time use" nuance (important — read this)

Exchange Online's server-side fetch is **not guaranteed to be a single GET**. It may issue a `HEAD` first, or **multiple range requests**, or retry. Strict "destroy on first touch" single-use will therefore *break legitimate attachments*.

**Recommended policy (default):**
- Token lifetime: **5 minutes** (primary protection).
- One-time enforcement: **enabled, but with a 2-minute replay grace** — the first request records `firstSeen`; subsequent requests with the same nonce are allowed only within 120 s of `firstSeen`, then rejected. This tolerates Exchange's multi-request fetch while still preventing later reuse.
- HEAD requests do **not** consume the token; only the first successful GET starts the grace clock.

**Simpler alternative (if Table Storage is undesirable):** skip the nonce ledger entirely and rely solely on the 5-minute expiry + single-file scoping. This is still a defensible posture (the link is short-lived and reveals only one file the user is already emailing). Document whichever the security team prefers.

### 6.5 Size limits

Exchange attachment limits apply downstream (typically 25–150 MB depending on tenant config). The Function streams (does not buffer) the file, so it is not the bottleneck. Reject > configured `MAX_FILE_MB` (default 150) with `413`.

---

## 7. Reference implementation (Azure Functions, Node.js v4)

> The add-in repo is already Node/JS; keep the gateway in the same stack. Place under a sibling repo or `gateway/` folder.

### 7.1 `package.json`

```json
{
  "name": "refuah-attach-gateway",
  "version": "1.0.0",
  "main": "src/functions/*.js",
  "dependencies": {
    "@azure/functions": "^4.5.0",
    "@azure/identity": "^4.4.0",
    "@azure/storage-file-share": "^12.25.0",
    "@azure/data-tables": "^13.3.0"
  }
}
```

### 7.2 `src/functions/attach.js`

```js
const { app } = require("@azure/functions");
const { ShareServiceClient } = require("@azure/storage-file-share");
const { DefaultAzureCredential } = require("@azure/identity");
const { TableClient } = require("@azure/data-tables");
const crypto = require("crypto");

const ACCOUNT   = process.env.STORAGE_ACCOUNT;        // "refuahfiles"
const SECRET    = process.env.GATEWAY_HMAC_SECRET;    // from Key Vault ref (base64)
const MAX_BYTES = (parseInt(process.env.MAX_FILE_MB || "150", 10)) * 1024 * 1024;
const ENFORCE_ONE_TIME = (process.env.ENFORCE_ONE_TIME || "true") === "true";
const GRACE_MS  = 120000;

// Derive independent encryption + MAC keys from the single shared secret (domain separation).
// Salesforce derives the identical pair the same way.
const SECRET_BYTES = Buffer.from(SECRET, "base64");
const ENC_KEY = crypto.createHash("sha256").update(Buffer.concat([SECRET_BYTES, Buffer.from("enc")])).digest(); // 32 bytes
const MAC_KEY = crypto.createHash("sha256").update(Buffer.concat([SECRET_BYTES, Buffer.from("mac")])).digest(); // 32 bytes

// Managed-identity client for Azure Files (keyless). fileRequestIntent is required for AAD auth.
const credential = new DefaultAzureCredential();
const shareService = new ShareServiceClient(
  `https://${ACCOUNT}.file.core.windows.net`,
  credential,
  { fileRequestIntent: "backup" }
);

function b64urlDecode(s) {
  s = s.replace(/-/g, "+").replace(/_/g, "/");
  while (s.length % 4) s += "=";
  return Buffer.from(s, "base64");
}

function verifyToken(token) {
  const dot = token.indexOf(".");
  if (dot < 0) return { ok: false, code: 400, msg: "invalid token" };
  const seg1 = token.slice(0, dot);     // base64url( IV || ciphertext )
  const sigSeg = token.slice(dot + 1);

  // 1) Verify MAC over the ciphertext segment BEFORE decrypting (encrypt-then-MAC).
  const expected = crypto.createHmac("sha256", MAC_KEY).update(seg1).digest();
  const given = b64urlDecode(sigSeg);
  if (expected.length !== given.length ||
      !crypto.timingSafeEqual(expected, given)) {
    return { ok: false, code: 401, msg: "unauthorized" };
  }

  // 2) Decrypt the payload (first 16 bytes are the IV).
  let p;
  try {
    const blob = b64urlDecode(seg1);
    const iv = blob.subarray(0, 16);
    const ct = blob.subarray(16);
    const decipher = crypto.createDecipheriv("aes-256-cbc", ENC_KEY, iv);
    const plain = Buffer.concat([decipher.update(ct), decipher.final()]);
    p = JSON.parse(plain.toString("utf8"));
  } catch { return { ok: false, code: 400, msg: "invalid token" }; }

  if (p.v !== 1) return { ok: false, code: 400, msg: "invalid token" };
  if (typeof p.exp !== "number" || (Date.now() / 1000) > p.exp + 60) {
    return { ok: false, code: 410, msg: "link expired" };
  }
  // path safety
  const path = String(p.p || "");
  if (!path || path.includes("..") || path.includes("\\") ||
      path.startsWith("/") || /^[a-z]+:/i.test(path)) {
    return { ok: false, code: 400, msg: "invalid path" };
  }
  return { ok: true, payload: p };
}

async function checkNonce(context, nonce) {
  if (!ENFORCE_ONE_TIME) return { allowed: true };
  const table = TableClient.fromConnectionString(
    process.env.AzureWebJobsStorage, "gatewaytokens");
  try {
    const row = await table.getEntity("t", nonce);
    const age = Date.now() - new Date(row.firstSeen).getTime();
    return { allowed: age <= GRACE_MS, table };
  } catch {
    // first time we've seen this nonce
    await table.upsertEntity(
      { partitionKey: "t", rowKey: nonce, firstSeen: new Date().toISOString() },
      "Merge");
    return { allowed: true, table };
  }
}

app.http("attach", {
  methods: ["GET", "HEAD"],
  authLevel: "function",
  route: "attach",
  handler: async (req, context) => {
    const token = req.query.get("token");
    if (!token) return { status: 400, body: "invalid token" };

    const v = verifyToken(token);
    if (!v.ok) return { status: v.code, body: v.msg };
    const { p: path, sh: share, fn } = v.payload;

    // HEAD does not consume the token
    if (req.method !== "HEAD") {
      const n = await checkNonce(context, v.payload.n);
      if (!n.allowed) return { status: 410, body: "link already used" };
    }

    try {
      const fileClient = shareService
        .getShareClient(share)
        .rootDirectoryClient
        .getFileClient(path);

      const props = await fileClient.getProperties();
      if (props.contentLength > MAX_BYTES) return { status: 413, body: "file too large" };

      // Strip quotes AND CR/LF to prevent response-header (Content-Disposition) injection.
      const filename = (fn || path.split("/").pop()).replace(/["\r\n]/g, "");
      const headers = {
        "Content-Type": props.contentType || "application/octet-stream",
        "Content-Length": String(props.contentLength),
        "Content-Disposition": `attachment; filename="${filename}"`,
        "Cache-Control": "no-store"
      };

      // Log nonce only — NEVER path or fn (they carry patient identifiers).
      context.log("attach", { nonce: v.payload.n, bytes: props.contentLength,
                              ip: req.headers.get("x-forwarded-for") });

      if (req.method === "HEAD") return { status: 200, headers };

      const dl = await fileClient.download();
      return { status: 200, headers, body: dl.readableStreamBody };
    } catch (e) {
      context.error("upstream error", e.message);
      const status = /NotFound|404/.test(e.message) ? 404 : 502;
      return { status, body: status === 404 ? "not found" : "upstream error" };
    }
  }
});
```

### 7.3 App settings (Function App configuration)

| Setting | Value |
|---|---|
| `STORAGE_ACCOUNT` | `refuahfiles` |
| `GATEWAY_HMAC_SECRET` | `@Microsoft.KeyVault(SecretUri=https://kv-refuah-gateway.vault.azure.net/secrets/gateway-hmac/)` |
| `MAX_FILE_MB` | `150` |
| `ENFORCE_ONE_TIME` | `true` |
| `WEBSITE_VNET_ROUTE_ALL` | `1` (route outbound through the VNet) |
| `AzureWebJobsStorage` | (Function's own storage; hosts the `gatewaytokens` table) |
| `APPLICATIONINSIGHTS_CONNECTION_STRING` | (App Insights) |

---

## 8. Salesforce-side changes

The existing `AzureFileStorageController.generateFileSasUrl` stays for the in-browser (OWA/Gmail) path. Add a **parallel method** that mints the gateway token, and switch the LWC drag payload (the `text/plain` / `text/uri-list` values) to the gateway URL so **all** Outlook versions work.

### 8.1 New Custom Labels

| Label | Example value |
|---|---|
| `Gateway_Base_Url` | `https://refuah-attach-gateway.azurewebsites.net/api/attach` |
| `Gateway_Function_Key` | (the Function access key) |
| `Gateway_Hmac_Secret` | (same base64 secret stored in Key Vault — used to derive **both** the AES enc key and the MAC key) |
| `Gateway_Share_Name` | (the Azure Files share name; same as `Azure_Share_Name`) |
| `Gateway_Token_Ttl_Min` | `5` |

> Store `Gateway_Hmac_Secret` and `Gateway_Function_Key` as **Protected** Custom Labels or, better, in a Protected Custom Metadata / Named Credential. They are write-rarely, read-by-Apex-only. Both are **rotated on the schedule in §13**.

### 8.2 New Apex method (sketch)

> **Two security findings are addressed here (must land before production):**
> 1. **Authorization (finding #2).** The original sketch minted a valid token for *any* `filePath` string passed from the browser. Because the method is `@AuraEnabled` and the gateway bypasses the storage firewall, any authenticated org user could mint a URL for any file. The fix: require a **Case Id**, enforce **record-level sharing** via a SOQL query in this `with sharing` class, and confirm the requested path is **within that Case's patient folder** (the same folder `getPatientFolderPath` resolves). A token is then only ever minted for a file the running user is authorized to see.
> 2. **Encryption (finding #1).** The payload is **AES-256-CBC encrypted** (managed IV) before signing, so the file path never appears in clear in the URL. The Function decrypts (§7.2).

```apex
@AuraEnabled
public static String generateGatewayUrl(Id caseId, String filePath) {
    // --- 1) AUTHORIZATION GATE -------------------------------------------
    // 'with sharing' on this class makes this SELECT enforce the running user's
    // record access. No row => user can't see the Case => refuse to mint.
    List<Case> cases = [SELECT Id FROM Case WHERE Id = :caseId LIMIT 1];
    if (cases.isEmpty()) {
        throw new AuraHandledException('You do not have access to this record.');
    }
    // Confine the path to THIS case's patient folder. Defeats a caller who has
    // access to Case A trying to mint a URL for Case B's files.
    String allowedPrefix = getPatientFolderPath(caseId).folderPath; // existing method
    String requested = filePath.startsWith('/') ? filePath.substring(1) : filePath;
    if (String.isBlank(allowedPrefix) ||
        !requested.toLowerCase().startsWith(allowedPrefix.toLowerCase()) ||
        requested.contains('..')) {
        throw new AuraHandledException('Requested file is outside the patient folder.');
    }

    // --- 2) MINT (encrypt-then-MAC) --------------------------------------
    String secretB64 = Label.Gateway_Hmac_Secret;   // base64 of the shared secret
    String share     = Label.Gateway_Share_Name;
    String baseUrl   = Label.Gateway_Base_Url;
    String funcKey   = Label.Gateway_Function_Key;
    Integer ttl      = Integer.valueOf(Label.Gateway_Token_Ttl_Min);

    Long expEpoch = DateTime.now().addMinutes(ttl).getTime() / 1000;
    String nonce  = EncodingUtil.convertToHex(Crypto.generateAesKey(128)).substring(0, 16);
    String fn     = filePath.substringAfterLast('/');   // BARE filename only (no path)

    String payloadJson = '{"v":1,"p":' + JSON.serialize(requested)
        + ',"sh":' + JSON.serialize(share)
        + ',"n":"' + nonce + '"'
        + ',"exp":' + expEpoch
        + ',"fn":' + JSON.serialize(fn) + '}';

    // Derive enc/mac keys from the shared secret (must match the Function, §7.2):
    Blob secret  = EncodingUtil.base64Decode(secretB64);
    Blob encKey  = Crypto.generateDigest('SHA-256', concat(secret, Blob.valueOf('enc')));
    Blob macKey  = Crypto.generateDigest('SHA-256', concat(secret, Blob.valueOf('mac')));

    // AES-256-CBC with managed IV: Apex PREPENDS the 16-byte IV to the ciphertext,
    // which is exactly the layout the Function expects (IV || ciphertext).
    Blob ivAndCt = Crypto.encryptWithManagedIV('AES256', encKey, Blob.valueOf(payloadJson));
    String seg1  = base64Url(ivAndCt);
    Blob sig     = Crypto.generateMac('HmacSHA256', Blob.valueOf(seg1), macKey);
    String token = seg1 + '.' + base64Url(sig);

    return baseUrl + '?token=' + EncodingUtil.urlEncode(token, 'UTF-8')
                   + '&code=' + EncodingUtil.urlEncode(funcKey, 'UTF-8');
}

private static Blob concat(Blob a, Blob b) {
    // Blob concat via hex round-trip (Apex has no native Blob concat)
    return EncodingUtil.convertFromHex(
        EncodingUtil.convertToHex(a) + EncodingUtil.convertToHex(b));
}

// base64url helper (Apex base64 -> url-safe, strip padding)
private static String base64Url(Blob b) {
    return EncodingUtil.base64Encode(b)
        .replace('+', '-').replace('/', '_').replace('=', '');
}
```

> **Verify before relying on it:** confirm `getPatientFolderPath` returns a wrapper with a `folderPath` string (adjust the accessor to match the real return type), and confirm `Crypto.encryptWithManagedIV` produces `IV(16) || ciphertext` — it does in current Apex, but the unit test (§ implementation C3) should assert a decrypt round-trip rather than assume it. For the `searchScope = "all"` (cross-patient browse) case, the prefix-confinement check will block drags outside the current Case's folder; if that workflow must support attaching other patients' files, it needs its own explicit authorization rule rather than loosening this check.

### 8.3 LWC change (`azureFileBrowser.js`)

In `prefetchForDrag` / hover, also cache a gateway URL. **Pass the Case Id (`this.recordId`)** so Apex can run the authorization gate (§8.2):

```js
this.gatewayUrlCache[filePath] = await generateGatewayUrl({ caseId: this.recordId, filePath });
```

In `handleFileDragStart`, set the **gateway URL** (not the raw SAS URL) as the text payload, because the gateway URL is the one Exchange can actually fetch on every Outlook version:

```js
const gatewayUrl = this.gatewayUrlCache[filePath];
if (gatewayUrl) {
    event.dataTransfer.setData("text/uri-list", gatewayUrl);
    event.dataTransfer.setData("text/plain", gatewayUrl);
}
// keep the blob: DownloadURL branch unchanged for in-browser OWA/Gmail
```

No change is required in the add-in's `taskpane.js` — it already takes the first URL it finds and calls `addFileAttachmentAsync`. We are only changing *which* URL Salesforce hands it.

---

## 9. End-to-end request flow (the full round trip)

1. User opens a Case in Salesforce, hovers a file → LWC pre-fetches a **gateway token URL** from Apex.
2. User drags the file into the Outlook add-in task pane → `dataTransfer` carries the gateway URL as plain text.
3. Add-in calls `addFileAttachmentAsync(gatewayUrl, fileName)`.
4. Exchange Online performs a **server-side GET** on the gateway URL.
5. The Function validates the HMAC token + expiry (+ nonce), then reads the one named file from `refuahfiles` via **Managed Identity over the Private Endpoint**.
6. The Function streams the bytes back with `Content-Disposition: attachment`.
7. Exchange attaches the file to the composing email. Done — works identically on Outlook 2019, New Outlook, OWA, and Mac.

---

## 10. Cost estimate (rough, monthly)

| Resource | Est. cost |
|---|---|
| Function App (Flex Consumption, low volume) | ~$0–5 |
| Private Endpoint | ~$7 + minimal data |
| Private DNS Zone | ~$0.50 |
| Application Insights (low volume) | ~$0–3 |
| Table Storage (nonce ledger) | < $1 |
| **Total** | **~$10–20 / month** |

(Elastic Premium EP1 instead of Flex would add ~$150/mo; prefer Flex Consumption.)

---

## 11. Open questions for Positive MSP / Refuah

1. **Private Endpoint vs. resource-instance rule** — Private Endpoint is recommended and is the cleanest security story. Confirm the team is comfortable adding one (it does **not** open public access).
2. **Region (BLOCKER — must be answered before any Azure resource is created)** — confirm the region of `refuahfiles`. The Function, VNet, and Private Endpoint **must** be co-located in that same region; a PE in a different region than the storage account will not resolve to a private IP inside the VNet and the whole private path fails. Fill the `LOC` blank in the implementation guide before the working session.
3. **Key Vault** — use an existing vault or stand up `kv-refuah-gateway`?
4. **One-time enforcement** — enable the nonce ledger (recommended) or rely on 5-minute expiry alone?
5. **Who owns the Azure resources** — created in Refuah's subscription (recommended) with Logicfold given Contributor on a dedicated resource group?
6. **RBAC scope** — confirm the role assignment is scoped to the single patient-files share (§4, step A5), not the storage account, and that this least-privileged role is sufficient for Entra/Files data-plane read.
7. **Ongoing ownership** — who owns runtime patching, key rotation, monitoring, and DR after go-live? (See §14.)

---

## 12. Acceptance criteria

- [ ] Storage account `refuahfiles` public network access remains **Disabled** after rollout.
- [ ] A dragged patient file attaches successfully on **Outlook 2019**, New Outlook, OWA, and Outlook for Mac.
- [ ] An expired token returns `410` and does not stream any bytes.
- [ ] A tampered token returns `401`.
- [ ] A token for `file A` cannot retrieve `file B` (path is encrypted + signed).
- [ ] **The token URL reveals no PHI** — base64-decoding the token segment yields ciphertext, not a readable patient path.
- [ ] **`generateGatewayUrl` refuses to mint** for a Case the running user cannot access, and for a path outside that Case's patient folder.
- [ ] Every attach is visible in Application Insights by **nonce and outcome — with no file path or filename in the log**.
- [ ] No account key is stored in the Function App configuration.
- [ ] The managed identity's RBAC role is scoped to the **single share**, not the storage account.
- [ ] A documented **secret/key rotation** procedure exists and has been dry-run once (§13).

---

## 13. Secret & key rotation

Two credentials are shared or exposed and must be rotatable without breaking in-flight attachments:

### 13.1 Shared secret (`Gateway_Hmac_Secret` ↔ Key Vault `gateway-hmac`)

Because tokens live at most **5 minutes** (the TTL), a hard cut-over creates at most a 5-minute window of broken links — but we still do it gracefully:

1. Generate a new secret; add it as a **new version** of the Key Vault secret.
2. **Dual-accept window:** temporarily have the Function try the new key first, then the old, for verification/decryption (carry `gateway-hmac` and `gateway-hmac-prev` app settings). This lets already-minted tokens keep working.
3. Update the Salesforce `Gateway_Hmac_Secret` label to the new value (new mints now use the new key).
4. After the TTL window (≥5 min) **plus a safety margin**, remove the old key from the Function. 
5. If a graceful overlap is not implemented, schedule the swap in a low-traffic window and accept the ≤5-minute gap.

**Cadence:** annually, and immediately on any suspected exposure.

### 13.2 Function access key (`code=` / `Gateway_Function_Key`)

This key ships in every URL and therefore lands in mail logs — treat it as low-trust (the HMAC + encryption are the real guard) but still rotate it:

1. Create a **second** function key in Azure (Functions supports multiple named keys).
2. Update the Salesforce `Gateway_Function_Key` label to the new key.
3. Confirm new mints work, then **delete** the old key. Zero downtime (both keys valid during the overlap).

**Cadence:** annually, and on suspected exposure.

### 13.3 Owner

Logicfold executes rotation; Positive MSP approves and witnesses. Record each rotation date in the run-book.

---

## 14. Post-go-live ownership & operations

| Area | Owner | Cadence / trigger |
|---|---|---|
| App Insights monitoring (spikes in 401/410, anomalous IPs, error rate) | **Logicfold** | Weekly review + alert on threshold |
| Function runtime / Node.js version patching (Azure EOLs language versions) | **Logicfold** | On Azure deprecation notice; check semi-annually |
| Secret & function-key rotation (§13) | **Logicfold** executes / **Positive** approves | Annual + on exposure |
| Private Endpoint / VNet / DNS health | **Positive MSP** (Azure infra) | On Azure platform changes |
| RBAC role assignment review (still least-privilege, still share-scoped) | **Positive MSP** | Annual access review |
| DR / outage response (gateway down ⇒ Outlook-2019 attach fails for all staff) | **Logicfold** + **Positive** | On incident; see 14.1 |

### 14.1 Outage behavior & fallback

If the Function App is down, Outlook 2019 attachments fail (the URL Exchange fetches returns 5xx). New Outlook / OWA are unaffected (they use the in-browser base64 path). **Fallback:** the LWC can be reverted in one line to hand out the raw SAS URL again (see implementation guide rollback) — this restores OWA/New-Outlook fully and is a no-op for 2019 (which never worked without the gateway). Target: Flex Consumption + App Insights availability alert; document the redeploy/restart run-book step.

---

*See `azure-gateway-implementation-steps.md` for the exact click-by-click / command-by-command build, split by Azure side, Outlook/Exchange side, and Salesforce side.*
