# Azure File Attachment Gateway — Step-by-Step Implementation

Companion to `azure-gateway-spec.md`. The build is split into **three independent parts**, each a self-contained work package with its own owner, inputs, and outputs:

| Part | Name | Owner | What it delivers |
|---|---|---|---|
| **PART 1** | **Gateway setup** (Azure) | Azure admin (Refuah/Positive) + Logicfold | A live, locked-down Function that streams one file per signed token over a private path |
| **PART 2** | **Add-in setup** (Outlook / Exchange) | Logicfold (manifest) + M365 admin (deploy) | The "Attach Files" add-in available in staff Outlook, including Outlook 2019 |
| **PART 3** | **Salesforce setup** | Logicfold | The LWC mints an authorized, encrypted gateway URL and hands it to the add-in on drag |

### How the parts connect (dependency order)

```
PART 1 (Gateway)  ──┬── outputs: Function base URL, function key, shared secret, share name
                    │
                    ▼
PART 3 (Salesforce) ── consumes those 4 values → mints the URL the add-in attaches
                    │
                    ▼
PART 2 (Add-in)  ── already deployed; carries whatever URL Salesforce hands it. Final UAT happens here.
```

> **Sequence:** **Part 1 → Part 3 → Part 2 verification.** Part 1 must exist before Part 3 can be configured (Part 3 needs Part 1's URL/key/secret). Part 2's add-in is already built and deployed, so it's mostly confirmation + end-to-end UAT once Parts 1 and 3 are live. Parts 1 and 2's manifest commit (B1) can proceed in parallel.

Each step lists who does it and how to verify before moving on.

Legend: 🟦 = Azure Portal · 💻 = CLI (`az`) · 🧪 = verification step · 📋 = value to record for a later part

---

## Prerequisites / decisions to lock before starting

- [ ] **(BLOCKER)** Region of the `refuahfiles` storage account (everything co-locates here): `__________`  ← must be filled before A1; a Private Endpoint in a different region than the storage account will not resolve to a private IP and the whole private path fails.
- [ ] Subscription + resource group to hold the gateway: `__________`
- [ ] The Azure Files **share name** inside `refuahfiles`: `__________`  *(matches Salesforce Custom Label `Azure_Share_Name`)*
- [ ] Key Vault to use (existing or new `kv-refuah-gateway`): `__________`
- [ ] One-time enforcement on? (recommended **yes**): `__________`
- [ ] Who creates resources & who gets Contributor on the RG: `__________`
- [ ] Who owns ongoing ops post go-live (monitoring, patching, rotation, DR — see "Post go-live ownership"): `__________`

CLI variables used throughout (set these once in a shell):

```bash
RG=rg-refuah-gateway
LOC=eastus                      # <-- the refuahfiles region
SUB=<subscription-id>
STORAGE=refuahfiles
SHARE=<share-name>
FUNC=refuah-attach-gateway
PLAN=plan-refuah-gateway
VNET=vnet-refuah-gateway
SUBNET=snet-func
KV=kv-refuah-gateway
az account set --subscription "$SUB"
```

---

# PART 1 — Gateway setup (Azure)

**Goal:** stand up `refuah-attach-gateway` — a Function that, given a signed token, streams exactly one file from `refuahfiles` over a Private Endpoint, while the storage account stays firewalled to the public internet.

**Owner:** Azure admin (Refuah/Positive) for the infra (A1–A6); Logicfold for the code + config (A7–A11).

**Needs going in (from Prerequisites):** confirmed region, subscription/RG, share name, Key Vault choice.

**Hands off to Part 3 (record these as you go):**
- 📋 Function base URL → Salesforce label `Gateway_Base_Url` (from A9)
- 📋 Function key → `Gateway_Function_Key` (from A9)
- 📋 Shared secret → `Gateway_Hmac_Secret` (from A6)
- 📋 Share name → `Gateway_Share_Name` (from Prerequisites)

**Done when:** A11 smoke test passes (file streams, tampered/expired tokens rejected, token is PHI-opaque, logs carry no path).

---

### A1. Create the resource group 🟦/💻
*Owner: Azure admin*

```bash
az group create -n "$RG" -l "$LOC"
```

🧪 `az group show -n "$RG"` returns the group.

---

### A2. Create the VNet + delegated subnet for the Function 🟦/💻
*Owner: Azure admin*

The Function needs **regional VNet integration** so its outbound calls to storage travel the private path. The subnet must be **delegated to `Microsoft.App`** (Flex Consumption) or `Microsoft.Web/serverFarms` (Elastic Premium).

```bash
az network vnet create -g "$RG" -n "$VNET" -l "$LOC" \
  --address-prefix 10.20.0.0/16 \
  --subnet-name "$SUBNET" --subnet-prefix 10.20.1.0/24

# Delegation (Flex Consumption):
az network vnet subnet update -g "$RG" --vnet-name "$VNET" -n "$SUBNET" \
  --delegations Microsoft.App/environments

# Add a second subnet to host the storage Private Endpoint:
az network vnet subnet create -g "$RG" --vnet-name "$VNET" \
  -n snet-pe --address-prefix 10.20.2.0/24 \
  --disable-private-endpoint-network-policies true
```

🧪 `az network vnet subnet list -g "$RG" --vnet-name "$VNET" -o table` shows `snet-func` (delegated) and `snet-pe`.

---

### A3. Create the Private Endpoint to `refuahfiles` (file sub-resource) 🟦/💻
*Owner: Azure admin (this is the key "grant the Function private access" step)*

```bash
STORAGE_ID=$(az storage account show -n "$STORAGE" -g <storage-rg> --query id -o tsv)

az network private-endpoint create -g "$RG" -n pe-refuahfiles-file -l "$LOC" \
  --vnet-name "$VNET" --subnet snet-pe \
  --private-connection-resource-id "$STORAGE_ID" \
  --group-id file \
  --connection-name pe-refuahfiles-file-conn
```

Then wire up **Private DNS** so `refuahfiles.file.core.windows.net` resolves to the private IP:

```bash
az network private-dns zone create -g "$RG" -n privatelink.file.core.windows.net
az network private-dns link vnet create -g "$RG" -n dns-link \
  --zone-name privatelink.file.core.windows.net --virtual-network "$VNET" --registration-enabled false
az network private-endpoint dns-zone-group create -g "$RG" \
  --endpoint-name pe-refuahfiles-file -n default \
  --private-dns-zone privatelink.file.core.windows.net --zone-name file
```

🧪 From a VM/Function in the VNet, `nslookup refuahfiles.file.core.windows.net` resolves to a `10.20.2.x` private IP (not a public IP).

> **Storage firewall stays Disabled for public.** The Private Endpoint is what grants access. Do **not** add public IP rules.

---

### A4. Create the Function App (Flex Consumption) 🟦/💻
*Owner: Azure admin / Logicfold*

```bash
# Function needs its own storage for runtime + the nonce table; can be a new small account
az storage account create -n stgatewayruntime -g "$RG" -l "$LOC" --sku Standard_LRS

az functionapp create -g "$RG" -n "$FUNC" -l "$LOC" \
  --flexconsumption-location "$LOC" \
  --runtime node --runtime-version 20 \
  --storage-account stgatewayruntime \
  --os-type Linux
```

Enable **system-assigned managed identity**:

```bash
az functionapp identity assign -g "$RG" -n "$FUNC"
PRINCIPAL_ID=$(az functionapp identity show -g "$RG" -n "$FUNC" --query principalId -o tsv)
```

Connect the Function's **outbound** to the VNet:

```bash
az functionapp vnet-integration add -g "$RG" -n "$FUNC" --vnet "$VNET" --subnet "$SUBNET"
az functionapp config appsettings set -g "$RG" -n "$FUNC" --settings WEBSITE_VNET_ROUTE_ALL=1
```

🧪 `az functionapp show -g "$RG" -n "$FUNC" --query state` = `Running`; identity `principalId` is non-empty.

---

### A5. Grant the managed identity read access to Azure Files 🟦/💻
*Owner: Azure admin*

Azure Files data-plane over Azure AD requires an RBAC role. Use **Storage File Data Privileged Reader** (reads file data bypassing SMB permissions — appropriate for a server gateway).

> **Scope it to the single share, not the storage account (security finding #3).** Scoping to `$STORAGE_ID` grants read on **every** share on `refuahfiles`. The gateway only ever needs the one patient-files share, so scope the assignment to that share's resource ID — the gateway then physically cannot read any other share.

```bash
SHARE_SCOPE="$STORAGE_ID/fileServices/default/shares/$SHARE"

az role assignment create \
  --assignee-object-id "$PRINCIPAL_ID" --assignee-principal-type ServicePrincipal \
  --role "Storage File Data Privileged Reader" \
  --scope "$SHARE_SCOPE"
```

🧪 `az role assignment list --assignee "$PRINCIPAL_ID" --scope "$SHARE_SCOPE" -o table` shows the role at **share** scope (the Scope column ends in `/shares/<share>`, not the account).

> If the data-plane role cannot be assigned at share scope in your tenant, fall back to account scope **and document explicitly** that only the one patient-files share exists / contains PHI on this account. Prefer share scope.

> Note: AAD auth to Azure **Files** over REST requires `x-ms-version: 2022-11-02`+ and the `x-ms-file-request-intent: backup` header — the SDK sends these when `fileRequestIntent: "backup"` is set (see spec §7.2).

---

### A6. Create the HMAC shared secret in Key Vault 🟦/💻
*Owner: Azure admin*

Generate a strong random secret (base64, 32 bytes) and store it:

```bash
az keyvault create -g "$RG" -n "$KV" -l "$LOC" --enable-rbac-authorization true
SECRET_VALUE=$(openssl rand -base64 32)
az keyvault secret set --vault-name "$KV" -n gateway-hmac --value "$SECRET_VALUE"

# Let the Function's identity read the secret:
KV_ID=$(az keyvault show -n "$KV" --query id -o tsv)
az role assignment create --assignee-object-id "$PRINCIPAL_ID" --assignee-principal-type ServicePrincipal \
  --role "Key Vault Secrets User" --scope "$KV_ID"
```

📋 **Record `$SECRET_VALUE`** — the identical value goes into Salesforce (Track C2). This is the only secret shared between the two systems. Both sides derive an **AES encryption key** and a **MAC key** from it (`SHA256(secret||"enc")`, `SHA256(secret||"mac")`) — see spec §5.1 / §7.2. Rotation procedure: spec §13.

🧪 `az keyvault secret show --vault-name "$KV" -n gateway-hmac --query value -o tsv` returns the value.

---

### A7. Create the nonce table (if one-time enforcement is on) 💻
*Owner: Logicfold*

```bash
az storage table create --account-name stgatewayruntime --name gatewaytokens \
  --auth-mode login
```

(Optionally set a lifecycle/TTL cleanup; rows are tiny and self-expire in logic via the grace window.)

---

### A8. Configure Function app settings 💻
*Owner: Logicfold*

```bash
az functionapp config appsettings set -g "$RG" -n "$FUNC" --settings \
  STORAGE_ACCOUNT=$STORAGE \
  GATEWAY_HMAC_SECRET="@Microsoft.KeyVault(SecretUri=https://$KV.vault.azure.net/secrets/gateway-hmac/)" \
  MAX_FILE_MB=150 \
  ENFORCE_ONE_TIME=true
```

🧪 In Portal → Function App → Environment variables, `GATEWAY_HMAC_SECRET` shows a green "Key Vault Reference" resolved status.

---

### A9. Deploy the Function code 💻
*Owner: Logicfold*

Use the reference implementation in spec §7. From the gateway project folder:

```bash
func azure functionapp publish "$FUNC"
```

Grab the function key for Salesforce:

```bash
az functionapp function keys list -g "$RG" -n "$FUNC" --function-name attach -o table
# record the "default" key  -> Salesforce Label Gateway_Function_Key
```

🧪 The function shows in Portal under Functions; `Invoke URL` is `https://refuah-attach-gateway.azurewebsites.net/api/attach`.

---

### A10. Lock down ingress (CORS / TLS) 🟦/💻
*Owner: Logicfold*

- This is a **server-to-server** endpoint (Exchange fetches it). Set **CORS allowed origins to empty** (no browser origins) so it can't be abused from a page.
- Ensure **HTTPS Only = On** and **Minimum TLS = 1.2**.

```bash
az functionapp update -g "$RG" -n "$FUNC" --set httpsOnly=true
az functionapp config set -g "$RG" -n "$FUNC" --min-tls-version 1.2
```

---

### A11. Smoke test the gateway in isolation 🧪
*Owner: Logicfold*

Mint a token by hand (small Node script using the same secret) for a known test file, then:

```bash
curl -v "https://refuah-attach-gateway.azurewebsites.net/api/attach?token=<token>&code=<funcKey>" -o out.bin
```

Verify:
- [ ] `200` with `Content-Disposition: attachment` and correct bytes.
- [ ] Tampered token → `401`.
- [ ] Expired token → `410`.
- [ ] Token for file A cannot fetch file B.
- [ ] **Token is PHI-opaque:** base64url-decoding the token's first segment yields ciphertext (binary), **not** readable JSON containing the patient path.
- [ ] **`Content-Disposition` filename** is the bare filename only (no patient folder / DOB), and a filename containing `\r`/`\n` does not inject a header.
- [ ] App Insights shows the request by **nonce + outcome only — no path, no filename** in the log line.

---

# PART 2 — Add-in setup (Outlook / Exchange)

**Goal:** the "Attach Files" task-pane add-in is installed and working in staff Outlook — **including classic Outlook 2019** — and correctly attaches whatever URL Salesforce hands it.

**Owner:** Logicfold (manifest) + M365 admin / Positive (centralized deploy + UAT).

**Needs going in:** nothing from Parts 1/3 for B1–B3 (the manifest carries no secrets — it just points the browser at the add-in). B4 end-to-end UAT needs Parts 1 and 3 live.

**Hands off:** confirmation the add-in loads on every Outlook client; final proof the whole chain works.

> The add-in manifest is **already built and (per the May thread) deployed for testing** via Positive's GitHub fork + the M365 admin center. Most of this part is verification, plus one manifest item to confirm. B1 (the manifest commit) can run in parallel with Part 1.

---

### B1. Confirm the manifest MinVersion is 1.5 (not 1.8) ✅
*Owner: Logicfold*

Outlook 2019 is requirement set 1.6. The manifest must advertise **MinVersion 1.5** so 2019 loads it (the JS feature-detects 1.8 at runtime). This was changed in the working copy but **is currently uncommitted** in `outlook-addin/manifest.xml`.

- [ ] Commit & push the manifest change (`Mailbox MinVersion 1.5` in both `<Set>` and `<bt:Sets DefaultMinVersion>`).
- [ ] Republish to GitHub Pages (`sweinfeld1234.github.io/file-dropper/`) or Refuah's fork.

🧪 `https://sweinfeld1234.github.io/file-dropper/manifest.xml` shows `MinVersion="1.5"`.

### B2. Confirm centralized deployment in M365 admin center
*Owner: M365 admin (Positive)*

- M365 Admin → **Settings → Integrated apps** → the "Azure File Attacher" custom app is uploaded and **assigned to the relevant users/group**.
- The manifest URL points to the production host (GitHub Pages / Refuah fork).

🧪 In a target user's Outlook (any version), composing a new email shows the **"Attach Files"** button in the ribbon.

### B3. No Exchange allow-listing needed ✅
*Owner: M365 admin*

There is **nothing to allow-list on the Exchange/M365 side.** Exchange fetches the gateway URL over the public internet like any link attachment. The security boundary lives at the gateway (HMAC token) and the storage (Private Endpoint). Document this explicitly so the security team isn't expecting an Exchange firewall change.

### B4. End-to-end test on Outlook 2019
*Owner: Refuah staff + Logicfold*

After Track C is live:
- [ ] Outlook 2019 (Build 1808 / 10417.20132, Click-to-Run — the confirmed staff build): drag a patient file from Salesforce → add-in → **file attaches**.
- [ ] Repeat on New Outlook, OWA, Outlook for Mac.

---

# PART 3 — Salesforce setup

**Goal:** when a user drags a patient file, Salesforce mints an **authorized, encrypted, short-lived** gateway URL (pointing at the Part 1 Function) and puts it on the clipboard for the add-in to attach.

**Owner:** Logicfold.

**Needs going in (from Part 1):** the four recorded values — `Gateway_Base_Url`, `Gateway_Function_Key`, `Gateway_Hmac_Secret`, `Gateway_Share_Name`. Do not start C1 until Part 1's A9 is done.

**Hands off to Part 2:** the live URL the add-in attaches (verified end-to-end in B4).

**Done when:** C5 deploy passes with tests, and a real drag in Salesforce produces a gateway URL (verified in the browser console).

---

### C1. Add Custom Labels
*Owner: Logicfold*

Create (Setup → Custom Labels), values from Track A:

| Label | Value |
|---|---|
| `Gateway_Base_Url` | `https://refuah-attach-gateway.azurewebsites.net/api/attach` |
| `Gateway_Function_Key` | (from A9) |
| `Gateway_Hmac_Secret` | (the base64 secret from A6 — **identical**) |
| `Gateway_Share_Name` | (the share name; same as `Azure_Share_Name`) |
| `Gateway_Token_Ttl_Min` | `5` |

> Treat `Gateway_Hmac_Secret` and `Gateway_Function_Key` as protected. Do not log them.

### C2. Add the `generateGatewayUrl` Apex method
*Owner: Logicfold*

Add the method from spec §8.2 to `AzureFileStorageController.cls`. Key points:
- **Signature is `generateGatewayUrl(Id caseId, String filePath)`** — caseId is required for the authorization gate.
- **Authorization gate (finding #2):** SELECT the Case (the class is `with sharing`, so this enforces the running user's record access); if no row, throw. Then confirm `filePath` starts with the Case's resolved patient-folder prefix (`getPatientFolderPath`) and contains no `..`; else throw. Never mint for a file the user isn't authorized to see.
- **Encrypt-then-MAC (finding #1):** derive `encKey = SHA256(secret||'enc')` and `macKey = SHA256(secret||'mac')`; `Crypto.encryptWithManagedIV('AES256', encKey, payload)` (Apex prepends the 16-byte IV); base64url that as seg1; `Crypto.generateMac('HmacSHA256', seg1, macKey)` as the signature.
- `fn` must be the bare filename (`substringAfterLast('/')`), not the path.
- Return `baseUrl?token=<urlencoded>&code=<urlencoded funcKey>`.

🧪 In anonymous Apex (use a real Case Id the running user can access, and a path inside that Case's folder):
```apex
System.debug(AzureFileStorageController.generateGatewayUrl('500XXXXXXXXXXXX', 'PatientFolder/test.pdf'));
```
Copy the URL, `curl` it (Track A11) → file downloads. Confirm a Case the user **cannot** access throws, and a path **outside** the Case folder throws.

### C3. Add a matching unit test
*Owner: Logicfold*

Extend `AzureFileStorageControllerTest.cls`:
- [ ] Token has two dot-separated segments.
- [ ] **Segment 1 is ciphertext, not plaintext JSON** — decoding it does NOT yield a readable patient path (proves finding #1 is fixed). Then decrypt it with `encKey` and assert the round-trip yields the expected `p`, `sh`, `exp` in the future.
- [ ] Recomputed HMAC (with `macKey`) over segment 1 equals segment 2.
- [ ] **Authorization:** as a user without access to the Case, `generateGatewayUrl` throws (use `System.runAs` + a Case the user can't see). A path outside the Case's patient folder throws.
- [ ] `fn` in the decrypted payload is the bare filename (no `/`).
- [ ] URL starts with the `Gateway_Base_Url` label and contains `&code=`.

(No callout in this method, so no `HttpCalloutMock` needed — it's pure crypto/string + a SOQL access check.)

### C4. Update the LWC drag payload
*Owner: Logicfold*

In `force-app/main/default/lwc/azureFileBrowser/azureFileBrowser.js`:

1. Import the new Apex method:
   ```js
   import generateGatewayUrl from "@salesforce/apex/AzureFileStorageController.generateGatewayUrl";
   ```
2. Add a `gatewayUrlCache = {}` field.
3. In the hover pre-fetch (where `sasUrlCache` is populated, ~line 805), also cache the gateway URL — **passing the Case Id** so Apex can run the authorization gate:
   ```js
   this.gatewayUrlCache[filePath] = await generateGatewayUrl({ caseId: this.recordId, filePath });
   ```
4. In `handleFileDragStart` (~line 834), set the **gateway URL** into `text/uri-list` and `text/plain` (replacing the raw SAS URL for the Outlook-desktop path). Keep the `DownloadURL` blob branch unchanged for in-browser OWA/Gmail.

🧪 In the browser console during a drag, `[DragDrop] dragstart types:` includes `text/plain`, and its value is the `…azurewebsites.net/api/attach?token=…` URL.

### C5. Deploy & run tests
*Owner: Logicfold*

```bash
sf project deploy start --test-level RunLocalTests
```

🧪 Deploy succeeds; `AzureFileStorageControllerTest` passes; org coverage ≥ 80%.

---

## Cut-over & rollback

**Cut-over:** Once C4 is deployed, every new drag uses the gateway URL automatically — no user action needed. The in-browser path is unaffected.

**Rollback:** Revert the LWC `handleFileDragStart` to set the raw `sasUrlCache[filePath]` again (one-line change) and redeploy. The gateway can stay deployed but idle. No Azure teardown required for rollback. This also doubles as the **outage fallback** (spec §14.1): if the Function is down, reverting restores OWA/New-Outlook immediately (Outlook 2019 simply returns to its pre-project state of not working without the gateway).

---

## Post go-live ownership, rotation & DR

(Full detail in spec §13–§14. Lock owners during the prerequisites step.)

- **Monitoring** — Logicfold reviews App Insights weekly; set an alert on 401/410 spikes and on availability. *(Owner: ________)*
- **Runtime patching** — Azure EOLs Node.js versions; bump the Function runtime before EOL. Check semi-annually. *(Owner: ________)*
- **Secret rotation** (`Gateway_Hmac_Secret` ↔ Key Vault) — annual + on suspected exposure. New Key Vault version → optional dual-accept overlap → update Salesforce label → retire old after ≥ TTL window. *(Owner: ________)*
- **Function-key rotation** (`code=`) — annual + on exposure. Create 2nd key → update Salesforce label → delete old (zero downtime). *(Owner: ________)*
- **RBAC review** — annual confirmation the role is still share-scoped and least-privilege. *(Owner: ________)*
- **DR / outage** — gateway down ⇒ Outlook-2019 attach fails; use the one-line LWC rollback as fallback; redeploy/restart run-book. *(Owner: ________)*

---

## Final verification matrix

| Client | Path used | Expected |
|---|---|---|
| Outlook 2019 (classic) | gateway URL → Exchange server-side fetch | ✅ attaches |
| New Outlook (Windows) | gateway URL (or base64) | ✅ attaches |
| Outlook on the Web | blob `DownloadURL` (same Chrome) | ✅ attaches |
| Outlook for Mac | gateway URL | ✅ attaches |
| Storage public access | — | ❌ still **Disabled** |
| Expired / tampered token | — | ❌ `410` / `401`, no bytes |
| Token URL inspected | base64-decode segment 1 | ❌ ciphertext only — **no readable patient path** |
| Mint for inaccessible Case | `generateGatewayUrl` | ❌ throws — no token issued |
| Mint for path outside Case folder | `generateGatewayUrl` | ❌ throws — no token issued |
| App Insights log line | — | ✅ nonce + outcome; ❌ no path/filename |
| RBAC scope | role assignment | ✅ single share, not account |

---

## Who does what — quick reference

| Part | Step | Owner |
|---|---|---|
| **1 — Gateway** | A1–A6 (RG, VNet, Private Endpoint, identity, RBAC, Key Vault) | **Azure admin (Refuah/Positive)** — the ~"15–20 min working session" from the email, realistically ~1–2 hrs |
| **1 — Gateway** | A7–A11 (table, settings, deploy, test) | **Logicfold** |
| **2 — Add-in** | B1–B2 (manifest + centralized deploy) | **Logicfold** (manifest) + **M365 admin** (assign) |
| **2 — Add-in** | B3–B4 (verify, UAT) | **M365 admin + Refuah staff** |
| **3 — Salesforce** | C1–C5 (Labels, Apex, LWC, deploy) | **Logicfold** |

> **Critical-path summary:** Part 1 (A1–A11) is the long pole and unblocks everything. Part 3 (C1–C5) follows once Part 1 hands over its four values. Part 2 is a manifest commit (B1, parallelizable) plus end-to-end UAT (B4) at the very end.
