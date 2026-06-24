const { app } = require("@azure/functions");
const { ShareServiceClient } = require("@azure/storage-file-share");
const { DefaultAzureCredential } = require("@azure/identity");
const { TableClient } = require("@azure/data-tables");
const crypto = require("crypto");

/* ============================================================================
   Refuah Attach Gateway — Azure Function
   ----------------------------------------------------------------------------
   Given a short-lived, AES-encrypted + HMAC-signed token (minted by Salesforce),
   stream exactly ONE file from the firewalled `refuahfiles` Azure Files share.
   The Function reaches storage over a Private Endpoint using its Managed
   Identity (no account keys); the storage account stays closed to the public.

   Token format (must match the Salesforce minter):
     seg1  = base64url( IV(16) || AES-256-CBC(payloadJson) )
     token = seg1 + "." + base64url( HMAC_SHA256(seg1, macKey) )
   enc/mac keys are derived from the shared secret via SHA-256 domain separation:
     encKey = SHA256(secret || "enc")
     macKey = SHA256(secret || "mac")
   ========================================================================== */

const ACCOUNT   = process.env.STORAGE_ACCOUNT;        // "refuahfiles"
const MAX_BYTES = (parseInt(process.env.MAX_FILE_MB || "150", 10)) * 1024 * 1024;
const ENFORCE_ONE_TIME = (process.env.ENFORCE_ONE_TIME || "true") === "true";
const GRACE_MS  = 120000; // Exchange may issue multiple range GETs / retries.

/*
 * Build candidate key pairs from the current secret AND an optional previous
 * secret (GATEWAY_HMAC_SECRET_PREV). During a rotation overlap window both are
 * accepted, so in-flight tokens keep working with zero downtime.
 * (Security punch-list item #5.)
 */
function keyPairsFromEnv() {
  const secrets = [process.env.GATEWAY_HMAC_SECRET, process.env.GATEWAY_HMAC_SECRET_PREV]
    .filter(Boolean);
  return secrets.map((b64) => {
    const secret = Buffer.from(b64, "base64");
    return {
      enc: crypto.createHash("sha256").update(Buffer.concat([secret, Buffer.from("enc")])).digest(),
      mac: crypto.createHash("sha256").update(Buffer.concat([secret, Buffer.from("mac")])).digest()
    };
  });
}
const KEY_PAIRS = keyPairsFromEnv();

// Managed-identity client for Azure Files (keyless). fileRequestIntent is
// required for Azure AD (Entra) auth to the Files data plane.
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

function validatePayload(p) {
  if (!p || p.v !== 1) return { ok: false, code: 400, msg: "invalid token" };
  if (typeof p.exp !== "number" || (Date.now() / 1000) > p.exp + 60) {
    return { ok: false, code: 410, msg: "link expired" };
  }
  const path = String(p.p || "");
  if (!path || path.includes("..") || path.includes("\\") ||
      path.startsWith("/") || /^[a-z]+:/i.test(path)) {
    return { ok: false, code: 400, msg: "invalid path" };
  }
  return { ok: true, payload: p };
}

function verifyToken(token) {
  const dot = token.indexOf(".");
  if (dot < 0) return { ok: false, code: 400, msg: "invalid token" };
  const seg1 = token.slice(0, dot);     // base64url( IV || ciphertext )
  const given = b64urlDecode(token.slice(dot + 1));

  // Try each candidate key pair (current first, then previous). MAC-verify
  // before decrypting (encrypt-then-MAC).
  for (let i = 0; i < KEY_PAIRS.length; i++) {
    const kp = KEY_PAIRS[i];
    const expected = crypto.createHmac("sha256", kp.mac).update(seg1).digest();
    if (expected.length === given.length && crypto.timingSafeEqual(expected, given)) {
      try {
        const blob = b64urlDecode(seg1);
        const iv = blob.subarray(0, 16);
        const ct = blob.subarray(16);
        const decipher = crypto.createDecipheriv("aes-256-cbc", kp.enc, iv);
        const plain = Buffer.concat([decipher.update(ct), decipher.final()]);
        return validatePayload(JSON.parse(plain.toString("utf8")));
      } catch (e) {
        return { ok: false, code: 400, msg: "invalid token" };
      }
    }
  }
  return { ok: false, code: 401, msg: "unauthorized" };
}

async function checkNonce(nonce) {
  if (!ENFORCE_ONE_TIME) return { allowed: true };
  const table = TableClient.fromConnectionString(process.env.AzureWebJobsStorage, "gatewaytokens");
  try {
    const row = await table.getEntity("t", nonce);
    const age = Date.now() - new Date(row.firstSeen).getTime();
    return { allowed: age <= GRACE_MS };
  } catch (e) {
    // First time we've seen this nonce — record firstSeen, allow.
    await table.upsertEntity(
      { partitionKey: "t", rowKey: nonce, firstSeen: new Date().toISOString() }, "Merge");
    return { allowed: true };
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

    // HEAD does not consume the token (Exchange often probes with HEAD first).
    if (req.method !== "HEAD") {
      const n = await checkNonce(v.payload.n);
      if (!n.allowed) return { status: 410, body: "link already used" };
    }

    try {
      const fileClient = shareService
        .getShareClient(share)
        .rootDirectoryClient
        .getFileClient(path);

      const props = await fileClient.getProperties();
      if (props.contentLength > MAX_BYTES) return { status: 413, body: "file too large" };

      // Strip quotes AND CR/LF to prevent response-header injection.
      const filename = (fn || path.split("/").pop()).replace(/["\r\n]/g, "");
      const headers = {
        "Content-Type": props.contentType || "application/octet-stream",
        "Content-Length": String(props.contentLength),
        "Content-Disposition": `attachment; filename="${filename}"`,
        "Cache-Control": "no-store"
      };

      // Log nonce + outcome ONLY — never path or fn (they carry patient identifiers).
      context.log("attach ok", {
        nonce: v.payload.n,
        bytes: props.contentLength,
        ip: req.headers.get("x-forwarded-for")
      });

      if (req.method === "HEAD") return { status: 200, headers };

      const dl = await fileClient.download();
      return { status: 200, headers, body: dl.readableStreamBody };
    } catch (e) {
      // SECURITY punch-list item #1: the Storage SDK's e.message can embed the
      // request URL (i.e. the patient file PATH). Log ONLY e.code / e.statusCode,
      // NEVER e.message, so PHI never lands in Application Insights.
      const code = e && (e.code || e.statusCode);
      context.error("upstream error", { nonce: v.payload.n, code: code });
      const notFound =
        code === 404 || code === "ResourceNotFound" ||
        code === "ShareNotFound" || code === "ParentNotFound";
      return { status: notFound ? 404 : 502, body: notFound ? "not found" : "upstream error" };
    }
  }
});
