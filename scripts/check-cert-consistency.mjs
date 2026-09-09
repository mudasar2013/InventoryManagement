#!/usr/bin/env node
// Diagnoses AADSTS700027 ("client assertion contains an invalid signature")
// for this app's certificate-based Azure AD sign-in. Run this from the
// project root, where it can find .env.local and certs/.
//
// Usage:
//   node scripts/check-cert-consistency.mjs
//
// SAFE TO SHARE: this script never prints the private key, the certificate
// body, or the raw base64 env values. It only prints thumbprints (meant to
// be compared against the Azure Portal), dates, IDs, and PASS/FAIL results.

import { readFileSync, existsSync } from "node:fs";
import { X509Certificate, createHash } from "node:crypto";
import { importPKCS8, importX509, SignJWT, jwtVerify } from "jose";

function parseEnvLocal(path) {
  if (!existsSync(path)) {
    console.error(`Missing ${path}. Run this from the project root.`);
    process.exit(1);
  }
  const out = {};
  for (const line of readFileSync(path, "utf8").split("\n")) {
    const trimmed = line.trim();
    if (!trimmed || trimmed.startsWith("#")) continue;
    const eq = trimmed.indexOf("=");
    if (eq === -1) continue;
    const key = trimmed.slice(0, eq).trim();
    let value = trimmed.slice(eq + 1).trim();
    if (
      (value.startsWith('"') && value.endsWith('"')) ||
      (value.startsWith("'") && value.endsWith("'"))
    ) {
      value = value.slice(1, -1);
    }
    out[key] = value;
  }
  return out;
}

function thumbprintSHA256Base64Url(certificatePem) {
  const cert = new X509Certificate(certificatePem);
  return createHash("sha256").update(cert.raw).digest("base64url");
}

async function main() {
  console.log("=== Azure AD certificate consistency check ===\n");

  const env = parseEnvLocal(".env.local");
  const clientId = env.AZURE_AD_CLIENT_ID ?? "(not set)";
  const tenantId = env.AZURE_AD_TENANT_ID ?? "(not set)";
  console.log(`AZURE_AD_CLIENT_ID (from .env.local): ${clientId}`);
  console.log(`AZURE_AD_TENANT_ID (from .env.local): ${tenantId}`);
  console.log(
    "-> Compare AZURE_AD_CLIENT_ID above against the Azure Portal app registration's\n" +
      "   'Application (client) ID', and against the app id in the failing sign-in log.\n" +
      "   A mismatch here means .env.local points at the wrong app registration\n" +
      "   entirely (a different app than the one you uploaded the certificate to).\n",
  );

  const certBase64 = env.AZURE_AD_CERT_BASE64;
  const keyBase64 = env.AZURE_AD_CERT_PRIVATE_KEY_BASE64;
  if (!certBase64 || !keyBase64) {
    console.error(
      "AZURE_AD_CERT_BASE64 and/or AZURE_AD_CERT_PRIVATE_KEY_BASE64 missing from .env.local.",
    );
    process.exit(1);
  }

  const certificatePem = Buffer.from(certBase64, "base64").toString("utf8");
  const privateKeyPem = Buffer.from(keyBase64, "base64").toString("utf8");

  // Compare against the on-disk files in certs/, if present.
  const diskCertPath = "certs/azure-ad-certificate.pem";
  const diskKeyPath = "certs/azure-ad-private-key.pem";
  if (existsSync(diskCertPath)) {
    const diskCert = readFileSync(diskCertPath, "utf8").trim();
    const match = diskCert === certificatePem.trim();
    console.log(
      `certs/azure-ad-certificate.pem matches AZURE_AD_CERT_BASE64: ${match ? "YES" : "NO — these are DIFFERENT certificates"}`,
    );
  } else {
    console.log(`(${diskCertPath} not found on disk — skipping file comparison)`);
  }
  if (existsSync(diskKeyPath)) {
    const diskKey = readFileSync(diskKeyPath, "utf8").trim();
    const match = diskKey === privateKeyPem.trim();
    console.log(
      `certs/azure-ad-private-key.pem matches AZURE_AD_CERT_PRIVATE_KEY_BASE64: ${match ? "YES" : "NO — these are DIFFERENT keys"}`,
    );
  } else {
    console.log(`(${diskKeyPath} not found on disk — skipping file comparison)`);
  }

  // Certificate identity + validity window.
  const cert = new X509Certificate(certificatePem);
  const now = new Date();
  const validFrom = new Date(cert.validFrom);
  const validTo = new Date(cert.validTo);
  const inWindow = now >= validFrom && now <= validTo;
  console.log(`\nCertificate subject: ${cert.subject.split("\n").join(", ")}`);
  console.log(`Certificate valid from: ${validFrom.toISOString()}`);
  console.log(`Certificate valid to:   ${validTo.toISOString()}`);
  console.log(`Certificate currently within its validity window: ${inWindow ? "YES" : "NO — this is likely the problem"}`);

  const sha1Thumbprint = cert.fingerprint.replace(/:/g, "").toUpperCase();
  const sha256HexThumbprint = cert.fingerprint256.replace(/:/g, "").toUpperCase();
  const sha256B64Url = thumbprintSHA256Base64Url(certificatePem);
  console.log(`\nSHA-1 thumbprint (hex, matches Azure Portal's "Thumbprint" column): ${sha1Thumbprint}`);
  console.log(`SHA-256 thumbprint (hex): ${sha256HexThumbprint}`);
  console.log(`SHA-256 thumbprint (base64url, this is the x5t#S256 the app sends): ${sha256B64Url}`);
  console.log(
    "-> Go to Azure Portal -> your app registration -> Certificates & secrets -> Certificates,\n" +
      "   and check whether the SHA-1 thumbprint above appears in that list.\n" +
      "   If it does NOT appear there, that's the bug: the certificate this app is signing\n" +
      "   with was never uploaded (or a different/old certificate is uploaded instead).\n" +
      "   Fix: upload certs/azure-ad-certificate.pem there now (no need to regenerate).\n",
  );

  // Prove the private key actually pairs with this certificate, using the
  // exact same PS256 + x5t#S256 construction lib/auth/certificate.ts uses.
  try {
    const privateKey = await importPKCS8(privateKeyPem, "PS256");
    const jwt = await new SignJWT({})
      .setProtectedHeader({ alg: "PS256", typ: "JWT", "x5t#S256": sha256B64Url })
      .setIssuer("diagnostic")
      .setSubject("diagnostic")
      .setAudience("diagnostic")
      .setIssuedAt()
      .setExpirationTime("1m")
      .sign(privateKey);

    const publicKey = await importX509(certificatePem, "PS256");
    await jwtVerify(jwt, publicKey);
    console.log("Private key cryptographically matches this certificate: YES");
  } catch (error) {
    console.log("Private key cryptographically matches this certificate: NO");
    console.log(
      "-> This means AZURE_AD_CERT_BASE64 and AZURE_AD_CERT_PRIVATE_KEY_BASE64 in .env.local\n" +
        "   are NOT a matching pair (e.g. a private key from one generate-cert run paired\n" +
        "   with a certificate from a different run, or a copy/paste mistake). This alone\n" +
        "   is enough to cause AADSTS700027. Fix: delete certs/*.pem, run\n" +
        "   `npm run generate-cert` again, upload the new certificate to Azure AD, and\n" +
        "   replace BOTH lines in .env.local with the newly printed values.",
    );
    console.log(`   (verify error: ${error.message})`);
  }

  console.log("\n=== End of report — no secret values were printed above. ===");
}

main().catch((error) => {
  console.error("Diagnostic script failed:", error);
  process.exit(1);
});
