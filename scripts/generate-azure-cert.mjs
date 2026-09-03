#!/usr/bin/env node
// Generates a self-signed certificate for Azure AD certificate-based
// client authentication (see README.md "Sign-in setup"). Run this on
// your own machine — the private key it creates never needs to leave
// it, and never needs to pass through anyone else.
//
// Usage:
//   npm run generate-cert
//
// Output:
//   certs/azure-ad-private-key.pem   <- keep secret, never commit, never share
//   certs/azure-ad-certificate.pem   <- upload this one to Azure AD
//   Two base64 values printed to the terminal for .env.local
//
// The certs/ directory is already in .gitignore.

import { mkdirSync, writeFileSync, existsSync } from "node:fs";
import { generate } from "selfsigned";

const OUT_DIR = "certs";
const PRIVATE_KEY_PATH = `${OUT_DIR}/azure-ad-private-key.pem`;
const CERTIFICATE_PATH = `${OUT_DIR}/azure-ad-certificate.pem`;

async function main() {
  if (existsSync(PRIVATE_KEY_PATH) || existsSync(CERTIFICATE_PATH)) {
    console.error(
      `${PRIVATE_KEY_PATH} or ${CERTIFICATE_PATH} already exists.\n` +
        "Delete them first if you really want to generate a new certificate " +
        "(you'll need to upload the new one to Azure AD and update .env.local too).",
    );
    process.exit(1);
  }

  const notBefore = new Date();
  const notAfter = new Date(notBefore);
  notAfter.setFullYear(notAfter.getFullYear() + 2);

  const pems = await generate(
    [{ name: "commonName", value: "parts-inventory-sharepoint-sync" }],
    {
      keySize: 2048,
      algorithm: "sha256",
      notBeforeDate: notBefore,
      notAfterDate: notAfter,
    },
  );

  mkdirSync(OUT_DIR, { recursive: true });
  writeFileSync(PRIVATE_KEY_PATH, pems.private, { mode: 0o600 });
  writeFileSync(CERTIFICATE_PATH, pems.cert);

  const privateKeyBase64 = Buffer.from(pems.private).toString("base64");
  const certificateBase64 = Buffer.from(pems.cert).toString("base64");

  console.log(`Wrote ${PRIVATE_KEY_PATH} and ${CERTIFICATE_PATH}.`);
  console.log(`Valid until ${notAfter.toISOString().slice(0, 10)}.\n`);
  console.log("Next steps:\n");
  console.log(
    `1. In the Azure Portal, go to your app registration -> Certificates & secrets\n` +
      `   -> Certificates tab -> Upload certificate, and upload ${CERTIFICATE_PATH}.\n`,
  );
  console.log("2. Add these two lines to .env.local:\n");
  console.log(`AZURE_AD_CERT_PRIVATE_KEY_BASE64=${privateKeyBase64}`);
  console.log(`AZURE_AD_CERT_BASE64=${certificateBase64}`);
  console.log(
    "\n(Remove AZURE_AD_CLIENT_SECRET from .env.local if it's still there — it's not used anymore.)",
  );
}

main().catch((error) => {
  console.error(error);
  process.exit(1);
});
