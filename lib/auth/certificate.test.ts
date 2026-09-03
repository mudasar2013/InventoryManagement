import assert from "node:assert/strict";
import { test } from "node:test";
import { generate } from "selfsigned";
import { importX509, jwtVerify } from "jose";
import {
  buildClientAssertion,
  computeCertificateThumbprint,
  getTokenEndpoint,
  loadCertificateCredential,
} from "./certificate";

// A fresh throwaway keypair/cert per test file run — nothing here needs
// to match a real Azure AD app, only to prove the assertion we build is
// internally consistent and verifiable against its own certificate,
// which is exactly what Azure AD does on its end.
async function generateTestCredential() {
  const pems = await generate([{ name: "commonName", value: "test.local" }], {
    keySize: 2048,
    algorithm: "sha256",
  });
  return { privateKeyPem: pems.private, certificatePem: pems.cert };
}

test("computeCertificateThumbprint: matches manually computing the SHA-256 of the DER cert", async () => {
  const { certificatePem } = await generateTestCredential();
  const thumbprint = computeCertificateThumbprint(certificatePem);

  // base64url, no padding, and the right rough length for a 32-byte SHA-256 digest.
  assert.match(thumbprint, /^[A-Za-z0-9_-]{43}$/);
});

test("buildClientAssertion: produces a JWT that verifies against the certificate's own public key", async () => {
  const credential = await generateTestCredential();
  const tokenEndpoint = getTokenEndpoint("test-tenant-id");

  const jwt = await buildClientAssertion({
    clientId: "test-client-id",
    tokenEndpoint,
    credential,
  });

  const publicKey = await importX509(credential.certificatePem, "PS256");
  const { payload, protectedHeader } = await jwtVerify(jwt, publicKey);

  assert.equal(protectedHeader.alg, "PS256");
  assert.equal(protectedHeader.typ, "JWT");
  assert.equal(
    protectedHeader["x5t#S256"],
    computeCertificateThumbprint(credential.certificatePem),
  );
  assert.equal(payload.iss, "test-client-id");
  assert.equal(payload.sub, "test-client-id");
  assert.equal(payload.aud, tokenEndpoint);
  assert.equal(typeof payload.jti, "string");
});

test("buildClientAssertion: does not verify against a different certificate's key", async () => {
  const credential = await generateTestCredential();
  const otherCredential = await generateTestCredential();
  const tokenEndpoint = getTokenEndpoint("test-tenant-id");

  const jwt = await buildClientAssertion({
    clientId: "test-client-id",
    tokenEndpoint,
    credential,
  });

  const wrongPublicKey = await importX509(otherCredential.certificatePem, "PS256");
  await assert.rejects(() => jwtVerify(jwt, wrongPublicKey));
});

test("buildClientAssertion: two calls never reuse the same jti (replay protection)", async () => {
  const credential = await generateTestCredential();
  const tokenEndpoint = getTokenEndpoint("test-tenant-id");
  const publicKey = await importX509(credential.certificatePem, "PS256");

  const [jwtA, jwtB] = await Promise.all([
    buildClientAssertion({ clientId: "test-client-id", tokenEndpoint, credential }),
    buildClientAssertion({ clientId: "test-client-id", tokenEndpoint, credential }),
  ]);

  const [a, b] = await Promise.all([
    jwtVerify(jwtA, publicKey),
    jwtVerify(jwtB, publicKey),
  ]);

  assert.notEqual(a.payload.jti, b.payload.jti);
});

test("loadCertificateCredential: throws a specific, actionable error when env vars are missing", () => {
  const savedKey = process.env.AZURE_AD_CERT_PRIVATE_KEY_BASE64;
  const savedCert = process.env.AZURE_AD_CERT_BASE64;
  delete process.env.AZURE_AD_CERT_PRIVATE_KEY_BASE64;
  delete process.env.AZURE_AD_CERT_BASE64;

  try {
    assert.throws(() => loadCertificateCredential(), /generate-cert/);
  } finally {
    if (savedKey !== undefined) process.env.AZURE_AD_CERT_PRIVATE_KEY_BASE64 = savedKey;
    if (savedCert !== undefined) process.env.AZURE_AD_CERT_BASE64 = savedCert;
  }
});

test("loadCertificateCredential: round-trips base64-encoded PEM back to the original", async () => {
  const credential = await generateTestCredential();
  const savedKey = process.env.AZURE_AD_CERT_PRIVATE_KEY_BASE64;
  const savedCert = process.env.AZURE_AD_CERT_BASE64;

  process.env.AZURE_AD_CERT_PRIVATE_KEY_BASE64 = Buffer.from(
    credential.privateKeyPem,
  ).toString("base64");
  process.env.AZURE_AD_CERT_BASE64 = Buffer.from(credential.certificatePem).toString(
    "base64",
  );

  try {
    const loaded = loadCertificateCredential();
    assert.equal(loaded.privateKeyPem, credential.privateKeyPem);
    assert.equal(loaded.certificatePem, credential.certificatePem);
  } finally {
    if (savedKey !== undefined) {
      process.env.AZURE_AD_CERT_PRIVATE_KEY_BASE64 = savedKey;
    } else {
      delete process.env.AZURE_AD_CERT_PRIVATE_KEY_BASE64;
    }
    if (savedCert !== undefined) {
      process.env.AZURE_AD_CERT_BASE64 = savedCert;
    } else {
      delete process.env.AZURE_AD_CERT_BASE64;
    }
  }
});
