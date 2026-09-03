import { X509Certificate, createHash, randomUUID } from "node:crypto";
import { SignJWT, importPKCS8 } from "jose";

/**
 * Certificate-based client authentication for the Microsoft identity
 * platform's token endpoint — used instead of a client secret. Some
 * tenants block client secrets by policy (see the README's "Sign-in
 * setup" section), and Microsoft is pushing certificates as the
 * replacement.
 *
 * This is NOT generic OAuth `private_key_jwt` (which openid-client, the
 * library NextAuth uses internally, already supports out of the box by
 * setting `token_endpoint_auth_method` on the provider). Azure AD's
 * token endpoint has its own, stricter requirements verified against
 * https://learn.microsoft.com/entra/identity-platform/certificate-credentials
 * while building this: `alg` must be PS256 (not the RS256 most generic
 * examples use), and the JWS header must carry `x5t#S256` — the
 * base64url SHA-256 thumbprint of the certificate's DER encoding —
 * which openid-client's built-in assertion builder does not set (it
 * only sets `kid`). That mismatch fails silently at Azure's token
 * endpoint rather than at compile time, so this is hand-built and
 * covered by certificate.test.ts instead of delegated to the library.
 */

export interface CertificateCredential {
  /** PKCS8 PEM private key, e.g. "-----BEGIN PRIVATE KEY-----...". */
  privateKeyPem: string;
  /** X.509 certificate PEM — the public half of privateKeyPem, and the
   *  same file uploaded to the Azure AD app registration's
   *  Certificates & secrets. */
  certificatePem: string;
}

/**
 * Reads AZURE_AD_CERT_PRIVATE_KEY_BASE64 / AZURE_AD_CERT_BASE64 — each
 * the base64 encoding of a PEM file, since env vars can't reliably hold
 * literal newlines — and decodes them back to PEM. `npm run generate-cert`
 * prints both values ready to paste into .env.local.
 */
export function loadCertificateCredential(): CertificateCredential {
  const privateKeyBase64 = process.env.AZURE_AD_CERT_PRIVATE_KEY_BASE64;
  const certificateBase64 = process.env.AZURE_AD_CERT_BASE64;

  if (!privateKeyBase64 || !certificateBase64) {
    throw new Error(
      "Missing AZURE_AD_CERT_PRIVATE_KEY_BASE64 and/or AZURE_AD_CERT_BASE64. " +
        "Run `npm run generate-cert` and copy its output into .env.local " +
        '(see README.md "Sign-in setup").',
    );
  }

  return {
    privateKeyPem: Buffer.from(privateKeyBase64, "base64").toString("utf8"),
    certificatePem: Buffer.from(certificateBase64, "base64").toString("utf8"),
  };
}

/**
 * base64url SHA-256 thumbprint of the certificate's DER encoding — the
 * `x5t#S256` header value Azure AD matches against the certificate
 * uploaded to the app registration.
 */
export function computeCertificateThumbprint(certificatePem: string): string {
  const cert = new X509Certificate(certificatePem);
  return createHash("sha256").update(cert.raw).digest("base64url");
}

/**
 * Builds one signed JWT client assertion for a single token-endpoint
 * request. Short-lived and single-use by design (fresh `jti`, 5 minute
 * expiry) — never cache or reuse the return value across requests.
 */
export async function buildClientAssertion({
  clientId,
  tokenEndpoint,
  credential,
}: {
  clientId: string;
  tokenEndpoint: string;
  credential: CertificateCredential;
}): Promise<string> {
  const privateKey = await importPKCS8(credential.privateKeyPem, "PS256");
  const x5tS256 = computeCertificateThumbprint(credential.certificatePem);

  return new SignJWT({})
    .setProtectedHeader({ alg: "PS256", typ: "JWT", "x5t#S256": x5tS256 })
    .setIssuer(clientId)
    .setSubject(clientId)
    .setAudience(tokenEndpoint)
    .setJti(randomUUID())
    .setIssuedAt()
    .setNotBefore("0s")
    .setExpirationTime("5m")
    .sign(privateKey);
}

/** The Microsoft identity platform v2 token endpoint for a tenant — both
 *  the URL requests are POSTed to and the `aud` claim of the assertion. */
export function getTokenEndpoint(tenantId: string): string {
  return `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`;
}
