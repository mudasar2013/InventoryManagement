import type { NextAuthOptions } from "next-auth";
import AzureADProvider from "next-auth/providers/azure-ad";
import {
  buildClientAssertion,
  getTokenEndpoint,
  loadCertificateCredential,
} from "./certificate";

/**
 * Delegated Microsoft sign-in. The signed-in user's own Graph access
 * token is what lib/sources/sharepoint-excel-source.ts uses to read the
 * inventory workbook — the app never has its own standing credential,
 * it can only see what the signed-in technician can see in SharePoint.
 *
 * This app authenticates itself to Azure AD with a CERTIFICATE, not a
 * client secret — some tenants block client secrets by a tenant-wide
 * app management policy (see the README's "Sign-in setup" section), and
 * Microsoft is steering everyone toward certificates anyway. See
 * lib/auth/certificate.ts for why that couldn't just be
 * `token_endpoint_auth_method: "private_key_jwt"` on the provider (the
 * generic OAuth version openid-client supports out of the box doesn't
 * satisfy Azure AD's specific header requirements) and why the token
 * exchange below is hand-built instead.
 *
 * Required environment variables (see .env.example):
 *   AZURE_AD_CLIENT_ID
 *   AZURE_AD_TENANT_ID
 *   AZURE_AD_CERT_PRIVATE_KEY_BASE64   (from `npm run generate-cert`)
 *   AZURE_AD_CERT_BASE64               (from `npm run generate-cert`)
 *   NEXTAUTH_SECRET
 *   NEXTAUTH_URL          (e.g. http://localhost:43127 in dev)
 *
 * The app registration's redirect URI must be:
 *   {NEXTAUTH_URL}/api/auth/callback/azure-ad
 *
 * Delegated Graph permissions to request admin consent for:
 *   offline_access, openid, profile, email, Sites.Selected
 * (Sites.Selected grants this app NO access to any SharePoint site by
 * default — narrower than Sites.Read.All/Sites.ReadWrite.All, which
 * hand it read (or read+write) access to every site the signed-in user
 * can already reach. With Sites.Selected, an admin must separately
 * grant this app access to each specific site that holds an inventory
 * workbook, with the "write" role (needed since the app writes part
 * updates/additions back into the workbook — see
 * lib/sources/sharepoint-excel-source.ts's updatePartInWorkbook /
 * addPartToWorkbook) — see the README's "Sign-in setup" section for the
 * exact PnP PowerShell command. That per-site grant is the finest
 * scoping Graph offers; there's no permission model that scopes down to
 * one specific file while leaving the rest of a site untouched — this
 * gets close by only ever touching the site(s) you explicitly grant,
 * not the whole tenant. After changing this scope in the app
 * registration's API permissions and granting admin consent, everyone
 * needs to sign out and back in — a stored refresh token from before
 * the change won't carry the new scope.)
 */

const GRAPH_SCOPES = "offline_access openid profile email Sites.Selected";

interface AzureTokenResponse {
  access_token: string;
  refresh_token?: string;
  expires_in: number;
  error?: string;
  error_description?: string;
  // Azure AD's token response carries more fields (id_token, token_type,
  // scope, ext_expires_in, ...) that next-auth's TokenSet reads from
  // `token.request`'s return value — allow them through untyped rather
  // than dropping them.
  [key: string]: unknown;
}

async function requestToken(
  tenantId: string,
  body: URLSearchParams,
): Promise<AzureTokenResponse> {
  const response = await fetch(getTokenEndpoint(tenantId), {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body,
  });

  const payload = (await response.json()) as AzureTokenResponse;
  if (!response.ok) {
    throw new Error(
      `Azure AD token request failed: ${payload.error ?? response.status} ${
        payload.error_description ?? ""
      }`.trim(),
    );
  }
  return payload;
}

async function refreshAccessToken(refreshToken: string): Promise<AzureTokenResponse> {
  const clientId = process.env.AZURE_AD_CLIENT_ID ?? "";
  const tenantId = process.env.AZURE_AD_TENANT_ID ?? "";
  const assertion = await buildClientAssertion({
    clientId,
    tokenEndpoint: getTokenEndpoint(tenantId),
    credential: loadCertificateCredential(),
  });

  return requestToken(
    tenantId,
    new URLSearchParams({
      client_id: clientId,
      grant_type: "refresh_token",
      refresh_token: refreshToken,
      scope: GRAPH_SCOPES,
      client_assertion_type: "urn:ietf:params:oauth:client-assertion-type:jwt-bearer",
      client_assertion: assertion,
    }),
  );
}

export const authOptions: NextAuthOptions = {
  providers: [
    AzureADProvider({
      clientId: process.env.AZURE_AD_CLIENT_ID ?? "",
      // Required by next-auth's OAuthUserConfig type, but never actually
      // used — the token.request override below authenticates with a
      // certificate instead and never reaches the code path that would
      // read this.
      clientSecret: "unused-app-authenticates-with-a-certificate-instead",
      tenantId: process.env.AZURE_AD_TENANT_ID,
      authorization: { params: { scope: GRAPH_SCOPES } },
      checks: ["pkce", "state"],
      client: { token_endpoint_auth_method: "none" },
      token: {
        async request({ provider, params, checks }) {
          // openid-client's built-in state check is bypassed by
          // providing a custom token.request — restore it here so a
          // forged callback can't complete sign-in (CSRF protection).
          if (checks.state !== undefined && checks.state !== params.state) {
            throw new Error("OAuth state mismatch on Azure AD callback.");
          }

          const clientId = provider.clientId as string;
          const tenantId = process.env.AZURE_AD_TENANT_ID ?? "";
          const assertion = await buildClientAssertion({
            clientId,
            tokenEndpoint: getTokenEndpoint(tenantId),
            credential: loadCertificateCredential(),
          });

          const body = new URLSearchParams({
            client_id: clientId,
            grant_type: "authorization_code",
            code: params.code as string,
            redirect_uri: provider.callbackUrl,
            client_assertion_type:
              "urn:ietf:params:oauth:client-assertion-type:jwt-bearer",
            client_assertion: assertion,
          });
          if (checks.code_verifier) {
            body.set("code_verifier", checks.code_verifier as string);
          }

          const tokens = await requestToken(tenantId, body);
          return { tokens };
        },
      },
    }),
  ],
  callbacks: {
    async jwt({ token, account }) {
      // Initial sign-in: Azure AD hands back the Graph access/refresh
      // tokens on the `account` object. Stash them on the JWT so every
      // later request has something to read the workbook with.
      if (account) {
        return {
          ...token,
          accessToken: account.access_token,
          refreshToken: account.refresh_token,
          accessTokenExpires: (account.expires_at ?? 0) * 1000,
        };
      }

      // Still valid — nothing to do.
      if (Date.now() < (token.accessTokenExpires as number)) {
        return token;
      }

      // Expired — refresh it. If refresh fails (revoked, expired refresh
      // token), mark the token so pages can detect it and prompt a
      // re-sign-in instead of silently reading stale/no SharePoint data.
      try {
        const refreshed = await refreshAccessToken(token.refreshToken as string);
        return {
          ...token,
          accessToken: refreshed.access_token,
          refreshToken: refreshed.refresh_token ?? token.refreshToken,
          accessTokenExpires: Date.now() + refreshed.expires_in * 1000,
          error: undefined,
        };
      } catch {
        return { ...token, error: "RefreshAccessTokenError" as const };
      }
    },
    async session({ session, token }) {
      return {
        ...session,
        accessToken: token.accessToken as string | undefined,
        error: token.error as string | undefined,
      };
    },
  },
};
