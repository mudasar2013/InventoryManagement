import type { NextAuthOptions } from "next-auth";
import AzureADProvider from "next-auth/providers/azure-ad";

/**
 * Delegated Microsoft sign-in. The signed-in user's own Graph access
 * token is what lib/sources/sharepoint-excel-source.ts uses to read the
 * inventory workbook — the app never has its own standing credential,
 * it can only see what the signed-in technician can see in SharePoint.
 *
 * Required environment variables (see .env.example):
 *   AZURE_AD_CLIENT_ID
 *   AZURE_AD_CLIENT_SECRET
 *   AZURE_AD_TENANT_ID
 *   NEXTAUTH_SECRET
 *   NEXTAUTH_URL          (e.g. http://localhost:43127 in dev)
 *
 * The app registration's redirect URI must be:
 *   {NEXTAUTH_URL}/api/auth/callback/azure-ad
 *
 * Delegated Graph permissions to request admin consent for:
 *   offline_access, openid, profile, email, Sites.Read.All
 * (Sites.Read.All lets a signed-in user read any SharePoint site they
 * already have access to. If your Azure AD admin prefers scoping this
 * to one specific site, swap it for Sites.Selected and grant that one
 * site explicitly — see the Graph docs for Sites.Selected.)
 */

const GRAPH_SCOPES = "offline_access openid profile email Sites.Read.All";

interface AzureTokenSet {
  access_token: string;
  refresh_token?: string;
  expires_in: number;
}

async function refreshAccessToken(refreshToken: string): Promise<AzureTokenSet> {
  const tenantId = process.env.AZURE_AD_TENANT_ID;
  const response = await fetch(
    `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`,
    {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: new URLSearchParams({
        client_id: process.env.AZURE_AD_CLIENT_ID ?? "",
        client_secret: process.env.AZURE_AD_CLIENT_SECRET ?? "",
        grant_type: "refresh_token",
        refresh_token: refreshToken,
        scope: GRAPH_SCOPES,
      }),
    },
  );

  if (!response.ok) {
    throw new Error(`Failed to refresh Azure AD token: ${response.status}`);
  }

  return response.json();
}

export const authOptions: NextAuthOptions = {
  providers: [
    AzureADProvider({
      clientId: process.env.AZURE_AD_CLIENT_ID ?? "",
      clientSecret: process.env.AZURE_AD_CLIENT_SECRET ?? "",
      tenantId: process.env.AZURE_AD_TENANT_ID,
      authorization: { params: { scope: `${GRAPH_SCOPES}` } },
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
