import type { DefaultSession } from "next-auth";

// Module augmentation so `session.accessToken` and `token.accessToken` are
// typed instead of `any` — these carry the Graph access token the
// SharePoint source reads with (see lib/auth/options.ts).
declare module "next-auth" {
  interface Session extends DefaultSession {
    accessToken?: string;
    error?: string;
  }
}

declare module "next-auth/jwt" {
  interface JWT {
    accessToken?: string;
    refreshToken?: string;
    accessTokenExpires?: number;
    error?: string;
  }
}
