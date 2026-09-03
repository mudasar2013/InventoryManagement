import { withAuth } from "next-auth/middleware";

// Next.js 16 renamed the middleware.ts convention to proxy.ts (same
// request/response contract, see node_modules/next/dist/docs/01-app/
// 03-api-reference/03-file-conventions/proxy.md). next-auth's withAuth()
// still returns a standard middleware-shaped function, so it plugs in
// here unchanged — only the file name and this export changed.
//
// This gates the whole shop-floor app behind Microsoft sign-in: every
// technician needs their own Graph-delegated session anyway for the
// SharePoint inventory source to read anything on their behalf, so
// there's no useful "signed out" view of the app to show.
export const proxy = withAuth({
  callbacks: {
    authorized: ({ token }) => Boolean(token),
  },
});

export const config = {
  matcher: ["/((?!api/auth|_next/static|_next/image|favicon.ico).*)"],
};
