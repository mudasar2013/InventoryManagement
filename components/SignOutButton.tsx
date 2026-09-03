"use client";

import { signOut } from "next-auth/react";

export function SignOutButton() {
  return (
    <button
      type="button"
      onClick={() => signOut()}
      className="shrink-0 font-semibold text-stone-600 underline-offset-2 hover:underline"
    >
      Sign out
    </button>
  );
}
