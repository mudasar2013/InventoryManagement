"use client";

import { AlertTriangle, Pencil, Plus, Tag, Trash2, X } from "lucide-react";
import Link from "next/link";
import { useRouter } from "next/navigation";
import { useState } from "react";
import { AppHeader } from "@/components/AppHeader";

/**
 * Where app-native settings live — today just the tag vocabulary (see
 * lib/tags-store.ts). Category and Location aren't managed here: they
 * come straight from whatever each SharePoint source's own sheet
 * contains (see the "Category column" field on the Data sources page),
 * so there's no separate list to curate for those. Tags are different
 * — no sheet has a tags column, so the vocabulary itself has to live
 * somewhere, and this is that somewhere.
 */
export function SettingsView({
  tags,
  canManageTags,
}: {
  tags: string[];
  canManageTags: boolean;
}) {
  const router = useRouter();
  const [newTag, setNewTag] = useState("");
  const [adding, setAdding] = useState(false);
  const [renamingTag, setRenamingTag] = useState<string | null>(null);
  const [renameValue, setRenameValue] = useState("");
  const [busyTag, setBusyTag] = useState<string | null>(null);
  const [error, setError] = useState<string | null>(null);

  async function handleAdd(event: React.FormEvent) {
    event.preventDefault();
    if (!newTag.trim()) return;
    setAdding(true);
    setError(null);
    try {
      const response = await fetch("/api/tags", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ name: newTag.trim() }),
      });
      const payload = await response.json().catch(() => ({}));
      if (!response.ok) {
        throw new Error(payload.error ?? "Failed to add tag.");
      }
      setNewTag("");
      router.refresh();
    } catch (err) {
      setError(err instanceof Error ? err.message : "Failed to add tag.");
    } finally {
      setAdding(false);
    }
  }

  function startRename(tag: string) {
    setError(null);
    setRenamingTag(tag);
    setRenameValue(tag);
  }

  async function handleRename(event: React.FormEvent) {
    event.preventDefault();
    if (!renamingTag || !renameValue.trim()) return;
    setBusyTag(renamingTag);
    setError(null);
    try {
      const response = await fetch(`/api/tags/${encodeURIComponent(renamingTag)}`, {
        method: "PATCH",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ newName: renameValue.trim() }),
      });
      const payload = await response.json().catch(() => ({}));
      if (!response.ok) {
        throw new Error(payload.error ?? "Failed to rename tag.");
      }
      setRenamingTag(null);
      router.refresh();
    } catch (err) {
      setError(err instanceof Error ? err.message : "Failed to rename tag.");
    } finally {
      setBusyTag(null);
    }
  }

  async function handleDelete(tag: string) {
    setBusyTag(tag);
    setError(null);
    try {
      const response = await fetch(`/api/tags/${encodeURIComponent(tag)}`, {
        method: "DELETE",
      });
      if (!response.ok) {
        const payload = await response.json().catch(() => ({}));
        throw new Error(payload.error ?? "Failed to delete tag.");
      }
      router.refresh();
    } catch (err) {
      setError(err instanceof Error ? err.message : "Failed to delete tag.");
    } finally {
      setBusyTag(null);
    }
  }

  return (
    <main className="space-y-6">
      <AppHeader
        title="Settings"
        subtitle="Manage the tag vocabulary used to label parts across every source."
      />

      {error ? (
        <p className="inline-flex items-start gap-2 rounded-2xl border border-rose-200 bg-rose-50 px-3.5 py-3 text-sm leading-5 text-rose-900">
          <AlertTriangle className="mt-0.5 size-4 shrink-0" />
          {error}
        </p>
      ) : null}

      <section className="space-y-3">
        <h2 className="text-sm font-semibold uppercase tracking-wide text-stone-500">Tags</h2>

        {!canManageTags ? (
          <div className="rounded-2xl border border-dashed border-stone-300 bg-white p-4">
            <p className="text-sm leading-6 text-stone-600">
              Tags need somewhere to save — this project doesn&apos;t have a Redis
              store yet. Add a Redis integration in the Vercel dashboard (Project →
              Storage → Marketplace Database Providers), then this page will let you
              add and manage tags.
            </p>
          </div>
        ) : (
          <>
            {tags.length === 0 ? (
              <p className="rounded-2xl border border-dashed border-stone-300 bg-white px-4 py-6 text-center text-sm text-stone-500">
                No tags yet — add one below (e.g. &quot;supplies&quot;).
              </p>
            ) : (
              <ul className="space-y-2">
                {tags.map((tag) => (
                  <li
                    key={tag}
                    className="flex items-center justify-between gap-3 rounded-2xl border border-stone-200 bg-white p-3.5"
                  >
                    {renamingTag === tag ? (
                      <form onSubmit={handleRename} className="flex flex-1 items-center gap-2">
                        <input
                          autoFocus
                          value={renameValue}
                          onChange={(event) => setRenameValue(event.target.value)}
                          className="h-9 flex-1 rounded-lg border border-stone-200 px-2.5 text-sm outline-none focus:border-amber-400 focus:ring-2 focus:ring-amber-100"
                        />
                        <button
                          type="submit"
                          disabled={busyTag === tag}
                          className="rounded-lg bg-stone-900 px-2.5 py-1.5 text-xs font-semibold text-white disabled:opacity-60"
                        >
                          Save
                        </button>
                        <button
                          type="button"
                          onClick={() => setRenamingTag(null)}
                          className="rounded-full p-1 text-stone-400"
                          aria-label="Cancel rename"
                        >
                          <X className="size-4" />
                        </button>
                      </form>
                    ) : (
                      <>
                        <span className="inline-flex items-center gap-2 text-sm font-medium text-stone-900">
                          <Tag className="size-3.5 text-amber-700" />
                          {tag}
                        </span>
                        <div className="flex items-center gap-1">
                          <button
                            type="button"
                            onClick={() => startRename(tag)}
                            disabled={busyTag === tag}
                            aria-label={`Rename ${tag}`}
                            className="rounded-full p-1.5 text-stone-400 hover:bg-amber-50 hover:text-amber-700 disabled:opacity-50"
                          >
                            <Pencil className="size-3.5" />
                          </button>
                          <button
                            type="button"
                            onClick={() => handleDelete(tag)}
                            disabled={busyTag === tag}
                            aria-label={`Delete ${tag}`}
                            className="rounded-full p-1.5 text-stone-400 hover:bg-rose-50 hover:text-rose-600 disabled:opacity-50"
                          >
                            <Trash2 className="size-3.5" />
                          </button>
                        </div>
                      </>
                    )}
                  </li>
                ))}
              </ul>
            )}

            <form onSubmit={handleAdd} className="flex items-center gap-2">
              <input
                value={newTag}
                onChange={(event) => setNewTag(event.target.value)}
                placeholder="New tag, e.g. supplies"
                className="h-11 flex-1 rounded-xl border border-stone-200 bg-white px-3 text-sm outline-none focus:border-amber-400 focus:ring-2 focus:ring-amber-100"
              />
              <button
                type="submit"
                disabled={adding || !newTag.trim()}
                className="inline-flex h-11 shrink-0 items-center gap-1.5 rounded-xl bg-stone-900 px-4 text-sm font-semibold text-white disabled:opacity-60"
              >
                <Plus className="size-4" />
                Add
              </button>
            </form>
          </>
        )}
      </section>

      <section className="rounded-2xl border border-stone-200 bg-white p-4">
        <h2 className="text-sm font-semibold text-stone-900">Category &amp; location filters</h2>
        <p className="mt-1.5 text-sm leading-6 text-stone-600">
          Category and bin location filters on the Parts page read directly from
          each source&apos;s own sheet — there&apos;s no separate list to manage
          here. To make a source&apos;s Category column show up as a filter, set
          its &quot;Category column&quot; on the{" "}
          <Link href="/sources" className="font-semibold text-amber-700 underline">
            Data sources
          </Link>{" "}
          page.
        </p>
      </section>
    </main>
  );
}
