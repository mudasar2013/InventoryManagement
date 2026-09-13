import { getRedis, isRedisConfigured } from "./redis";

/**
 * App-managed tags (e.g. "supplies") and which parts carry them —
 * neither lives in any SharePoint sheet, so both are stored here
 * instead, in the same Redis store used for source configs (see
 * lib/sources/sharepoint-source-store.ts). Two keys: the tag
 * vocabulary itself (managed from the Settings page — see
 * app/settings/page.tsx), and a part_number -> tags[] map recording
 * which parts carry which tags.
 *
 * Keyed by part_number rather than a part's id or (sourceId,
 * part_number) pair because part_number is the one identity that's
 * stable across a source re-fetch and shared across sources — exactly
 * the identity lib/sources/merge.ts already uses to dedupe the same
 * part reported by more than one source. A tag assigned to a part
 * survives that part's row being re-read (even re-ordered, even
 * reported by a different source than before) as long as its part
 * number doesn't change; renaming a part's part_number is treated the
 * same as anywhere else in this app — the old identity's tags don't
 * automatically follow (nothing to key them by once the part_number
 * itself is gone).
 */
const TAG_LIST_KEY = "inventory:tags";
const PART_TAGS_KEY = "inventory:part-tags";

export const isTagStoreConfigured = isRedisConfigured;

function requireRedis() {
  const redis = getRedis();
  if (!redis) {
    throw new Error(
      "No Redis store configured. Add a Redis integration to this Vercel project " +
        "(Storage tab → Marketplace Database Providers) to enable tags.",
    );
  }
  return redis;
}

/** Returns [] (never throws) when Redis isn't configured — same
 *  fail-open philosophy as listStoredSharePointSources: a missing store
 *  just means no tags exist yet, not a broken page. */
export async function listTags(): Promise<string[]> {
  const redis = getRedis();
  if (!redis) {
    return [];
  }
  const tags = await redis.get<string[]>(TAG_LIST_KEY);
  return tags ?? [];
}

/** Adds a tag to the vocabulary, case-insensitively deduped against
 *  what's already there (adding "Supplies" when "supplies" exists is a
 *  no-op, not a second near-duplicate entry). Returns the full,
 *  alphabetized list either way. */
export async function addTag(name: string): Promise<string[]> {
  const redis = requireRedis();
  const trimmed = name.trim();
  if (!trimmed) {
    throw new Error("Tag name can't be empty.");
  }
  const current = await listTags();
  if (current.some((tag) => tag.toLowerCase() === trimmed.toLowerCase())) {
    return current;
  }
  const next = [...current, trimmed].sort((a, b) => a.localeCompare(b));
  await redis.set(TAG_LIST_KEY, next);
  return next;
}

/** Renames a tag in the vocabulary and in every part currently carrying
 *  it, so a rename doesn't silently orphan existing assignments. */
export async function renameTag(oldName: string, newName: string): Promise<string[]> {
  const redis = requireRedis();
  const trimmed = newName.trim();
  if (!trimmed) {
    throw new Error("Tag name can't be empty.");
  }
  const current = await listTags();
  if (!current.includes(oldName)) {
    throw new Error("That tag no longer exists — it may have been removed already.");
  }
  const next = current
    .map((tag) => (tag === oldName ? trimmed : tag))
    .sort((a, b) => a.localeCompare(b));
  await redis.set(TAG_LIST_KEY, next);

  const allPartTags = await getAllPartTags();
  let changed = false;
  for (const partNumber of Object.keys(allPartTags)) {
    const idx = allPartTags[partNumber].indexOf(oldName);
    if (idx !== -1) {
      allPartTags[partNumber][idx] = trimmed;
      changed = true;
    }
  }
  if (changed) {
    await redis.set(PART_TAGS_KEY, allPartTags);
  }
  return next;
}

/** Removes a tag from the vocabulary and un-assigns it from every part
 *  that had it. */
export async function deleteTag(name: string): Promise<string[]> {
  const redis = requireRedis();
  const current = await listTags();
  const next = current.filter((tag) => tag !== name);
  await redis.set(TAG_LIST_KEY, next);

  const allPartTags = await getAllPartTags();
  let changed = false;
  for (const partNumber of Object.keys(allPartTags)) {
    if (allPartTags[partNumber].includes(name)) {
      allPartTags[partNumber] = allPartTags[partNumber].filter((tag) => tag !== name);
      changed = true;
    }
  }
  if (changed) {
    await redis.set(PART_TAGS_KEY, allPartTags);
  }
  return next;
}

async function getAllPartTags(): Promise<Record<string, string[]>> {
  const redis = getRedis();
  if (!redis) {
    return {};
  }
  const data = await redis.get<Record<string, string[]>>(PART_TAGS_KEY);
  return data ?? {};
}

/** Looks up tags for a batch of part numbers in one Redis read — used
 *  by loadInventory() to attach `tags` to every merged Part in one
 *  shot rather than one lookup per part. Part numbers with no tags are
 *  simply absent from the result (not present with an empty array). */
export async function getTagsForParts(
  partNumbers: string[],
): Promise<Record<string, string[]>> {
  if (partNumbers.length === 0) {
    return {};
  }
  const all = await getAllPartTags();
  const wanted = new Set(partNumbers);
  const result: Record<string, string[]> = {};
  for (const [partNumber, tags] of Object.entries(all)) {
    if (wanted.has(partNumber) && tags.length > 0) {
      result[partNumber] = tags;
    }
  }
  return result;
}

/** Replaces a part's tag list outright (empty array clears it) — the
 *  part detail page's edit form submits the full set it wants, not a
 *  delta, so there's no separate add/remove entry point. */
export async function setTagsForPart(partNumber: string, tags: string[]): Promise<void> {
  const redis = requireRedis();
  const all = await getAllPartTags();
  const deduped = Array.from(new Set(tags.map((tag) => tag.trim()).filter(Boolean)));
  if (deduped.length === 0) {
    delete all[partNumber];
  } else {
    all[partNumber] = deduped;
  }
  await redis.set(PART_TAGS_KEY, all);
}
