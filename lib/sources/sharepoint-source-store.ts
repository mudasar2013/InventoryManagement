import { randomUUID } from "node:crypto";
import { Redis } from "@upstash/redis";

/**
 * A SharePoint workbook a technician added from the "Data sources" page
 * (see app/sources/page.tsx), stored in Redis so it survives without a
 * redeploy — unlike the single legacy source configured via
 * SHAREPOINT_SITE_HOSTNAME etc. (see readSharePointExcelConfig), which
 * still works and is kept for backward compatibility, but can't be added
 * to or removed without editing environment variables.
 *
 * Uses whatever Redis a Vercel Marketplace storage integration injects.
 * In practice the Vercel Marketplace "Upstash for Redis" integration sets
 * KV_REST_API_URL / KV_REST_API_TOKEN (it keeps the legacy Vercel KV
 * variable names for compatibility with existing KV code — Vercel KV
 * itself was retired and existing stores auto-migrated to Upstash in
 * December 2024, see https://vercel.com/docs/redis). Some setups instead
 * expose Upstash's own UPSTASH_REDIS_REST_URL / UPSTASH_REDIS_REST_TOKEN
 * names (e.g. Redis.fromEnv() conventions, or a manually-created Upstash
 * database), so both pairs are checked here — UPSTASH_* first, then
 * falling back to the KV_* names actually present on this project.
 */
export interface StoredSharePointSource {
  id: string;
  label: string;
  siteHostname: string;
  sitePath: string;
  filePath: string;
  tableName: string;
  createdAt: string;
}

export type NewSharePointSource = Omit<StoredSharePointSource, "id" | "createdAt">;

const STORE_KEY = "inventory:sharepoint-sources";

/**
 * Reads whichever REST Redis credentials the current Vercel project
 * actually has, trying Upstash's own env var names first and falling
 * back to the KV_REST_API_* names the Vercel Marketplace "Upstash for
 * Redis" integration sets by default.
 */
function getRedisCredentials(): { url: string; token: string } | null {
  const url =
    process.env.UPSTASH_REDIS_REST_URL ?? process.env.KV_REST_API_URL;
  const token =
    process.env.UPSTASH_REDIS_REST_TOKEN ?? process.env.KV_REST_API_TOKEN;
  if (!url || !token) {
    return null;
  }
  return { url, token };
}

function getRedis(): Redis | null {
  const credentials = getRedisCredentials();
  return credentials ? new Redis(credentials) : null;
}

/**
 * Whether a Redis store is reachable at all — the "Data sources" page
 * uses this to decide whether to show the add-source form or a note
 * pointing at Vercel's storage marketplace instead.
 */
export function isSourceStoreConfigured(): boolean {
  return getRedisCredentials() !== null;
}

/** Returns [] (never throws) when Redis isn't configured — same
 *  fail-open philosophy as readSharePointExcelConfig(): a missing store
 *  just means zero user-added sources, not a broken page. */
export async function listStoredSharePointSources(): Promise<
  StoredSharePointSource[]
> {
  const redis = getRedis();
  if (!redis) {
    return [];
  }
  const stored = await redis.get<StoredSharePointSource[]>(STORE_KEY);
  return stored ?? [];
}

export async function addStoredSharePointSource(
  input: NewSharePointSource,
): Promise<StoredSharePointSource> {
  const redis = getRedis();
  if (!redis) {
    throw new Error(
      "No Redis store configured. Add a Redis integration to this Vercel project " +
        "(Storage tab → Marketplace Database Providers) to enable adding sources here.",
    );
  }

  const current = await listStoredSharePointSources();
  const entry: StoredSharePointSource = {
    ...input,
    id: `sharepoint-${randomUUID()}`,
    createdAt: new Date().toISOString(),
  };
  await redis.set(STORE_KEY, [...current, entry]);
  return entry;
}

/**
 * Overwrites an existing entry's fields in place (id and createdAt are
 * kept) — used by the "Data sources" page to fix a typo'd hostname/path
 * without deleting and re-adding the source (which would also lose its
 * position in the merge-priority list). Throws if the store isn't
 * configured, or if no entry with that id exists.
 */
export async function updateStoredSharePointSource(
  id: string,
  updates: NewSharePointSource,
): Promise<StoredSharePointSource> {
  const redis = getRedis();
  if (!redis) {
    throw new Error(
      "No Redis store configured. Add a Redis integration to this Vercel project " +
        "(Storage tab → Marketplace Database Providers) to enable editing sources here.",
    );
  }

  const current = await listStoredSharePointSources();
  const index = current.findIndex((source) => source.id === id);
  if (index === -1) {
    throw new Error("That source no longer exists — it may have been removed already.");
  }

  const updated: StoredSharePointSource = {
    ...current[index],
    ...updates,
  };
  const next = [...current];
  next[index] = updated;
  await redis.set(STORE_KEY, next);
  return updated;
}

export async function removeStoredSharePointSource(id: string): Promise<void> {
  const redis = getRedis();
  if (!redis) {
    return;
  }
  const current = await listStoredSharePointSources();
  await redis.set(
    STORE_KEY,
    current.filter((source) => source.id !== id),
  );
}
