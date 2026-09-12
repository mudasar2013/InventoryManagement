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
 * Uses whatever Redis a Vercel Marketplace storage integration injects
 * (Upstash's UPSTASH_REDIS_REST_URL / UPSTASH_REDIS_REST_TOKEN — Vercel
 * KV itself was retired and existing stores auto-migrated to Upstash in
 * December 2024, see https://vercel.com/docs/redis). Nothing here is
 * Upstash-specific beyond the client and env var names — swapping to a
 * different REST-compatible Redis provider only touches getRedis().
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

function getRedis(): Redis | null {
  const url = process.env.UPSTASH_REDIS_REST_URL;
  const token = process.env.UPSTASH_REDIS_REST_TOKEN;
  if (!url || !token) {
    return null;
  }
  return new Redis({ url, token });
}

/**
 * Whether a Redis store is reachable at all — the "Data sources" page
 * uses this to decide whether to show the add-source form or a note
 * pointing at Vercel's storage marketplace instead.
 */
export function isSourceStoreConfigured(): boolean {
  return Boolean(
    process.env.UPSTASH_REDIS_REST_URL && process.env.UPSTASH_REDIS_REST_TOKEN,
  );
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
