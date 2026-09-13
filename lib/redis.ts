import { Redis } from "@upstash/redis";

/**
 * Shared Redis access for every app-managed store (SharePoint source
 * configs, tags, tag-to-part assignments — see
 * lib/sources/sharepoint-source-store.ts and lib/tags-store.ts). Pulled
 * out of sharepoint-source-store.ts (the first store that needed it) so
 * a second, unrelated store doesn't re-implement the same credential
 * lookup and risk drifting from it.
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
export function getRedisCredentials(): { url: string; token: string } | null {
  const url = process.env.UPSTASH_REDIS_REST_URL ?? process.env.KV_REST_API_URL;
  const token = process.env.UPSTASH_REDIS_REST_TOKEN ?? process.env.KV_REST_API_TOKEN;
  if (!url || !token) {
    return null;
  }
  return { url, token };
}

export function getRedis(): Redis | null {
  const credentials = getRedisCredentials();
  return credentials ? new Redis(credentials) : null;
}

/** Whether a Redis store is reachable at all — pages that manage
 *  app-native data (Data sources, Settings) use this to decide whether
 *  to show their add/edit form or a note pointing at Vercel's storage
 *  marketplace instead. */
export function isRedisConfigured(): boolean {
  return getRedisCredentials() !== null;
}
