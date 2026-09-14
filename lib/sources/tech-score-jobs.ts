import type { Job, JobStatus } from "../types";

/**
 * Fetches real Housecall Pro jobs from the tech-score app's read-only
 * job feed (see tech-score's app/api/public/inventory-jobs/route.ts),
 * so "Attach to a job" offers actual current work instead of the fixed
 * local demo list in lib/sources/local-source.ts.
 *
 * Configured via two env vars:
 * - TECH_SCORE_JOBS_API_URL — tech-score's own deployed base URL (e.g.
 *   "https://tech-score.example.com"), no trailing path.
 * - TECH_SCORE_JOBS_API_KEY — must match that deployment's own
 *   INVENTORY_API_KEY env var, sent as an `x-api-key` header.
 *
 * Both must be set for this to run at all. When either is missing,
 * loadInventory (see lib/getInventory.ts) skips this entirely and keeps
 * the bundled local demo jobs — this module never partially runs.
 * tech-score's own database is never reached directly from here: this
 * app only ever talks to that one narrow, authenticated HTTP endpoint,
 * so a schema change or a payroll/warranty column living right next to
 * `jobs` in that database is never something this app could see even
 * by accident.
 */

/** The shape tech-score's /api/public/inventory-jobs endpoint returns
 *  per job — see that route's own doc comment for what's deliberately
 *  left out (payroll/warranty fields, the raw HCP payload, etc.). */
export interface TechScoreJobRow {
  id: string;
  hcpJobId: string;
  jobNumber: string;
  customerName: string;
  description: string;
  workStatus: string | null;
  completedAt: string | null;
  updatedAt: string;
}

export function isTechScoreJobsConfigured(): boolean {
  return Boolean(process.env.TECH_SCORE_JOBS_API_URL && process.env.TECH_SCORE_JOBS_API_KEY);
}

/**
 * Collapses Housecall Pro's own free-form work_status text down to this
 * app's 3-value JobStatus. HCP's exact wording varies by job type and
 * account ("needs scheduling", "in_progress", "in progress", ...), so
 * this matches on substrings rather than an exact enum. A completed
 * job is judged by `completedAt` being set (the column tech-score's own
 * ingest treats as authoritative), not by the status text, since a
 * workStatus string alone isn't reliably "completed" across every job
 * type tech-score has ever synced. Anything not recognizably completed
 * or in progress defaults to "Scheduled" (open/upcoming work) — the
 * safer default for a list a technician is choosing a job to attach a
 * part to from.
 */
export function mapWorkStatus(workStatus: string | null, completedAt: string | null): JobStatus {
  if (completedAt) return "Completed";
  const normalized = (workStatus ?? "").toLowerCase();
  if (normalized.includes("complet")) return "Completed";
  if (normalized.includes("progress")) return "In Progress";
  return "Scheduled";
}

/**
 * Maps one row from tech-score's job feed to this app's Job shape.
 * Pulled out as a pure function (see mapTableRowsToRawParts in
 * sharepoint-excel-source.ts for the same pattern) so the shaping logic
 * is unit-testable without a network call.
 */
export function mapTechScoreJob(row: TechScoreJobRow): Job {
  return {
    id: `tech-score-${row.id}`,
    job_number: row.jobNumber,
    customer_name: row.customerName,
    // tech-score's `jobs` table doesn't separate "appliance" from the
    // rest of the job description the way this app's own Job shape
    // does — the full description goes in `issue`, and `appliance` is
    // left blank rather than guessed at by splitting free text.
    appliance: "",
    issue: row.description,
    status: mapWorkStatus(row.workStatus, row.completedAt),
    // tech-score's `public.jobs` table has no appointment time of its
    // own — that lives in a separate `dispatch.jobs` table this feed
    // doesn't join against. `updatedAt` is the closest available
    // timestamp: a best-effort "last activity" time, not a real
    // scheduled appointment.
    scheduled_for: row.updatedAt,
  };
}

export async function fetchTechScoreJobs(): Promise<Job[]> {
  const baseUrl = process.env.TECH_SCORE_JOBS_API_URL;
  const apiKey = process.env.TECH_SCORE_JOBS_API_KEY;
  if (!baseUrl || !apiKey) {
    throw new Error("TECH_SCORE_JOBS_API_URL/TECH_SCORE_JOBS_API_KEY are not configured");
  }

  const url = new URL("/api/public/inventory-jobs", baseUrl).toString();
  const response = await fetch(url, {
    headers: { "x-api-key": apiKey },
    cache: "no-store",
  });
  if (!response.ok) {
    throw new Error(`tech-score jobs feed returned HTTP ${response.status}`);
  }

  const payload: { jobs: TechScoreJobRow[] } = await response.json();
  return payload.jobs.map(mapTechScoreJob);
}
