import assert from "node:assert/strict";
import { test } from "node:test";
import { isTechScoreJobsConfigured, mapTechScoreJob, mapWorkStatus } from "./tech-score-jobs";
import type { TechScoreJobRow } from "./tech-score-jobs";

function row(overrides: Partial<TechScoreJobRow> = {}): TechScoreJobRow {
  return {
    id: "412",
    hcpJobId: "hcp-abc123",
    jobNumber: "6083",
    customerName: "Jane Doe",
    description: "Refrigerator not cooling",
    workStatus: "scheduled",
    completedAt: null,
    updatedAt: "2026-09-01T12:00:00.000Z",
    ...overrides,
  };
}

test("mapWorkStatus: a completedAt timestamp always means Completed, regardless of workStatus text", () => {
  assert.equal(mapWorkStatus("scheduled", "2026-08-01T00:00:00.000Z"), "Completed");
  assert.equal(mapWorkStatus(null, "2026-08-01T00:00:00.000Z"), "Completed");
});

test("mapWorkStatus: matches 'complet' and 'progress' as substrings, case-insensitively", () => {
  assert.equal(mapWorkStatus("Completed", null), "Completed");
  assert.equal(mapWorkStatus("job_completed", null), "Completed");
  assert.equal(mapWorkStatus("In Progress", null), "In Progress");
  assert.equal(mapWorkStatus("in_progress", null), "In Progress");
});

test("mapWorkStatus: anything unrecognized (or null) defaults to Scheduled", () => {
  assert.equal(mapWorkStatus("needs_scheduling", null), "Scheduled");
  assert.equal(mapWorkStatus(null, null), "Scheduled");
  assert.equal(mapWorkStatus("canceled", null), "Scheduled");
});

test("mapTechScoreJob: maps fields, prefixes the id, and leaves appliance blank", () => {
  const job = mapTechScoreJob(row());

  assert.equal(job.id, "tech-score-412");
  assert.equal(job.job_number, "6083");
  assert.equal(job.customer_name, "Jane Doe");
  assert.equal(job.appliance, "");
  assert.equal(job.issue, "Refrigerator not cooling");
  assert.equal(job.status, "Scheduled");
  assert.equal(job.scheduled_for, "2026-09-01T12:00:00.000Z");
});

test("mapTechScoreJob: a completed job maps to status Completed", () => {
  const job = mapTechScoreJob(row({ completedAt: "2026-08-15T09:00:00.000Z", workStatus: "completed" }));
  assert.equal(job.status, "Completed");
});

test("isTechScoreJobsConfigured: true only when both env vars are set", () => {
  const original = {
    url: process.env.TECH_SCORE_JOBS_API_URL,
    key: process.env.TECH_SCORE_JOBS_API_KEY,
  };
  try {
    delete process.env.TECH_SCORE_JOBS_API_URL;
    delete process.env.TECH_SCORE_JOBS_API_KEY;
    assert.equal(isTechScoreJobsConfigured(), false);

    process.env.TECH_SCORE_JOBS_API_URL = "https://tech-score.example.com";
    assert.equal(isTechScoreJobsConfigured(), false);

    process.env.TECH_SCORE_JOBS_API_KEY = "secret";
    assert.equal(isTechScoreJobsConfigured(), true);
  } finally {
    if (original.url === undefined) delete process.env.TECH_SCORE_JOBS_API_URL;
    else process.env.TECH_SCORE_JOBS_API_URL = original.url;
    if (original.key === undefined) delete process.env.TECH_SCORE_JOBS_API_KEY;
    else process.env.TECH_SCORE_JOBS_API_KEY = original.key;
  }
});
