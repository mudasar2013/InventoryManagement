import type { Job, JobPart, Part } from "./types";

export const parts: Part[] = [
  {
    id: "prt-001",
    part_number: "WR17X11705",
    description: "Refrigerator water filter cartridge",
    bin_location: "A-12-04",
    quantity_on_hand: 14,
    status: "In Stock",
  },
  {
    id: "prt-002",
    part_number: "WPW10311524",
    description: "Washer drain pump assembly",
    bin_location: "B-03-11",
    quantity_on_hand: 3,
    status: "Low Stock",
  },
  {
    id: "prt-003",
    part_number: "5303918277",
    description: "Electric dryer heating element",
    bin_location: "C-07-02",
    quantity_on_hand: 0,
    status: "Out of Stock",
  },
  {
    id: "prt-004",
    part_number: "WB27X10934",
    description: "Range infinite burner switch",
    bin_location: "A-08-15",
    quantity_on_hand: 8,
    status: "In Stock",
  },
  {
    id: "prt-005",
    part_number: "WP3392519",
    description: "Dryer thermal fuse 3392519",
    bin_location: "D-01-06",
    quantity_on_hand: 22,
    status: "In Stock",
  },
];

export const jobs: Job[] = [
  {
    id: "job-1042",
    job_number: "1042",
    customer_name: "Maria Lopez",
    appliance: "GE French-door refrigerator",
    issue: "Ice maker running, water tastes off",
    status: "Scheduled",
    scheduled_for: "2026-09-03T09:30:00",
  },
  {
    id: "job-1043",
    job_number: "1043",
    customer_name: "David Chen",
    appliance: "Whirlpool top-load washer",
    issue: "Cycle ends with standing water in the tub",
    status: "In Progress",
    scheduled_for: "2026-09-02T13:00:00",
  },
  {
    id: "job-1044",
    job_number: "1044",
    customer_name: "Priya Patel",
    appliance: "Frigidaire electric dryer",
    issue: "Drum turns but no heat",
    status: "Scheduled",
    scheduled_for: "2026-09-04T11:00:00",
  },
];

export const jobParts: JobPart[] = [
  {
    id: "jp-1",
    job_id: "job-1042",
    part_id: "prt-001",
    quantity_needed: 1,
  },
  {
    id: "jp-2",
    job_id: "job-1043",
    part_id: "prt-002",
    quantity_needed: 1,
  },
  {
    id: "jp-3",
    job_id: "job-1044",
    part_id: "prt-003",
    quantity_needed: 1,
  },
  {
    id: "jp-4",
    job_id: "job-1044",
    part_id: "prt-005",
    quantity_needed: 1,
  },
];
