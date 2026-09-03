import type { Part, PartStatus, RawPart } from "./types";

/**
 * A part at or below this quantity (but above zero) is "Low Stock".
 * Kept as one shared constant so every source normalizes the same way —
 * previously each mock record hand-picked its own status, which let
 * quantity and status drift out of sync.
 */
export const DEFAULT_LOW_STOCK_THRESHOLD = 5;

export function deriveStatus(
  quantityOnHand: number,
  lowStockThreshold: number = DEFAULT_LOW_STOCK_THRESHOLD,
): PartStatus {
  if (quantityOnHand <= 0) {
    return "Out of Stock";
  }
  if (quantityOnHand <= lowStockThreshold) {
    return "Low Stock";
  }
  return "In Stock";
}

/**
 * Normalize a source's raw record into the shape the rest of the app
 * renders, computing status instead of trusting a stored field.
 */
export function toPart(
  raw: RawPart,
  lowStockThreshold: number = DEFAULT_LOW_STOCK_THRESHOLD,
): Part {
  return {
    ...raw,
    status: deriveStatus(raw.quantity_on_hand, lowStockThreshold),
  };
}
