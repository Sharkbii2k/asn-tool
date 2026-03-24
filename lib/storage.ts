import { PackingRow, PairRow } from "./types";

const PACKING_KEY = "asn-tool-packing-master";
const PAIR_KEY = "asn-tool-pair-rules";

export function loadPacking(): PackingRow[] {
  if (typeof window === "undefined") return [];
  try { return JSON.parse(localStorage.getItem(PACKING_KEY) || "[]"); } catch { return []; }
}
export function savePacking(rows: PackingRow[]) { localStorage.setItem(PACKING_KEY, JSON.stringify(rows)); }

export function loadPairs(): PairRow[] {
  if (typeof window === "undefined") return [];
  try { return JSON.parse(localStorage.getItem(PAIR_KEY) || "[]"); } catch { return []; }
}
export function savePairs(rows: PairRow[]) { localStorage.setItem(PAIR_KEY, JSON.stringify(rows)); }
