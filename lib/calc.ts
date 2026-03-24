import { LinesRow, PackingRow, ParsedDoc } from "./types";

function normalizeRev(v: string) {
  const txt = String(v || "").trim().replace(/\.0$/, "");
  return /^\d+$/.test(txt) ? txt.padStart(2, "0") : txt;
}
function locFromLine(line: string) {
  if (String(line).includes("C2")) return "CPT";
  if (String(line).includes("C1")) return "OP";
  if (String(line).includes("GP")) return "GP";
  return "OTHER";
}

export function buildLinesByAsn(docs: ParsedDoc[], packing: PackingRow[]): LinesRow[] {
  const packMap = new Map<string, number>();
  packing.forEach((r) => {
    const pack = Number(r.pack);
    if (r.item && r.rev && pack > 0) packMap.set(`${String(r.item).trim()}__${normalizeRev(String(r.rev))}`, pack);
  });

  const out: LinesRow[] = [];
  for (const doc of docs) {
    for (const item of doc.items) {
      const key = `${String(item.itemNo).trim()}__${normalizeRev(item.rev)}`;
      const pack = packMap.get(key) || 0;
      if (pack > 0) {
        const full = Math.floor(Number(item.quantity) / pack);
        const loosePcs = Number(item.quantity) % pack;
        const looseCarton = loosePcs > 0 ? 1 : 0;
        out.push({
          "ASN": doc.asnNo,
          "Item": item.itemNo,
          "Rev": normalizeRev(item.rev),
          "Quantity": Number(item.quantity),
          "Packing": pack,
          "Thùng chẵn": full,
          "SL lẻ PCS": loosePcs,
          "Tổng Cartons": full + looseCarton,
          "Line No": doc.lineNo,
          "Location": locFromLine(doc.lineNo),
          "Packing Found": "YES",
          "Calc Status": "OK",
          "__loose_carton": looseCarton
        });
      } else {
        out.push({
          "ASN": doc.asnNo,
          "Item": item.itemNo,
          "Rev": normalizeRev(item.rev),
          "Quantity": Number(item.quantity),
          "Packing": "",
          "Thùng chẵn": "",
          "SL lẻ PCS": "",
          "Tổng Cartons": "",
          "Line No": doc.lineNo,
          "Location": locFromLine(doc.lineNo),
          "Packing Found": "NO",
          "Calc Status": "CHECK",
          "__loose_carton": 0
        });
      }
    }
  }
  return out;
}

export function buildSummary(lines: LinesRow[]) {
  const base = ["CPT", "OP", "GP"].map((loc) => ({ Location: loc, "Thùng chẵn": 0, "Tổng số thùng lẻ": 0, "Tổng": 0 }));
  lines.forEach((r) => {
    const row = base.find((x) => x.Location === r["Location"]);
    if (!row || r["Calc Status"] !== "OK") return;
    row["Thùng chẵn"] += Number(r["Thùng chẵn"] || 0);
    row["Tổng số thùng lẻ"] += Number(r["__loose_carton"] || 0);
    row["Tổng"] += Number(r["Tổng Cartons"] || 0);
  });
  return base;
}
