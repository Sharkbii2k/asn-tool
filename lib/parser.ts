"use client";

import * as pdfjsLib from "pdfjs-dist";
import Tesseract from "tesseract.js";
import { ParsedDoc, ParsedItem, HeaderRow } from "./types";

pdfjsLib.GlobalWorkerOptions.workerSrc = `https://cdnjs.cloudflare.com/ajax/libs/pdf.js/${pdfjsLib.version}/pdf.worker.min.mjs`;

function normalizeRev(v: string) {
  const txt = String(v || "").trim().replace(/\.0$/, "");
  return /^\d+$/.test(txt) ? txt.padStart(2, "0") : txt;
}
function find(pattern: RegExp, text: string): string {
  return text.match(pattern)?.[1]?.trim() || "";
}
function extractBlock(label: string, text: string, nextLabels: string[]) {
  const start = text.search(new RegExp(`${label}\s*:`, "i"));
  if (start < 0) return "";
  const slice = text.slice(start);
  let end = slice.length;
  for (const n of nextLabels) {
    const idx = slice.search(new RegExp(`\n\s*${n}\s*:`, "i"));
    if (idx > 0) end = Math.min(end, idx);
  }
  return slice.replace(new RegExp(`^.*?${label}\s*:\s*`, "i"), "").slice(0, end).replace(/\s+/g, " ").trim();
}

export async function fileToText(file: File): Promise<string> {
  const lower = file.name.toLowerCase();
  if (lower.endsWith(".pdf")) {
    const bytes = await file.arrayBuffer();
    const pdf = await pdfjsLib.getDocument({ data: bytes }).promise;
    let text = "";
    for (let i = 1; i <= pdf.numPages; i++) {
      const page = await pdf.getPage(i);
      const content = await page.getTextContent();
      text += content.items.map((it: any) => ("str" in it ? it.str : "")).join(" ") + "\n";
    }
    if (text.trim()) return text;
  }
  const result = await Tesseract.recognize(file, "eng");
  return result.data.text || "";
}

function parseItems(text: string, defaultLine: string): ParsedItem[] {
  const lines = text.split(/\r?\n/).map((l) => l.replace(/\s+/g, " ").trim()).filter(Boolean);
  const items: ParsedItem[] = [];
  for (const line of lines) {
    const m = line.match(/^(\d+)\s+([0-9A-Z\-]+)\s+([0-9]{6,})\s+(\d{1,2})\s+([0-9]+(?:\.[0-9]+)?)\s+([A-Z]{1,3})\s+([0-9]+(?:\.[0-9]+)?)/i);
    if (!m) continue;
    const [, seq, poNo, itemNo, rev, qty, uom, netWeight] = m;
    items.push({
      seq: Number(seq),
      poNo,
      itemNo: itemNo.trim(),
      rev: normalizeRev(rev),
      quantity: Number(qty),
      uom,
      netWeight: Number(netWeight),
      grossWeight: "",
      packingSpec: "",
      lotRef: "",
      lineNo: defaultLine
    });
  }
  return items;
}

export function parseTextToDoc(text: string, sourceFile: string): ParsedDoc {
  const asnNo = find(/ASN\s*No\.?\s*:?\s*([A-Z]{2}\d{6,})/i, text);
  const eta = find(/ETA\s*:?\s*([0-9\-\/]{8,10}\s+[0-9:]{4,5})/i, text);
  const etd = find(/ETD\s*:?\s*([0-9\-\/]{8,10}\s+[0-9:]{4,5})/i, text);
  const routeCode = find(/\b(XC\d+-TC\d+)\b/i, text);
  const soldTo = extractBlock("Sold To", text, ["Bill To", "Ship To", "Location", "Seq"]);
  const billTo = extractBlock("Bill To", text, ["Ship To", "Location", "Seq"]);
  const shipTo = extractBlock("Ship To", text, ["Location", "Seq"]);
  const location = extractBlock("Location", text, ["Seq", "PO No", "Item No"]);
  const lineCandidates = Array.from(text.matchAll(/\b(?:C\d|GP)[-A-Z0-9]+\b/g)).map((m) => m[0]);
  const lineNo = lineCandidates[lineCandidates.length - 1] || "";
  const [date = "", time = ""] = eta.split(" ");
  const items = parseItems(text, lineNo);

  return {
    sourceFile,
    asnNo,
    eta,
    etd,
    date,
    time,
    soldTo,
    billTo,
    shipTo,
    location,
    routeCode,
    lineNo,
    totalQuantity: items.reduce((sum, item) => sum + Number(item.quantity || 0), 0),
    items,
    rawText: text
  };
}

export function docsToHeaderRows(docs: ParsedDoc[]): HeaderRow[] {
  return docs.map((doc) => ({
    "ASN No": doc.asnNo,
    "ETA": doc.eta,
    "ETD": doc.etd,
    "Sold To": doc.soldTo,
    "Bill To": doc.billTo,
    "Ship To": doc.shipTo,
    "Location": doc.location,
    "Line No": doc.lineNo
  }));
}
