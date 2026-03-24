"use client";

import ExcelJS from "exceljs";
import { saveAs } from "file-saver";
import { HeaderRow, LinesRow, ParsedDoc } from "./types";

function style(cell: ExcelJS.Cell, opts: { fill?: string; bold?: boolean; center?: boolean; border?: boolean; size?: number } = {}) {
  if (opts.fill) cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: opts.fill } };
  cell.alignment = { vertical: "middle", horizontal: opts.center ? "center" : undefined, wrapText: true };
  if (opts.bold || opts.size) cell.font = { bold: !!opts.bold, size: opts.size || 11 };
  if (opts.border) {
    cell.border = {
      top: { style: "thin", color: { argb: "B7B7B7" } },
      left: { style: "thin", color: { argb: "B7B7B7" } },
      right: { style: "thin", color: { argb: "B7B7B7" } },
      bottom: { style: "thin", color: { argb: "B7B7B7" } }
    };
  }
}

export async function exportExcel(docs: ParsedDoc[], headers: HeaderRow[], lines: LinesRow[], summary: any[]) {
  const wb = new ExcelJS.Workbook();
  const title = "FFF2CC";
  const head = "D9EAD3";
  const sub = "EDEDED";
  const calc = "EAF2F8";
  const ok = "E2F0D9";
  const warn = "FCE4D6";
  const green = "C6E0B4";
  const blue = "BDD7EE";
  const orange = "FCE4D6";
  const red = "F4CCCC";
  const alt1 = "F7FBFF";
  const alt2 = "FFF9F0";

  // ASN
  const ws1 = wb.addWorksheet("ASN");
  ws1.columns = [
    { width: 8 }, { width: 16 }, { width: 15 }, { width: 8 }, { width: 12 }, { width: 8 },
    { width: 16 }, { width: 18 }, { width: 16 }, { width: 28 }, { width: 12 }, { width: 4 }, { width: 14 }, { width: 10 }
  ];
  const asnHeaders = ["Seq","PO No.","Item No.","Rev.","Quantity","Uom","Net Weight (KG)","Gross Weight (KG)","Packing Spec.","Lot/MI No./SO No./Invoice No","Line No."];
  let r = 1;
  docs.forEach((doc, idx) => {
    const blockFill = idx % 2 === 0 ? alt1 : alt2;
    ws1.mergeCells(r, 1, r, 11);
    for (let c = 1; c <= 11; c++) style(ws1.getCell(r, c), { fill: title, border: true });
    ws1.getCell(r, 1).value = doc.asnNo;
    style(ws1.getCell(r, 1), { fill: title, bold: true, center: true, border: true, size: 15 });
    ws1.getCell(r + 1, 1).value = "Date:";
    ws1.getCell(r + 1, 2).value = doc.date;
    ws1.getCell(r + 1, 4).value = "Time:";
    ws1.getCell(r + 1, 5).value = doc.time;
    ws1.getCell(r + 2, 1).value = "Route:";
    ws1.getCell(r + 2, 2).value = doc.routeCode;
    ws1.getCell(r + 2, 4).value = "Line No:";
    ws1.getCell(r + 2, 5).value = doc.lineNo;

    const hr = r + 4;
    asnHeaders.forEach((label, i) => {
      ws1.getCell(hr, i + 1).value = label;
      style(ws1.getCell(hr, i + 1), { fill: head, bold: true, center: true, border: true });
    });

    doc.items.forEach((item, i) => {
      const rr = hr + 1 + i;
      const vals = [item.seq, item.poNo, item.itemNo, item.rev, item.quantity, item.uom, item.netWeight, item.grossWeight, item.packingSpec, item.lotRef, item.lineNo];
      vals.forEach((v, c) => {
        ws1.getCell(rr, c + 1).value = v as any;
        style(ws1.getCell(rr, c + 1), { border: true, center: true });
        ws1.getCell(rr, c + 1).fill = { type: "pattern", pattern: "solid", fgColor: { argb: blockFill } };
      });
    });

    const tr = hr + 1 + doc.items.length;
    ws1.mergeCells(tr, 1, tr, 4);
    ws1.getCell(tr, 1).value = "Total Quantity";
    style(ws1.getCell(tr, 1), { fill: sub, bold: true, border: true });
    ws1.getCell(tr, 5).value = doc.totalQuantity;
    style(ws1.getCell(tr, 5), { fill: sub, bold: true, center: true, border: true });
    for (let c = 1; c <= 11; c++) style(ws1.getCell(tr, c), { border: true });
    r = tr + 3;
  });

  [["Tổng ASN", docs.length, green], ["CPT", docs.filter((d) => String(d.lineNo).startsWith("C2")).length, blue], ["OP", docs.filter((d) => String(d.lineNo).startsWith("C1")).length, orange], ["GP", docs.filter((d) => String(d.lineNo).startsWith("GP")).length, red]].forEach((row, idx) => {
    ws1.getCell(idx + 1, 13).value = row[0] as any;
    ws1.getCell(idx + 1, 14).value = row[1] as any;
    style(ws1.getCell(idx + 1, 13), { fill: row[2] as string, bold: true, center: true, border: true });
    style(ws1.getCell(idx + 1, 14), { fill: row[2] as string, bold: true, center: true, border: true });
  });

  // Header
  const ws2 = wb.addWorksheet("Header");
  ws2.columns = [{ width: 16 }, { width: 18 }, { width: 12 }, { width: 38 }, { width: 52 }, { width: 38 }, { width: 62 }, { width: 12 }];
  const headerCols = ["ASN No","ETA","ETD","Sold To","Bill To","Ship To","Location","Line No"];
  headerCols.forEach((label, idx) => {
    ws2.getCell(1, idx + 1).value = label;
    style(ws2.getCell(1, idx + 1), { fill: head, bold: true, center: true, border: true });
  });
  headers.forEach((row, idx) => {
    const fill = idx % 2 === 0 ? alt1 : alt2;
    headerCols.forEach((key, cidx) => {
      ws2.getCell(idx + 2, cidx + 1).value = (row as any)[key];
      style(ws2.getCell(idx + 2, cidx + 1), { border: true });
      ws2.getCell(idx + 2, cidx + 1).fill = { type: "pattern", pattern: "solid", fgColor: { argb: fill } };
    });
  });

  // Lines
  const ws3 = wb.addWorksheet("Lines");
  ws3.columns = [
    { width: 16 }, { width: 14 }, { width: 8 }, { width: 12 }, { width: 10 }, { width: 12 },
    { width: 10 }, { width: 12 }, { width: 14 }, { width: 10 }, { width: 12 }, { width: 12 },
    { width: 12 }, { width: 12 }, { width: 12 }, { width: 10 }
  ];
  const lineCols = ["ASN","Item","Rev","Quantity","Packing","Thùng chẵn","SL lẻ PCS","Tổng Cartons","Line No","Location","Packing Found","Calc Status"];
  lineCols.forEach((label, idx) => {
    ws3.getCell(1, idx + 1).value = label;
    style(ws3.getCell(1, idx + 1), { fill: head, bold: true, center: true, border: true });
  });

  let currentAsn = "";
  let block = -1;
  lines.forEach((row, idx) => {
    if (row["ASN"] !== currentAsn) { currentAsn = row["ASN"]; block += 1; }
    const fill = block % 2 === 0 ? alt1 : alt2;
    lineCols.forEach((key, cidx) => {
      ws3.getCell(idx + 2, cidx + 1).value = (row as any)[key];
      style(ws3.getCell(idx + 2, cidx + 1), { border: true, center: true });
      ws3.getCell(idx + 2, cidx + 1).fill = { type: "pattern", pattern: "solid", fgColor: { argb: fill } };
    });
    ws3.getCell(idx + 2, 5).fill = { type: "pattern", pattern: "solid", fgColor: { argb: row["Packing Found"] === "YES" ? ok : warn } };
    ws3.getCell(idx + 2, 6).fill = { type: "pattern", pattern: "solid", fgColor: { argb: calc } };
    ws3.getCell(idx + 2, 7).fill = { type: "pattern", pattern: "solid", fgColor: { argb: calc } };
    ws3.getCell(idx + 2, 8).fill = { type: "pattern", pattern: "solid", fgColor: { argb: calc } };
    ws3.getCell(idx + 2, 11).fill = { type: "pattern", pattern: "solid", fgColor: { argb: row["Packing Found"] === "YES" ? ok : warn } };
    ws3.getCell(idx + 2, 12).fill = { type: "pattern", pattern: "solid", fgColor: { argb: row["Calc Status"] === "OK" ? ok : warn } };
  });

  ws3.mergeCells(1, 13, 1, 16);
  ws3.getCell(1, 13).value = "Tổng Cartons";
  style(ws3.getCell(1, 13), { fill: title, bold: true, center: true, border: true });
  ["Location", "Thùng chẵn", "Tổng số thùng lẻ", "Tổng"].forEach((label, idx) => {
    ws3.getCell(2, 13 + idx).value = label;
    style(ws3.getCell(2, 13 + idx), { fill: head, bold: true, center: true, border: true });
  });
  summary.forEach((row, idx) => {
    ws3.getCell(idx + 3, 13).value = row.Location;
    ws3.getCell(idx + 3, 14).value = row["Thùng chẵn"];
    ws3.getCell(idx + 3, 15).value = row["Tổng số thùng lẻ"];
    ws3.getCell(idx + 3, 16).value = row["Tổng"];
    for (let c = 13; c <= 16; c++) style(ws3.getCell(idx + 3, c), { border: true, center: true });
  });

  const buffer = await wb.xlsx.writeBuffer();
  const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  saveAs(blob, "ASN_TOOL_GM_final.xlsx");
}
