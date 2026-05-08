"use strict";
const ExcelJS = require("exceljs");
const { calcTotals, fmtINR, toWords } = require("./invoice-data");

const NAVY   = "FF1F3864";
const SKYBLUE= "FFDBEAFE";
const LGREY  = "FFF1F5F9";
const ORANGE = "FFC65911";
const WHITE  = "FFFFFFFF";

const thin   = { style: "thin", color: { argb: "FFB8C4D4" } };
const allThin = { top: thin, bottom: thin, left: thin, right: thin };

function cell(ws, ref, value, opts = {}) {
  const c = ws.getCell(ref);
  c.value = value;
  if (opts.bold !== undefined || opts.size || opts.color || opts.italic) {
    c.font = {
      name: "Arial", bold: !!opts.bold, italic: !!opts.italic,
      size: opts.size || 9,
      color: { argb: opts.color || "FF111111" },
    };
  }
  if (opts.bg) c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: opts.bg } };
  if (opts.align || opts.valign || opts.wrap) {
    c.alignment = {
      horizontal: opts.align || "left",
      vertical:   opts.valign || "middle",
      wrapText:   !!opts.wrap,
    };
  }
  c.border = allThin;
  if (opts.numFmt) c.numFmt = opts.numFmt;
  return c;
}

function mergeRow(ws, ref, value, opts = {}) {
  ws.mergeCells(ref);
  const [top] = ref.split(":");
  return cell(ws, top, value, opts);
}

async function genXLSX(data) {
  const wb = new ExcelJS.Workbook();
  wb.creator = data.supplierName;
  const ws = wb.addWorksheet(`Invoice ${data.invoiceNo}`, { pageSetup: { fitToPage: true, fitToWidth: 1, orientation: "portrait" } });

  // Column widths
  ws.columns = [
    { width: 6 },   // A - S.No.
    { width: 30 },  // B - Description
    { width: 14 },  // C - Rate
    { width: 14 },  // D - Persons
    { width: 8 },   // E - Days
    { width: 16 },  // F - Amount
  ];

  let R = 1;

  // ── Title ──────────────────────────────────────────────────────────────────
  ws.mergeCells(`A${R}:F${R}`);
  cell(ws, `A${R}`, "INVOICE", { bold:true, size:16, color:"FFFFFFFF", bg:NAVY, align:"center" });
  ws.getRow(R).height = 26; R++;

  // ── Original for recipient ─────────────────────────────────────────────────
  ws.mergeCells(`A${R}:F${R}`);
  cell(ws, `A${R}`, "ORIGINAL FOR RECIPIENT", { bold:true, size:8, color:ORANGE, align:"right", bg:WHITE });
  ws.getRow(R).height = 14; R++;

  // ── Party headers ──────────────────────────────────────────────────────────
  ws.mergeCells(`A${R}:C${R}`);
  ws.mergeCells(`D${R}:F${R}`);
  cell(ws, `A${R}`, "Supplier", { bold:true, size:9, color:"FFFFFFFF", bg:NAVY, align:"center" });
  cell(ws, `D${R}`, "Receiver", { bold:true, size:9, color:"FFFFFFFF", bg:NAVY, align:"center" });
  ws.getRow(R).height = 16; R++;

  // ── Supplier info ──────────────────────────────────────────────────────────
  function partyRow(r, labelA, valA, labelD, valD) {
    ws.mergeCells(`B${r}:C${r}`); ws.mergeCells(`E${r}:F${r}`);
    cell(ws, `A${r}`, labelA, { bold:true, size:9, color:NAVY, bg:WHITE });
    cell(ws, `B${r}`, valA,   { size:9, bg:WHITE, wrap:true });
    if (labelD) cell(ws, `D${r}`, labelD, { bold:true, size:9, color:NAVY, bg:WHITE });
    if (valD)   cell(ws, `E${r}`, valD,   { size:9, bg:WHITE });
  }

  partyRow(R,   "Name",    data.supplierName,    "Name",   data.receiverName);   R++;
  partyRow(R,   "GSTIN",   data.supplierGSTIN,   "GSTIN",  data.receiverGSTIN);  R++;
  partyRow(R,   "PAN",     data.supplierPAN,     "PERIOD", data.period);         R++;
  ws.mergeCells(`B${R}:C${R}`); ws.mergeCells(`D${R}:F${R}`);
  cell(ws, `A${R}`, "Address", { bold:true, size:9, color:NAVY, bg:WHITE });
  const addrCell = cell(ws, `B${R}`, data.supplierAddress, { size:8, bg:WHITE, wrap:true });
  ws.getRow(R).height = 28; R++;

  // ── Meta row ──────────────────────────────────────────────────────────────
  ws.mergeCells(`A${R}:B${R}`); ws.mergeCells(`C${R}:D${R}`); ws.mergeCells(`E${R}:F${R}`);
  cell(ws, `A${R}`, "Place of Supply", { bold:true, size:9, color:NAVY, bg:LGREY, align:"center" });
  cell(ws, `C${R}`, "Invoice Date",    { bold:true, size:9, color:NAVY, bg:LGREY, align:"center" });
  cell(ws, `E${R}`, "Invoice No.",     { bold:true, size:9, color:NAVY, bg:LGREY, align:"center" });
  ws.getRow(R).height = 14; R++;

  ws.mergeCells(`A${R}:B${R}`); ws.mergeCells(`C${R}:D${R}`); ws.mergeCells(`E${R}:F${R}`);
  cell(ws, `A${R}`, data.placeOfSupply, { bold:true, size:10, align:"center" });
  cell(ws, `C${R}`, data.invoiceDate,   { bold:true, size:10, align:"center" });
  cell(ws, `E${R}`, data.invoiceNo,     { bold:true, size:10, align:"center" });
  ws.getRow(R).height = 16; R++;

  // ── Table headers ──────────────────────────────────────────────────────────
  ["S.No.", "Description & SAC Code", "Rate / Head", "No. of Persons", "Days", "Total Amount"]
    .forEach((h, i) => {
      const col = String.fromCharCode(65 + i);
      cell(ws, `${col}${R}`, h, { bold:true, size:9, color:"FFFFFFFF", bg:NAVY, align:"center", wrap:true });
    });
  ws.getRow(R).height = 22; R++;

  // ── Line items ─────────────────────────────────────────────────────────────
  const t = calcTotals(data);
  const startRow = R;
  t.lineItems.forEach((item, i) => {
    const bg = i % 2 === 1 ? SKYBLUE : WHITE;
    const cols = ["A","B","C","D","E","F"];
    const vals = [i+1, item.desc||"", item.rate||"", item.persons||"", item.days&&+item.days?item.days:"", item._amount||""];
    vals.forEach((v, ci) => {
      const ref = `${cols[ci]}${R}`;
      cell(ws, ref, v, {
        bg, size:9,
        align: ci <= 1 ? "left" : "right",
        numFmt: ci >= 2 && ci !== 4 && v ? "#,##0.00" : undefined,
      });
    });
    cell(ws, `A${R}`, i+1, { bg, size:9, align:"center" });
    ws.getRow(R).height = 15; R++;
  });

  // ── Totals ─────────────────────────────────────────────────────────────────
  const sp = data.svcPct??5, cp = data.cgstPct??9, gp = data.sgstPct??9;

  function totRow(label, val, hilight=false, grand=false) {
    ws.mergeCells(`A${R}:E${R}`);
    const bg = grand ? NAVY : hilight ? SKYBLUE : WHITE;
    const clr = grand ? "FFFFFFFF" : "FF1F3864";
    cell(ws, `A${R}`, label, { bold:true, size:9, color:clr, bg, align:"left" });
    cell(ws, `F${R}`, val,   { bold:true, size:9, color:clr, bg, align:"right", numFmt:"#,##0.00" });
    ws.getRow(R).height = 15; R++;
  }

  totRow("Total Taxable Amount",        t.taxable, true);
  totRow(`Service Charges @${sp}%`,     t.svc);
  totRow("Total",                        t.sub, true);
  totRow(`CGST ${cp}%`,                 t.cgst);
  totRow(`SGST ${gp}%`,                 t.sgst);
  totRow("TOTAL AMOUNT",                t.grand, false, true);

  // ── Figures ────────────────────────────────────────────────────────────────
  ws.mergeCells(`A${R}:B${R}`); ws.mergeCells(`C${R}:F${R}`);
  cell(ws, `A${R}`, "Total Invoice Value (In Figures)", { bold:true, size:9 });
  cell(ws, `C${R}`, t.grand, { bold:true, size:9, numFmt:"#,##0.00" }); ws.getRow(R).height=15; R++;
  ws.mergeCells(`A${R}:B${R}`); ws.mergeCells(`C${R}:F${R}`);
  cell(ws, `A${R}`, "Total Invoice Value (In Words)", { bold:true, size:9 });
  cell(ws, `C${R}`, toWords(t.grand) + " only", { size:9, italic:true }); ws.getRow(R).height=15; R++;

  // ── Payment ────────────────────────────────────────────────────────────────
  ws.mergeCells(`A${R}:F${R}`);
  cell(ws, `A${R}`, "Payment Terms", { bold:true, size:9, color:"FFFFFFFF", bg:NAVY }); ws.getRow(R).height=16; R++;

  [["In favour of", data.payFavour],["Bank & Branch", data.bankBranch],
   ["Account No.", data.accountNo],["IFSC Code", data.ifscCode],["Comments", data.comments]]
  .forEach(([label, val], i) => {
    const bg = i%2===1 ? SKYBLUE : WHITE;
    ws.mergeCells(`B${R}:F${R}`);
    cell(ws, `A${R}`, label, { bold:true, size:9, bg });
    cell(ws, `B${R}`, val||"", { size:9, bg }); ws.getRow(R).height=15; R++;
  });

  // ── Footer ─────────────────────────────────────────────────────────────────
  ws.mergeCells(`A${R}:F${R}`);
  cell(ws, `A${R}`, data.footerText||"", { size:8, color:"FF64748B", align:"center", bg:LGREY }); ws.getRow(R).height=18;

  ws.printArea = `A1:F${R}`;

  const buf = await wb.xlsx.writeBuffer();
  return buf;
}

module.exports = { genXLSX };
