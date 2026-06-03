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
  const isClassic = (data.billVariant || "modern") === "classic";
  const theme = isClassic
    ? {
        titleBg: WHITE,
        titleFg: "FF111111",
        titleText: data.supplierName || "INVOICE",
        tagBg: WHITE,
        tagFg: "FF111111",
        tagText: "BILL",
        mainBg: WHITE,
        mainFg: "FF111111",
        softBg: WHITE,
        altBg: WHITE,
        labelFg: "FF111111",
      }
    : {
        titleBg: NAVY,
        titleFg: "FFFFFFFF",
        titleText: "INVOICE",
        tagBg: WHITE,
        tagFg: ORANGE,
        tagText: "ORIGINAL FOR RECIPIENT",
        mainBg: NAVY,
        mainFg: "FFFFFFFF",
        softBg: LGREY,
        altBg: SKYBLUE,
        labelFg: NAVY,
      };

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
  cell(ws, `A${R}`, theme.titleText, { bold:true, size:isClassic ? 13 : 16, color:theme.titleFg, bg:theme.titleBg, align:"center" });
  ws.getRow(R).height = 26; R++;

  // ── Original for recipient ─────────────────────────────────────────────────
  ws.mergeCells(`A${R}:F${R}`);
  cell(ws, `A${R}`, theme.tagText, { bold:true, size:8, color:theme.tagFg, align:"center", bg:theme.tagBg });
  ws.getRow(R).height = 14; R++;

  // ── Party headers ──────────────────────────────────────────────────────────
  ws.mergeCells(`A${R}:C${R}`);
  ws.mergeCells(`D${R}:F${R}`);
  cell(ws, `A${R}`, "Supplier", { bold:true, size:9, color:theme.mainFg, bg:theme.mainBg, align:"center" });
  cell(ws, `D${R}`, "Receiver", { bold:true, size:9, color:theme.mainFg, bg:theme.mainBg, align:"center" });
  ws.getRow(R).height = 16; R++;

  // ── Supplier info ──────────────────────────────────────────────────────────
  function partyRow(r, labelA, valA, labelD, valD) {
    ws.mergeCells(`B${r}:C${r}`); ws.mergeCells(`E${r}:F${r}`);
    cell(ws, `A${r}`, labelA, { bold:true, size:9, color:theme.labelFg, bg:WHITE });
    cell(ws, `B${r}`, valA,   { size:9, bg:WHITE, wrap:true });
    if (labelD) cell(ws, `D${r}`, labelD, { bold:true, size:9, color:theme.labelFg, bg:WHITE });
    if (valD)   cell(ws, `E${r}`, valD,   { size:9, bg:WHITE });
  }

  partyRow(R,   "Name",    data.supplierName,    "Name",   data.receiverName);   R++;
  partyRow(R,   "GSTIN",   data.supplierGSTIN,   "GSTIN",  data.receiverGSTIN);  R++;
  partyRow(R,   "PAN",     data.supplierPAN,     "PERIOD", data.period);         R++;
  ws.mergeCells(`B${R}:C${R}`); ws.mergeCells(`D${R}:F${R}`);
  cell(ws, `A${R}`, "Address", { bold:true, size:9, color:theme.labelFg, bg:WHITE });
  const addrCell = cell(ws, `B${R}`, data.supplierAddress, { size:8, bg:WHITE, wrap:true });
  ws.getRow(R).height = 28; R++;

  // ── Meta row ──────────────────────────────────────────────────────────────
  ws.mergeCells(`A${R}:B${R}`); ws.mergeCells(`C${R}:D${R}`); ws.mergeCells(`E${R}:F${R}`);
  cell(ws, `A${R}`, "Place of Supply", { bold:true, size:9, color:theme.labelFg, bg:theme.softBg, align:"center" });
  cell(ws, `C${R}`, "Invoice Date",    { bold:true, size:9, color:theme.labelFg, bg:theme.softBg, align:"center" });
  cell(ws, `E${R}`, "Invoice No.",     { bold:true, size:9, color:theme.labelFg, bg:theme.softBg, align:"center" });
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
      cell(ws, `${col}${R}`, h, { bold:true, size:9, color:theme.mainFg, bg:theme.mainBg, align:"center", wrap:true });
    });
  ws.getRow(R).height = 22; R++;

  // ── Line items ─────────────────────────────────────────────────────────────
  const t = calcTotals(data);
  const startRow = R;
  t.lineItems.forEach((item, i) => {
    const bg = i % 2 === 1 ? theme.altBg : WHITE;
    const cols = ["A","B","C","D","E","F"];
    const amountValue = item.rate && item.persons
      ? { formula: `C${R}*D${R}`, result: item._amount || 0 }
      : "";
    const vals = [i+1, item.desc||"", item.rate||"", item.persons||"", item.days&&+item.days?item.days:"", amountValue];
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
  const pastBalance = +(data.pastBalanceDue ?? 0) || 0;
  const taxableRow = R;

  function formulaTotRow(label, formula, result, hilight=false, grand=false) {
    ws.mergeCells(`A${R}:E${R}`);
    const bg = grand ? theme.mainBg : hilight ? theme.altBg : WHITE;
    const clr = grand ? theme.mainFg : theme.labelFg;
    cell(ws, `A${R}`, label, { bold:true, size:9, color:clr, bg, align:"left" });
    cell(ws, `F${R}`, { formula, result }, { bold:true, size:9, color:clr, bg, align:"right", numFmt:"#,##0.00" });
    ws.getRow(R).height = 15; R++;
  }

  function totRow(label, val, hilight=false, grand=false) {
    ws.mergeCells(`A${R}:E${R}`);
    const bg = grand ? theme.mainBg : hilight ? theme.altBg : WHITE;
    const clr = grand ? theme.mainFg : theme.labelFg;
    cell(ws, `A${R}`, label, { bold:true, size:9, color:clr, bg, align:"left" });
    cell(ws, `F${R}`, val, { bold:true, size:9, color:clr, bg, align:"right", numFmt:"#,##0.00" });
    ws.getRow(R).height = 15; R++;
  }

  formulaTotRow("Total Taxable Amount", `SUM(F${startRow}:F${R - 1})`, t.taxable, true);
  const serviceRow = R;
  formulaTotRow(`Service Charges @${sp}%`, `F${taxableRow}*${sp / 100}`, t.svc);
  const subRow = R;
  formulaTotRow("Total Amount Before Tax", `SUM(F${taxableRow}:F${serviceRow})`, t.sub, true);
  const cgstRow = R;
  formulaTotRow(`CGST ${cp}%`, `${cp / 100}*F${subRow}`, t.cgst);
  const sgstRow = R;
  formulaTotRow(`SGST ${gp}%`, `${gp / 100}*F${subRow}`, t.sgst);
  const pastBalanceRow = R;
  totRow("PAST BALANCE DUE", pastBalance, true);
  const grandRow = R;
  formulaTotRow("TOTAL AMOUNT", `SUM(F${subRow}:F${pastBalanceRow})`, t.grand, false, true);

  // ── Figures ────────────────────────────────────────────────────────────────
  const roundedGrand = Math.ceil(t.grand);
  ws.mergeCells(`A${R}:B${R}`); ws.mergeCells(`C${R}:F${R}`);
  cell(ws, `A${R}`, "Total Invoice Value (In Figures)", { bold:true, size:9 });
  cell(ws, `C${R}`, { formula: `ROUNDUP(F${grandRow},0)`, result: roundedGrand }, { bold:true, size:9, numFmt:"#,##0.00" }); ws.getRow(R).height=15; R++;
  ws.mergeCells(`A${R}:B${R}`); ws.mergeCells(`C${R}:F${R}`);
  cell(ws, `A${R}`, "Total Invoice Value (In Words)", { bold:true, size:9 });
  cell(ws, `C${R}`, toWords(roundedGrand) + " only", { size:9, italic:true }); ws.getRow(R).height=15; R++;

  // ── Payment ────────────────────────────────────────────────────────────────
  ws.mergeCells(`A${R}:F${R}`);
  cell(ws, `A${R}`, "Payment Terms", { bold:true, size:9, color:theme.mainFg, bg:theme.mainBg }); ws.getRow(R).height=16; R++;

  [["In favour of", data.payFavour],["Bank & Branch", data.bankBranch],
   ["Account No.", data.accountNo],["IFSC Code", data.ifscCode],["Comments", data.comments]]
  .forEach(([label, val], i) => {
    const bg = i%2===1 ? theme.altBg : WHITE;
    ws.mergeCells(`B${R}:F${R}`);
    cell(ws, `A${R}`, label, { bold:true, size:9, bg });
    cell(ws, `B${R}`, val||"", { size:9, bg }); ws.getRow(R).height=15; R++;
  });

  // ── Signature ──────────────────────────────────────────────────────────────
  const sigStart = R;
  for (let row = 0; row < 4; row++) {
    const current = sigStart + row;
    const leftFill = row === 3 ? theme.softBg : WHITE;
    const rightFill = row === 3 ? theme.altBg : WHITE;
    ["A", "B", "C"].forEach(col => cell(ws, `${col}${current}`, "", { bg:leftFill }));
    ["D", "E", "F"].forEach(col => cell(ws, `${col}${current}`, "", { bg:rightFill }));
    ws.getRow(current).height = row === 3 ? 18 : 16;
  }
  ws.mergeCells(`A${sigStart}:C${sigStart + 2}`);
  ws.mergeCells(`D${sigStart}:F${sigStart + 2}`);
  ws.mergeCells(`A${sigStart + 3}:C${sigStart + 3}`);
  ws.mergeCells(`D${sigStart + 3}:F${sigStart + 3}`);
  cell(ws, `A${sigStart}`, "", { bg:WHITE });
  cell(ws, `D${sigStart}`, "", { bg:WHITE });
  cell(ws, `A${sigStart + 3}`, "Authorised Signatory", { bold:true, size:9, color:theme.labelFg, bg:theme.softBg, align:"center" });
  cell(ws, `D${sigStart + 3}`, `by ${data.supplierName}`, { bold:true, size:9, color:theme.labelFg, bg:theme.altBg, align:"center" });
  R += 4;

  // ── Footer ─────────────────────────────────────────────────────────────────
  ws.mergeCells(`A${R}:F${R}`);
  cell(ws, `A${R}`, data.footerText||"", { size:8, color:isClassic ? "FF444444" : "FF64748B", align:"center", bg:theme.softBg }); ws.getRow(R).height=18;

  ws.printArea = `A1:F${R}`;

  const buf = await wb.xlsx.writeBuffer();
  return buf;
}

module.exports = { genXLSX };
