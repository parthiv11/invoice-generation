"use strict";
const PDFDocument = require("pdfkit");
const { calcTotals, fmtINR, toWords } = require("./invoice-data");

// ── Colour palette ────────────────────────────────────────────────────────────
const NAVY   = "#1F3864";
const SKYBLUE= "#DBEAFE";
const LGREY  = "#F1F5F9";
const ORANGE = "#C65911";
const WHITE  = "#FFFFFF";
const BLACK  = "#111111";
const MGREY  = "#64748B";
const BORDER = "#B8C4D4";
const CLASSIC_BORDER = "#111111";

function renderClassic(doc, data, PW, ML) {
  const t = calcTotals(data);
  let Y = ML;
  const headerH = 60;
  const headerSplit = ML + PW * 0.6;
  const rightW = ML + PW - headerSplit - 10;

  doc.rect(ML, Y, PW, headerH).lineWidth(0.8).stroke(CLASSIC_BORDER);
  doc.fillColor(BLACK).font("Helvetica-Bold").fontSize(26)
     .text((data.supplierName || "INVOICE").toUpperCase(), ML + 10, Y + 16, { width: headerSplit - ML - 16 });
  doc.fillColor(BLACK).font("Helvetica").fontSize(10)
     .text(data.supplierAddress || "", headerSplit + 4, Y + 14, { width: rightW, align: "right" });
  doc.font("Helvetica-Bold").fontSize(10)
     .text(`G.S.T No: ${data.supplierGSTIN || ""}`, headerSplit + 4, Y + 34, { width: rightW, align: "right" });
  Y += headerH;

  const infoH = 70;
  const splitX = ML + PW * 0.58;
  doc.rect(ML, Y, PW, infoH).lineWidth(0.8).stroke(CLASSIC_BORDER);
  doc.moveTo(splitX, Y).lineTo(splitX, Y + infoH).lineWidth(0.6).stroke(CLASSIC_BORDER);

  const leftW = splitX - ML - 16;
  let leftY = Y + 8;
  doc.font("Helvetica-Bold").fontSize(10).text("TO:", ML + 8, leftY);
  leftY += 14;
  const receiverLines = [data.receiverName, data.receiverAddress].filter(line => line && String(line).trim());
  if (receiverLines.length) {
    doc.font("Helvetica-Bold").fontSize(10)
       .text(receiverLines.join("\n"), ML + 8, leftY, { width: leftW });
  }
  doc.font("Helvetica-Bold").fontSize(9.5)
     .text(`GST NO : ${data.receiverGSTIN || ""}`, ML + 8, Y + infoH - 18, { width: leftW });

  const rightX = splitX + 8;
  const infoRightW = ML + PW - rightX - 8;
  let rightY = Y + 10;
  doc.font("Helvetica-Bold").fontSize(10)
     .text(`BILL NO: ${data.invoiceNo || ""}`, rightX, rightY, { width: infoRightW });
  rightY += 16;
  doc.text(`BILL DATE : ${data.invoiceDate || ""}`, rightX, rightY, { width: infoRightW });
  rightY += 16;
  doc.font("Helvetica").fontSize(9.5)
     .text(`Bill for the month of : ${data.period || ""}`, rightX, rightY, { width: infoRightW });
  Y += infoH;

  const colWidths = [PW * 0.44, PW * 0.18, PW * 0.18, PW * 0.2];
  const colX = [ML];
  colWidths.forEach((w, i) => { if (i) colX[i] = colX[i - 1] + colWidths[i - 1]; });

  function drawRow(y, h, values, opts = {}) {
    const { bold = false, align = [], fontSize = 9.5 } = opts;
    doc.rect(ML, y, PW, h).lineWidth(0.6).stroke(CLASSIC_BORDER);
    colX.slice(1).forEach(x => {
      doc.moveTo(x, y).lineTo(x, y + h).lineWidth(0.6).stroke(CLASSIC_BORDER);
    });
    values.forEach((val, i) => {
      doc.font(bold ? "Helvetica-Bold" : "Helvetica").fontSize(fontSize).fillColor(BLACK)
         .text(val || "", colX[i] + 4, y + 5, { width: colWidths[i] - 8, align: align[i] || "left" });
    });
  }

  drawRow(Y, 22, ["Particulars", "Quantity/LTR.", "Rate", "Amount Rs."], { bold: true, align: ["center", "center", "center", "center"] });
  Y += 22;

  const fmtQty = n => (n ? fmtINR(n, 0) : "");
  const fmtRate = n => (n ? fmtINR(n, 2) : "");
  const fmtAmt = n => (n || n === 0 ? fmtINR(n, 0) : "");
  const minRows = 6;
  const items = t.lineItems;
  const rowsToRender = Math.max(items.length, minRows);

  for (let i = 0; i < rowsToRender; i++) {
    const item = items[i];
    const desc = item?.desc || "";
    const persons = item && item.persons !== "" && item.persons !== undefined ? +item.persons : 0;
    const days = item && item.days !== "" && item.days !== undefined ? +item.days : 1;
    const qty = persons ? persons * (days || 1) : "";
    const rate = item && item.rate ? +item.rate : 0;
    const amt = item && item._amount ? item._amount : 0;
    drawRow(Y, 20, [desc, fmtQty(qty), fmtRate(rate), amt ? fmtAmt(amt) : ""], { align: ["left", "center", "right", "right"] });
    Y += 20;
  }

  const cp = data.cgstPct ?? 9;
  const gp = data.sgstPct ?? 9;
  const totals = [
    ["", "", "TOTAL", fmtAmt(t.sub)],
    ["", "", `CGST ${cp}%`, fmtAmt(t.cgst)],
    ["", "", `SGST ${gp}%`, fmtAmt(t.sgst)],
    ["", "", "GRAND TOTAL", fmtAmt(t.grand)],
  ];
  totals.forEach(row => {
    drawRow(Y, 20, row, { align: ["left", "center", "left", "right"], bold: true });
    Y += 20;
  });

  doc.rect(ML, Y, PW, 20).lineWidth(0.8).stroke(CLASSIC_BORDER);
  doc.font("Helvetica-Bold").fontSize(9.5)
     .text(`Rupees in Words: ${toWords(t.grand).toUpperCase()} ONLY`, ML + 6, Y + 5, { width: PW - 12 });
  Y += 20;

  const footH = 80;
  const footSplit = ML + PW * 0.6;
  doc.rect(ML, Y, PW, footH).lineWidth(0.8).stroke(CLASSIC_BORDER);
  doc.moveTo(footSplit, Y).lineTo(footSplit, Y + footH).lineWidth(0.6).stroke(CLASSIC_BORDER);

  const stateCode = ((data.placeOfSupply || "").match(/\b(\d{2})\b/) || [])[1] || "";
  let fy = Y + 10;
  doc.font("Helvetica-Bold").fontSize(9.5)
     .text(`PAN: ${data.supplierPAN || ""}`, ML + 8, fy);
  fy += 14;
  doc.text("P.TAX NO : ", ML + 8, fy);
  fy += 14;
  doc.text(`GST NO : ${data.supplierGSTIN || ""}`, ML + 8, fy);
  fy += 14;
  doc.text(`STATE CODE : ${stateCode}`, ML + 8, fy);

  const rightX2 = footSplit + 8;
  const rightW2 = ML + PW - rightX2 - 8;
  doc.font("Helvetica-Bold").fontSize(10)
     .text(`For ${data.supplierName || ""}`, rightX2, Y + footH - 28, { width: rightW2, align: "right" });
  doc.font("Helvetica").fontSize(9)
     .text("Authorized signatory", rightX2, Y + footH - 14, { width: rightW2, align: "right" });
}

function hex(h) {
  const r = parseInt(h.slice(1,3),16)/255;
  const g = parseInt(h.slice(3,5),16)/255;
  const b = parseInt(h.slice(5,7),16)/255;
  return [r,g,b];
}

function genPDF(data) {
  return new Promise((resolve, reject) => {
    const doc  = new PDFDocument({ size:"A4", margin:28, compress:true, info:{ Title:`Invoice ${data.invoiceNo}`, Author: data.supplierName } });
    const bufs = [];
    doc.on("data", d => bufs.push(d));
    doc.on("end",  () => resolve(Buffer.concat(bufs)));
    doc.on("error", reject);

    const PW = doc.page.width  - 56; // usable width
    const ML = 28;                   // margin left
    const isClassic = (data.billVariant || "modern") === "classic";
    if (isClassic) {
      renderClassic(doc, data, PW, ML);
      doc.end();
      return;
    }
    const theme = {
      titleBg: NAVY,
      titleFg: WHITE,
      titleText: "INVOICE",
      titleSize: 16,
      titleTracking: 4,
      tagBg: WHITE,
      tagFg: ORANGE,
      tagText: "ORIGINAL FOR RECIPIENT",
      mainBg: NAVY,
      altBg: SKYBLUE,
      softBg: LGREY,
      border: BORDER,
      label: NAVY,
      text: BLACK,
      footer: MGREY,
    };

    // ── helpers ──────────────────────────────────────────────────────────────
    function fillRect(x,y,w,h,color){
      doc.rect(x,y,w,h).fill(color);
    }
    function strokeRect(x,y,w,h,color="#B8C4D4",lw=0.5){
      doc.rect(x,y,w,h).lineWidth(lw).stroke(color);
    }
    function text(str,x,y,opts={}){
      doc.fillColor(opts.color||BLACK).fontSize(opts.size||8.5)
         .font(opts.bold?"Helvetica-Bold":"Helvetica")
         .text(str,x,y,{...opts,continued:false});
    }

    // ── TITLE ─────────────────────────────────────────────────────────────────
    let Y = ML;
    fillRect(ML, Y, PW, 26, theme.titleBg);
    doc.fillColor(theme.titleFg).fontSize(theme.titleSize).font("Helvetica-Bold")
       .text(theme.titleText, ML, Y+5, {width:PW, align:"center", characterSpacing: theme.titleTracking});
    Y += 26;

    // ── ORIGINAL TAG ──────────────────────────────────────────────────────────
    fillRect(ML, Y, PW, 14, theme.tagBg);
    doc.fillColor(theme.tagFg).fontSize(7.5).font("Helvetica-Bold")
       .text(theme.tagText, ML, Y+3, {width:PW, align:"center"});
    strokeRect(ML, Y, PW, 14, theme.border);
    Y += 14;

    // ── PARTIES ───────────────────────────────────────────────────────────────
    const partyH = 78;
    const half   = PW / 2;

    // Supplier header
    fillRect(ML,        Y, half, 14, theme.mainBg);
    fillRect(ML + half, Y, half, 14, theme.mainBg);
    doc.fillColor(theme.titleFg).fontSize(8).font("Helvetica-Bold")
       .text("Supplier", ML,         Y+3, {width:half,  align:"center"})
       .text("Receiver", ML + half,  Y+3, {width:half,  align:"center"});
    Y += 14;

    fillRect(ML,        Y, half, partyH, WHITE);
    fillRect(ML + half, Y, half, partyH, WHITE);
    strokeRect(ML, Y - 14, PW, partyH + 14, theme.border);
    // vertical divider
    doc.moveTo(ML + half, Y - 14).lineTo(ML + half, Y + partyH).lineWidth(0.5).stroke(theme.border);

    function partyField(label, val, x, yy, w) {
      doc.fillColor(theme.label).fontSize(7.5).font("Helvetica-Bold").text(label, x+4, yy, {width:52});
      doc.fillColor(theme.text).fontSize(8).font("Helvetica").text(val||"", x+56, yy, {width:w-60, lineBreak:false});
    }

    let sy = Y + 5;
    partyField("Name",    data.supplierName,    ML, sy, half); sy += 13;
    partyField("GSTIN",   data.supplierGSTIN,   ML, sy, half); sy += 13;
    partyField("PAN",     data.supplierPAN,     ML, sy, half); sy += 13;
    doc.fillColor(theme.label).fontSize(7.5).font("Helvetica-Bold").text("Address", ML+4, sy, {width:52});
    doc.fillColor(theme.text).fontSize(7).font("Helvetica")
       .text(data.supplierAddress||"", ML+56, sy, {width:half-62, lineBreak:true});

    let ry = Y + 5;
    partyField("Name",   data.receiverName,  ML+half, ry, half); ry += 13;
    partyField("GSTIN",  data.receiverGSTIN, ML+half, ry, half); ry += 13;
    partyField("PERIOD", data.period,        ML+half, ry, half);
    Y += partyH;

    // ── META ROW ──────────────────────────────────────────────────────────────
    const metaH = 26; const third = PW / 3;
    fillRect(ML,           Y, PW, metaH, theme.softBg);
    strokeRect(ML, Y, PW, metaH, theme.border);
    doc.moveTo(ML+third,Y).lineTo(ML+third,Y+metaH).lineWidth(0.5).stroke(theme.border);
    doc.moveTo(ML+third*2,Y).lineTo(ML+third*2,Y+metaH).lineWidth(0.5).stroke(theme.border);

    function metaCell(label, val, x, w) {
      doc.fillColor(theme.label).fontSize(7).font("Helvetica-Bold")
         .text(label.toUpperCase(), x+4, Y+4, {width:w-8});
      doc.fillColor(theme.text).fontSize(9).font("Helvetica-Bold")
         .text(val||"", x+4, Y+14, {width:w-8});
    }
    metaCell("Place of Supply", data.placeOfSupply, ML, third);
    metaCell("Invoice Date",    data.invoiceDate,   ML+third, third);
    metaCell("Invoice No.",     data.invoiceNo,     ML+third*2, third);
    Y += metaH;

    // ── TABLE HEADER ──────────────────────────────────────────────────────────
    const cols = [30, PW-30-74-74-36-80, 74, 74, 36, 80];
    const colX = [];
    let cx = ML;
    cols.forEach(w => { colX.push(cx); cx += w; });
    const hdrs = ["S.No.","Description & SAC Code","Rate/Head","No. of Persons","Days","Total Amt"];

    fillRect(ML, Y, PW, 20, theme.mainBg);
    hdrs.forEach((h,i) => {
      doc.fillColor(theme.titleFg).fontSize(7.5).font("Helvetica-Bold")
         .text(h, colX[i]+2, Y+4, {width:cols[i]-4, align:i===1?"left":"center"});
    });
    Y += 20;

    // ── TABLE ROWS ────────────────────────────────────────────────────────────
    const t = calcTotals(data);
    t.lineItems.forEach((item, i) => {
      const rh = 16;
      const bg = i % 2 === 1 ? theme.altBg : WHITE;
      fillRect(ML, Y, PW, rh, bg);
      strokeRect(ML, Y, PW, rh, theme.border);
      colX.forEach((x,ci) => {
        if(ci>0) doc.moveTo(x,Y).lineTo(x,Y+rh).lineWidth(0.3).stroke(theme.border);
      });

      const vals = [
        String(i+1),
        item.desc||"",
        item.rate ? fmtINR(item.rate,0) : "",
        item.persons ? fmtINR(item.persons,0) : "",
        item.days && +item.days ? String(item.days) : "",
        item._amount ? fmtINR(item._amount) : "",
      ];
      vals.forEach((v,ci) => {
        const align = ci === 0 ? "center" : ci === 1 ? "left" : "right";
        doc.fillColor(BLACK).fontSize(8).font("Helvetica")
           .text(v, colX[ci]+2, Y+4, {width:cols[ci]-4, align, lineBreak:false});
      });
      // fix S.No. center
      doc.fillColor(BLACK).fontSize(8).font("Helvetica")
         .text(vals[0], colX[0]+2, Y+4, {width:cols[0]-4, align:"center", lineBreak:false});
      Y += rh;
    });

    // ── TOTALS ROWS ───────────────────────────────────────────────────────────
    function totRow(label, val, highlight=false, grand=false) {
      const h = 15;
      const bg = grand ? theme.mainBg : highlight ? theme.altBg : WHITE;
      const fg = grand ? theme.titleFg : BLACK;
      fillRect(ML, Y, PW, h, bg);
      strokeRect(ML, Y, PW, h, theme.border);
      doc.fillColor(grand ? theme.titleFg : theme.label).fontSize(8).font("Helvetica-Bold")
         .text(label, ML+4, Y+3.5, {width:PW-84});
      doc.moveTo(ML+PW-80, Y).lineTo(ML+PW-80, Y+h).lineWidth(0.4).stroke(theme.border);
      doc.fillColor(fg).fontSize(8).font("Helvetica-Bold")
         .text(val, ML+PW-78, Y+3.5, {width:74, align:"right"});
      Y += h;
    }

    const sp = data.svcPct??5, cp = data.cgstPct??9, gp = data.sgstPct??9;
    totRow("Total Taxable Amount",         fmtINR(t.taxable), true);
    totRow(`Service Charges @${sp}%`,      fmtINR(t.svc));
    totRow("Total",                         fmtINR(t.sub), true);
    totRow(`CGST ${cp}%`,                  fmtINR(t.cgst));
    totRow(`SGST ${gp}%`,                  fmtINR(t.sgst), false);
    totRow("TOTAL AMOUNT",                 fmtINR(t.grand), false, true);

    // ── FIGURES ───────────────────────────────────────────────────────────────
    function figRow(label, val, italic=false) {
      const h = 15;
      fillRect(ML, Y, PW, h, WHITE);
      strokeRect(ML, Y, PW, h, theme.border);
      doc.moveTo(ML+155, Y).lineTo(ML+155, Y+h).lineWidth(0.4).stroke(theme.border);
      doc.fillColor(theme.label).fontSize(8).font("Helvetica-Bold").text(label, ML+4, Y+3.5, {width:148});
      doc.fillColor(BLACK).fontSize(8).font(italic?"Helvetica-Oblique":"Helvetica-Bold")
         .text(val, ML+158, Y+3.5, {width:PW-162});
      Y += h;
    }
    figRow("Total Invoice Value (In Figures)", fmtINR(t.grand));
    figRow("Total Invoice Value (In Words)",   toWords(t.grand) + " only", true);

    // ── PAYMENT ───────────────────────────────────────────────────────────────
    fillRect(ML, Y, PW, 14, theme.mainBg);
    doc.fillColor(theme.titleFg).fontSize(8).font("Helvetica-Bold").text("Payment Terms", ML+4, Y+3);
    Y += 14;

    const payRows = [
      ["In favour of", data.payFavour],
      ["Bank & Branch", data.bankBranch],
      ["Account No.",   data.accountNo],
      ["IFSC Code",     data.ifscCode],
      ["Comments",      data.comments],
    ];
    payRows.forEach(([label,val],i) => {
      const h = 14;
      const bg = i%2===1 ? theme.altBg : WHITE;
      fillRect(ML, Y, PW, h, bg);
      strokeRect(ML, Y, PW, h, theme.border);
      doc.moveTo(ML+90, Y).lineTo(ML+90, Y+h).lineWidth(0.4).stroke(theme.border);
      doc.fillColor(theme.label).fontSize(7.5).font("Helvetica-Bold").text(label, ML+4, Y+3, {width:84});
      doc.fillColor(BLACK).fontSize(8).font("Helvetica").text(val||"", ML+94, Y+3, {width:PW-98});
      Y += h;
    });

    // ── SIGNATURE ─────────────────────────────────────────────────────────────
    const sigH = 64;
    fillRect(ML,      Y, half, sigH, theme.softBg);
    fillRect(ML+half, Y, half, sigH, theme.altBg);
    strokeRect(ML, Y, PW, sigH, theme.border);
    doc.moveTo(ML+half, Y).lineTo(ML+half, Y+sigH).lineWidth(0.5).stroke(theme.border);
    doc.fillColor(theme.label).fontSize(8).font("Helvetica-Bold")
       .text("Authorised Signatory", ML, Y+sigH-14, {width:half, align:"center"})
       .text("by " + data.supplierName, ML+half, Y+sigH-14, {width:half, align:"center"});
    Y += sigH;

    // ── FOOTER ────────────────────────────────────────────────────────────────
    fillRect(ML, Y, PW, 18, theme.softBg);
    strokeRect(ML, Y, PW, 18, theme.border);
    doc.fillColor(theme.footer).fontSize(7.5).font("Helvetica")
       .text(data.footerText||"", ML, Y+5, {width:PW, align:"center"});

    doc.end();
  });
}

module.exports = { genPDF };
