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
    fillRect(ML, Y, PW, 26, NAVY);
    doc.fillColor(WHITE).fontSize(16).font("Helvetica-Bold")
       .text("INVOICE", ML, Y+5, {width:PW, align:"center"});
    Y += 26;

    // ── ORIGINAL TAG ──────────────────────────────────────────────────────────
    fillRect(ML, Y, PW, 14, WHITE);
    doc.fillColor(ORANGE).fontSize(7.5).font("Helvetica-Bold")
       .text("ORIGINAL FOR RECIPIENT", ML, Y+3, {width:PW, align:"right"});
    strokeRect(ML, Y, PW, 14);
    Y += 14;

    // ── PARTIES ───────────────────────────────────────────────────────────────
    const partyH = 78;
    const half   = PW / 2;

    // Supplier header
    fillRect(ML,        Y, half, 14, NAVY);
    fillRect(ML + half, Y, half, 14, NAVY);
    doc.fillColor(WHITE).fontSize(8).font("Helvetica-Bold")
       .text("Supplier", ML,         Y+3, {width:half,  align:"center"})
       .text("Receiver", ML + half,  Y+3, {width:half,  align:"center"});
    Y += 14;

    fillRect(ML,        Y, half, partyH, WHITE);
    fillRect(ML + half, Y, half, partyH, WHITE);
    strokeRect(ML, Y - 14, PW, partyH + 14);
    // vertical divider
    doc.moveTo(ML + half, Y - 14).lineTo(ML + half, Y + partyH).lineWidth(0.5).stroke(BORDER);

    function partyField(label, val, x, yy, w) {
      doc.fillColor(NAVY).fontSize(7.5).font("Helvetica-Bold").text(label, x+4, yy, {width:52});
      doc.fillColor(BLACK).fontSize(8).font("Helvetica").text(val||"", x+56, yy, {width:w-60, lineBreak:false});
    }

    let sy = Y + 5;
    partyField("Name",    data.supplierName,    ML, sy, half); sy += 13;
    partyField("GSTIN",   data.supplierGSTIN,   ML, sy, half); sy += 13;
    partyField("PAN",     data.supplierPAN,     ML, sy, half); sy += 13;
    doc.fillColor(NAVY).fontSize(7.5).font("Helvetica-Bold").text("Address", ML+4, sy, {width:52});
    doc.fillColor(BLACK).fontSize(7).font("Helvetica")
       .text(data.supplierAddress||"", ML+56, sy, {width:half-62, lineBreak:true});

    let ry = Y + 5;
    partyField("Name",   data.receiverName,  ML+half, ry, half); ry += 13;
    partyField("GSTIN",  data.receiverGSTIN, ML+half, ry, half); ry += 13;
    partyField("PERIOD", data.period,        ML+half, ry, half);
    Y += partyH;

    // ── META ROW ──────────────────────────────────────────────────────────────
    const metaH = 26; const third = PW / 3;
    fillRect(ML,           Y, PW, metaH, LGREY);
    strokeRect(ML, Y, PW, metaH);
    doc.moveTo(ML+third,Y).lineTo(ML+third,Y+metaH).lineWidth(0.5).stroke(BORDER);
    doc.moveTo(ML+third*2,Y).lineTo(ML+third*2,Y+metaH).lineWidth(0.5).stroke(BORDER);

    function metaCell(label, val, x, w) {
      doc.fillColor(NAVY).fontSize(7).font("Helvetica-Bold")
         .text(label.toUpperCase(), x+4, Y+4, {width:w-8});
      doc.fillColor(BLACK).fontSize(9).font("Helvetica-Bold")
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

    fillRect(ML, Y, PW, 20, NAVY);
    hdrs.forEach((h,i) => {
      doc.fillColor(WHITE).fontSize(7.5).font("Helvetica-Bold")
         .text(h, colX[i]+2, Y+4, {width:cols[i]-4, align:i===1?"left":"center"});
    });
    Y += 20;

    // ── TABLE ROWS ────────────────────────────────────────────────────────────
    const t = calcTotals(data);
    t.lineItems.forEach((item, i) => {
      const rh = 16;
      const bg = i % 2 === 1 ? SKYBLUE : WHITE;
      fillRect(ML, Y, PW, rh, bg);
      strokeRect(ML, Y, PW, rh);
      colX.forEach((x,ci) => {
        if(ci>0) doc.moveTo(x,Y).lineTo(x,Y+rh).lineWidth(0.3).stroke(BORDER);
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
        doc.fillColor(BLACK).fontSize(8).font("Helvetica")
           .text(v, colX[ci]+2, Y+4, {width:cols[ci]-4, align:ci<=1?"center":"right", lineBreak:false});
        if(ci===1) doc.text(v, colX[ci]+3, Y+4, {width:cols[ci]-5, align:"left", lineBreak:false});
      });
      // fix S.No. center
      doc.fillColor(BLACK).fontSize(8).font("Helvetica")
         .text(vals[0], colX[0]+2, Y+4, {width:cols[0]-4, align:"center", lineBreak:false});
      Y += rh;
    });

    // ── TOTALS ROWS ───────────────────────────────────────────────────────────
    function totRow(label, val, highlight=false, grand=false) {
      const h = 15;
      const bg = grand ? NAVY : highlight ? SKYBLUE : WHITE;
      const fg = grand ? WHITE : BLACK;
      fillRect(ML, Y, PW, h, bg);
      strokeRect(ML, Y, PW, h);
      doc.fillColor(grand? WHITE : NAVY).fontSize(8).font("Helvetica-Bold")
         .text(label, ML+4, Y+3.5, {width:PW-84});
      doc.moveTo(ML+PW-80, Y).lineTo(ML+PW-80, Y+h).lineWidth(0.4).stroke(BORDER);
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
      strokeRect(ML, Y, PW, h);
      doc.moveTo(ML+155, Y).lineTo(ML+155, Y+h).lineWidth(0.4).stroke(BORDER);
      doc.fillColor(NAVY).fontSize(8).font("Helvetica-Bold").text(label, ML+4, Y+3.5, {width:148});
      doc.fillColor(BLACK).fontSize(8).font(italic?"Helvetica-Oblique":"Helvetica-Bold")
         .text(val, ML+158, Y+3.5, {width:PW-162});
      Y += h;
    }
    figRow("Total Invoice Value (In Figures)", fmtINR(t.grand));
    figRow("Total Invoice Value (In Words)",   toWords(t.grand) + " only", true);

    // ── PAYMENT ───────────────────────────────────────────────────────────────
    fillRect(ML, Y, PW, 14, NAVY);
    doc.fillColor(WHITE).fontSize(8).font("Helvetica-Bold").text("Payment Terms", ML+4, Y+3);
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
      const bg = i%2===1 ? SKYBLUE : WHITE;
      fillRect(ML, Y, PW, h, bg);
      strokeRect(ML, Y, PW, h);
      doc.moveTo(ML+90, Y).lineTo(ML+90, Y+h).lineWidth(0.4).stroke(BORDER);
      doc.fillColor(NAVY).fontSize(7.5).font("Helvetica-Bold").text(label, ML+4, Y+3, {width:84});
      doc.fillColor(BLACK).fontSize(8).font("Helvetica").text(val||"", ML+94, Y+3, {width:PW-98});
      Y += h;
    });

    // ── SIGNATURE ─────────────────────────────────────────────────────────────
    const sigH = 40;
    fillRect(ML,      Y, half, sigH, LGREY);
    fillRect(ML+half, Y, half, sigH, SKYBLUE);
    strokeRect(ML, Y, PW, sigH);
    doc.moveTo(ML+half, Y).lineTo(ML+half, Y+sigH).lineWidth(0.5).stroke(BORDER);
    doc.fillColor(NAVY).fontSize(8).font("Helvetica-Bold")
       .text("Authorised Signatory", ML, Y+sigH-12, {width:half, align:"center"})
       .text("by " + data.supplierName, ML+half, Y+sigH-12, {width:half, align:"center"});
    Y += sigH;

    // ── FOOTER ────────────────────────────────────────────────────────────────
    fillRect(ML, Y, PW, 18, LGREY);
    strokeRect(ML, Y, PW, 18);
    doc.fillColor(MGREY).fontSize(7.5).font("Helvetica")
       .text(data.footerText||"", ML, Y+5, {width:PW, align:"center"});

    doc.end();
  });
}

module.exports = { genPDF };
