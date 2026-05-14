"use strict";
const express  = require("express");
const cors     = require("cors");
const path     = require("path");
const { mergeWithDefaults, calcTotals, fmtINR, toWords, DEFAULTS } = require("./lib/invoice-data");
const { genPDF }  = require("./lib/gen-pdf");
const { genXLSX } = require("./lib/gen-xlsx");

const app = express();
app.use(cors());
app.use(express.json({ limit: "2mb" }));
app.use(express.static(path.join(__dirname, "public")));

// ── Docs ──────────────────────────────────────────────────────────────────────
app.get("/api", (req, res) => res.json({
  name: "KISHAN ENTERPRISE — Invoice API",
  version: "2.0.0",
  endpoints: {
    "GET  /api":              "This help page",
    "GET  /api/health":       "Health check",
    "GET  /api/defaults":     "Default invoice fields",
    "POST /api/invoice/pdf":  "→ PDF file download",
    "POST /api/invoice/xlsx": "→ Excel file download",
    "POST /api/invoice/json": "→ Calculated JSON response",
  },
  fields: {
    invoiceNo:       "string  — e.g. 'RE-39'",
    invoiceDate:     "string  — e.g. '01-04-2026'",
    billVariant:     "string  — 'modern' (default) or 'classic'",
    placeOfSupply:   "string",
    supplierName:    "string",
    supplierGSTIN:   "string",
    supplierPAN:     "string",
    supplierAddress: "string",
    receiverName:    "string",
    receiverGSTIN:   "string",
    period:          "string  — e.g. '01/03/2026 to 31/03/2026'",
    svcPct:          "number  — service charge % (default 5)",
    cgstPct:         "number  — CGST % (default 9)",
    sgstPct:         "number  — SGST % (default 9)",
    items:           "array   — [{desc, rate, persons, days}]",
    payFavour:       "string",
    bankBranch:      "string",
    accountNo:       "string",
    ifscCode:        "string",
    comments:        "string",
    footerText:      "string",
  },
  note: "All unspecified fields fall back to KISHAN ENTERPRISE defaults. Only items[] replaces entirely."
}));

app.get("/api/health",   (_, r) => r.json({ status: "ok", ts: new Date().toISOString() }));
app.get("/api/defaults", (_, r) => r.json(DEFAULTS));

// ── POST /api/invoice/pdf ──────────────────────────────────────────────────
app.post("/api/invoice/pdf", async (req, res) => {
  try {
    const data = mergeWithDefaults(req.body);
    const buf  = await genPDF(data);
    const name = `invoice-${data.invoiceNo || "draft"}.pdf`;
    res.setHeader("Content-Type", "application/pdf");
    res.setHeader("Content-Disposition", `attachment; filename="${name}"`);
    res.send(buf);
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message });
  }
});

// ── POST /api/invoice/xlsx ─────────────────────────────────────────────────
app.post("/api/invoice/xlsx", async (req, res) => {
  try {
    const data = mergeWithDefaults(req.body);
    const buf  = await genXLSX(data);
    const name = `invoice-${data.invoiceNo || "draft"}.xlsx`;
    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    res.setHeader("Content-Disposition", `attachment; filename="${name}"`);
    res.send(buf);
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message });
  }
});

// ── POST /api/invoice/json ─────────────────────────────────────────────────
app.post("/api/invoice/json", (req, res) => {
  try {
    const data = mergeWithDefaults(req.body);
    const t    = calcTotals(data);
    res.json({
      invoice: {
        ...data,
        items: t.lineItems.map(({ _amount, ...rest }) => ({
          ...rest, amount: _amount, amountFormatted: fmtINR(_amount)
        })),
      },
      totals: {
        taxable:         t.taxable,
        serviceCharge:   t.svc,
        subtotal:        t.sub,
        cgst:            t.cgst,
        sgst:            t.sgst,
        grandTotal:      t.grand,
        grandTotalWords: toWords(t.grand) + " only",
        formatted: {
          taxable:       fmtINR(t.taxable),
          serviceCharge: fmtINR(t.svc),
          subtotal:      fmtINR(t.sub),
          cgst:          fmtINR(t.cgst),
          sgst:          fmtINR(t.sgst),
          grandTotal:    fmtINR(t.grand),
        }
      }
    });
  } catch (err) {
    res.status(400).json({ error: err.message });
  }
});

// ── Fallback to SPA ────────────────────────────────────────────────────────
app.use((req, res) => res.sendFile(path.join(__dirname, "public", "index.html")));

if (require.main === module) {
  const PORT = process.env.PORT || 3000;
  app.listen(PORT, () => {
    console.log(`\n✅  Invoice API + UI  →  http://localhost:${PORT}`);
    console.log(`    API docs           →  http://localhost:${PORT}/api\n`);
  });
}

module.exports = app;
