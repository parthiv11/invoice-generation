"use strict";

const DEFAULTS = {
  supplierName:    "KISHAN ENTERPRISE",
  supplierGSTIN:   "24EXJPR5313F1ZG",
  supplierPAN:     "EXJPR5313F",
  supplierAddress: "Rameshwar Nagar, Godhra Road, Halol, Panchmahal, Gujarat - 389350",
  receiverName:    "GR AGRO INDUSTRIES",
  receiverGSTIN:   "24HAYPP3387L1ZK",
  period:          "01/02/2026 to 28/02/2026",
  placeOfSupply:   "GUJARAT (24)",
  invoiceDate:     "05-03-2026",
  invoiceNo:       "KE-38",
  billVariant:     "modern",
  svcPct:          5,
  cgstPct:         9,
  sgstPct:         9,
  items: [
    { desc: "Loading /Unloading", rate: 400, persons: 245, days: 1 },
    { desc: "Loading /Unloading", rate: 0, persons: 0, days: 0 },
    { desc: "Skill",              rate: 0, persons: 0, days: 0 },
    { desc: "Provident Fund@13%", rate: 0, persons: 0, days: 0 },
    { desc: "ESIC@3.25%",         rate: 0, persons: 0, days: 0 },
    { desc: "Bonus",              rate: 0, persons: 0, days: 0 },
  ],
  payFavour:   "KISHAN ENTERPRISE",
  bankBranch:  "HDFC BANK",
  accountNo:   "50200119219082",
  ifscCode:    "HDFC0000954",
  comments:    "KISHAN ENTERPRISE",
  footerText:  "Rameshwar Nagar, Godhra Road, Halol | E-mail: makwanarajesh822@gmail.com",
};

function mergeWithDefaults(body = {}) {
  const merged = { ...DEFAULTS, ...body };
  if (body.items) merged.items = body.items;
  return merged;
}

function calcTotals(data) {
  const svcPct  = (data.svcPct  ?? 5)  / 100;
  const cgstPct = (data.cgstPct ?? 9)  / 100;
  const sgstPct = (data.sgstPct ?? 9)  / 100;

  let taxable = 0;
  const lineItems = (data.items || []).map(item => {
    const rate    = +item.rate    || 0;
    const persons = +item.persons || 0;
    const days    = +item.days    || (rate && persons ? 1 : 0);
    const amount  = rate && persons ? rate * persons * (days || 1) : 0;
    taxable += amount;
    return { ...item, _amount: amount };
  });

  const svc   = +(taxable * svcPct).toFixed(2);
  const sub   = +(taxable + svc).toFixed(2);
  const cgst  = +(sub * cgstPct).toFixed(2);
  const sgst  = +(sub * sgstPct).toFixed(2);
  const grand = +(sub + cgst + sgst).toFixed(2);
  return { taxable, svc, sub, cgst, sgst, grand, lineItems };
}

function fmtINR(n, d = 2) {
  return Number(n).toLocaleString("en-IN", { minimumFractionDigits: d, maximumFractionDigits: d });
}

function toWords(n) {
  n = Math.round(n); if (!n) return "Zero";
  const o = ["","One","Two","Three","Four","Five","Six","Seven","Eight","Nine","Ten",
    "Eleven","Twelve","Thirteen","Fourteen","Fifteen","Sixteen","Seventeen","Eighteen","Nineteen"];
  const t = ["","","Twenty","Thirty","Forty","Fifty","Sixty","Seventy","Eighty","Ninety"];
  const b = x => x < 20 ? o[x] : t[~~(x/10)] + (x%10 ? " " + o[x%10] : "");
  const c = x => x >= 100 ? o[~~(x/100)] + " Hundred" + (x%100 ? " " + b(x%100) : "") : b(x);
  let r = "";
  if (n >= 10000000) { r += c(~~(n/10000000)) + " Crore ";  n %= 10000000; }
  if (n >= 100000)   { r += c(~~(n/100000))   + " Lakh ";   n %= 100000;   }
  if (n >= 1000)     { r += c(~~(n/1000))     + " Thousand "; n %= 1000;   }
  if (n > 0)         { r += c(n); }
  return r.trim();
}

module.exports = { DEFAULTS, mergeWithDefaults, calcTotals, fmtINR, toWords };
