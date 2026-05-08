# 🧾 RIDHI ENTERPRISE — Invoice API + Builder

Send JSON → Get PDF or Excel invoice. Hosted free on Render.com.

## Run locally
```bash
npm install
node server.js
# → http://localhost:3000
```

## API Endpoints
| Method | Path | Returns |
|--------|------|---------|
| GET | /api | API docs |
| POST | /api/invoice/pdf | PDF file |
| POST | /api/invoice/xlsx | Excel file |
| POST | /api/invoice/json | Calculated JSON |

## Quick example
```bash
curl -X POST http://localhost:3000/api/invoice/pdf \
  -H "Content-Type: application/json" \
  -d '{"invoiceNo":"RE-39","items":[{"desc":"Labour","rate":400,"persons":100,"days":1}]}' \
  --output invoice.pdf
```

## Deploy FREE on Render.com
1. Push to GitHub
2. render.com → New Web Service → connect repo  
3. Build: `npm install` | Start: `node server.js`
4. Select Free plan → Deploy!

Your live URL: `https://your-app.onrender.com`
