# Local Excel Bridge (Laptop Mini Server)

Use this when your app is hosted on Vercel but you want Excel to open on your laptop automatically.

## 1) Start the bridge on your laptop

From `form137-react`:

```powershell
npm run bridge:excel
```

Default endpoint:

- `http://127.0.0.1:8787`
- Health check: `GET /health`
- Open endpoint: `POST /open-excel`

## 2) Configure your web app

Set these in Vercel (Project -> Settings -> Environment Variables):

- `VITE_EXCEL_BRIDGE_URL`
  Example: `http://127.0.0.1:8787`
- `VITE_EXCEL_BRIDGE_KEY` (optional but recommended)

If the bridge is reachable, clicking `Download Excel` in web mode will:

1. send student data to the bridge,
2. generate `.xlsx` on your laptop,
3. open it in desktop Excel.

If bridge call fails, the app falls back to browser download.

## 3) Optional security and CORS on bridge

You can run the bridge with environment variables:

- `BRIDGE_API_KEY` (requires `X-Bridge-Key` header)
- `BRIDGE_ALLOW_ORIGIN` (default `*`)
- `BRIDGE_HOST` (default `127.0.0.1`)
- `BRIDGE_PORT` (default `8787`)
- `BRIDGE_TEMPLATE_PATH` (default `public/Form 137-SHS-BLANK.xlsx`)
- `BRIDGE_OUTPUT_DIR` (default temp folder)

Example:

```powershell
$env:BRIDGE_API_KEY='your-secret'
$env:BRIDGE_ALLOW_ORIGIN='https://your-vercel-app.vercel.app'
npm run bridge:excel
```
