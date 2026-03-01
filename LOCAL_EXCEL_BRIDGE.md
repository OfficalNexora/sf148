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
  Example: `https://your-tunnel.ngrok-free.app`
- `VITE_EXCEL_BRIDGE_KEY` (optional but recommended)

If the bridge is reachable, clicking `Print to Excel` in web mode will:

1. send student data to the bridge,
2. generate `.xlsx` on your laptop,
3. auto-print to your default printer (Windows bridge host),
4. open it in desktop Excel.

If bridge call fails, the app falls back to browser download.

## 3) ngrok tunnel (for remote devices)

`127.0.0.1` only works on the same machine.  
If other devices will use your Vercel app and trigger Excel on your laptop, tunnel your bridge with ngrok.

### One-command helper

```powershell
npm run bridge:ngrok
```

Before first run, set ngrok token once:

```powershell
powershell -ExecutionPolicy Bypass -File scripts/setup-ngrok-auth.ps1 `
  -AuthToken "<YOUR_NGROK_AUTHTOKEN>" `
  -NgrokPath "C:\form137application\ngrok-bin\ngrok.exe" `
  -NgrokConfigPath "C:\form137application\form137-react\scripts\ngrok.yml"
```

Optional params:

```powershell
powershell -ExecutionPolicy Bypass -File scripts/start-bridge-ngrok.ps1 `
  -BridgePort 8787 `
  -VercelOrigin "https://your-app.vercel.app" `
  -BridgeKey "your-secret-key" `
  -NgrokPath "C:\path\to\ngrok.exe"
```

The script will:

1. start the local bridge,
2. start ngrok,
3. print the public URL,
4. print the exact `VITE_EXCEL_BRIDGE_URL` and `VITE_EXCEL_BRIDGE_KEY` values to place in Vercel.

## 4) Optional security and CORS on bridge

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
