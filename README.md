# AI Excel Add-in

Excel 365 AI assistant — task pane chat + custom functions `=AI()` via 9router (self-hosted AI gateway).

## Fitur

- **💬 Task Pane Chat** — Chat dengan AI langsung di Excel sidebar
- **📝 Write to Cell** — Klik tombol untuk tulis response ke cell aktif
- **⚡ Custom Functions** — `=AI("prompt")`, `=AI.EXPLAIN()`, `=AI.TABLE()`
- **🔗 Terhubung ke 9router** — Self-hosted AI gateway via Cloudflare tunnel
- **⚙️ Config Panel** — Ganti API key, model, temperature tanpa edit code

## Struktur

```
excel-ai-addin/
├── manifest.xml              # Manifest Office Add-in (sideload file)
├── server.js                 # HTTPS dev server (Node.js)
├── package.json
├── cert/
│   ├── key.pem               # Self-signed TLS key
│   └── cert.pem              # Self-signed TLS cert
├── assets/
│   ├── icon-32.png
│   └── icon-80.png
└── src/
    ├── taskpane/
    │   ├── taskpane.html     # Chat UI
    │   ├── taskpane.css      # Material Design 3 styling
    │   └── taskpane.js       # Chat logic + 9router API
    └── functions/
        ├── functions.html    # Custom functions page (shared runtime)
        ├── functions.js      # =AI(), =AI.EXPLAIN(), =AI.TABLE()
        └── functions.json    # Custom functions metadata
```

## Cara Pakai

### 1. Start Server

```bash
cd excel-ai-addin
node server.js
# HTTPS server running on https://localhost:3000
```

### 2. Sideload ke Excel Desktop (Windows)

1. Copy `manifest.xml` ke folder lokal (misal `C:\Users\brita\Desktop\`)
2. Buka Excel → **Insert** tab → **Add-ins** → **My Add-ins**
3. Klik **Upload My Add-in** → Browse ke `manifest.xml`
4. Add-in muncul di ribbon → klik untuk buka task pane

### 3. Sideload ke Excel for Web

1. Butuh HTTPS URL publik. Bisa via Cloudflare Tunnel:

```bash
# Install cloudflared, lalu tunnel ke localhost:3000
cloudflared tunnel --url https://localhost:3000
```

2. Dapat URL `https://xxxx.trycloudflare.com`
3. Edit `manifest.xml` — ganti `https://localhost:3000` ke URL tunnel
4. Buka Excel Online → **Insert** → **Add-ins** → **Upload My Add-in**
5. Pilih manifest dari URL atau upload file

### 4. Konfigurasi API Key

1. Buka task pane → klik ⚙️ Settings
2. Isi **API Key** (dari 9router admin)
3. Pilih **Model** (default: DeepSeek V4 Flash)
4. **Save** — siap dipakai

## 9router API

- Endpoint: `https://9router.britamax.my.id/v1/chat/completions`
- Format: OpenAI-compatible
- Streaming: ✅ SSE (server-sent events)
- Models: 100+ (OpenRouter, Gemini, Claude, GPT, locally-hosted)

## Custom Functions (Excel)

| Function | Description |
|----------|-------------|
| `=AI("question")` | Tanya AI apa pun |
| `=AI.EXPLAIN(A1)` | Jelaskan nilai/rumus di cell |
| `=AI.TABLE(A1:C10, "analyze")` | Analisis range data |

> **Note:** Custom functions jalan di **shared runtime** dengan task pane — jadi setting API key cukup sekali.

## Development

```bash
# Regenerate self-signed cert (1 tahun)
npm run cert

# Start dev server
npm start
```

## Tech Stack

- Office.js API 1.12
- Excel Custom Functions Runtime 1.1
- Shared Runtime (task pane + custom functions satu proses)
- Google Material Design 3 UI
- 9router OpenAI-compatible API
