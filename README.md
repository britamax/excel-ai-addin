# AI Excel Add-in

Excel 365 AI assistant — task pane chat + custom functions `=AI()` with any OpenAI-compatible API.

## Features

- **💬 Task Pane Chat** — Chat with AI directly in Excel sidebar
- **📝 Write to Cell** — Button to write AI response to active cell
- **⚡ Custom Functions** — `=AI("prompt")`, `=AI.EXPLAIN()`, `=AI.TABLE()`
- **⚙️ Bring Your Own Endpoint** — Configure any OpenAI-compatible API endpoint, model, and API key
- **🔗 Shared Runtime** — Task pane and custom functions share config (set once)

## Structure

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
    │   └── taskpane.js       # Chat logic
    └── functions/
        ├── functions.html    # Custom functions page (shared runtime)
        ├── functions.js      # =AI(), =AI.EXPLAIN(), =AI.TABLE()
        └── functions.json    # Custom functions metadata
```

## Setup

### 1. Get the manifest

Download `manifest.xml` from:  
`https://britamax.github.io/excel-ai-addin/manifest.xml`

Add your AI provider's domain(s) to the `<AppDomains>` section in `manifest.xml` before sideloading.

### 2. Sideload to Excel

**Desktop (Windows / Mac):**
1. Excel → **Insert** tab → **Add-ins** → **My Add-ins**
2. Click **Upload My Add-in** → select your `manifest.xml`
3. Add-in appears in ribbon → click to open task pane

**Excel for Web:**
1. Open Excel Online → **Insert** → **Add-ins** → **Upload My Add-in**
2. Upload `manifest.xml` or provide URL

### 3. Configure

1. Open task pane → click **⚙️ Settings**
2. Enter your **Endpoint URL** (e.g. `https://api.openai.com/v1/chat/completions`)
3. Enter your **API Key**
4. Enter **Model name** (optional — leave blank for endpoint default)
5. Adjust **Temperature** if needed
6. **Save** — ready to use

## Custom Functions (Excel)

| Function | Description |
|----------|-------------|
| `=AI("question")` | Ask AI anything |
| `=AI.EXPLAIN(A1)` | Explain value/formula in a cell |
| `=AI.TABLE(A1:C10, "analyze")` | Analyze a data range |

> **Note:** Custom functions run in **shared runtime** with the task pane — configure your endpoint and API key once in the task pane, and `=AI()` functions use the same settings automatically.

## Development

```bash
# Install dependencies
npm install

# Generate self-signed cert (1 year)
npm run cert

# Start HTTPS dev server
npm start
```

## Tech Stack

- Office.js API 1.12
- Excel Custom Functions Runtime 1.1
- Shared Runtime (task pane + custom functions in one process)
- Google Material Design 3 UI
- Any OpenAI-compatible API
