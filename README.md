# AVA Practice Operating System — Apps Script API + Vercel frontend

Two pieces, no server code of your own to host:

1. **Code.gs** — deployed inside Apps Script as a Web App. It now returns
   JSON (the same shape your old `getData()` produced), instead of HTML.
2. **index.html** — deployed on Vercel as a plain static page. It calls your
   Apps Script Web App URL with `fetch()` to get the data.

## Step 1 — Deploy Code.gs as a Web App

1. Open the Apps Script project bound to your Google Sheet (Extensions →
   Apps Script), or open script.google.com if it's a standalone project.
2. Replace the contents of `Code.gs` with the file in this folder.
3. Click **Deploy → New deployment**.
4. Next to "Select type", click the gear icon → choose **Web app**.
5. Set:
   - **Execute as:** Me
   - **Who has access:** Anyone
6. Click **Deploy**, authorize when prompted.
7. Copy the **Web app URL** — it looks like:
   ```
   https://script.google.com/macros/s/AKfycb.../exec
   ```

You can test it immediately by opening that URL in a browser — it should
show raw JSON like `{"depts":[...], "synced":"..."}`.

## Step 2 — Paste the URL into index.html

Open `index.html` in this folder and find this near the top of the
`<script>` tag:

```js
var APPS_SCRIPT_URL = 'PASTE_YOUR_APPS_SCRIPT_WEB_APP_URL_HERE';
```

Replace the placeholder with the URL you copied in Step 1.

## Step 3 — Deploy index.html on Vercel

Any of these work, since it's a single static HTML file:

**Option A — Vercel CLI**
```bash
npm install -g vercel
vercel        # from this folder, follow the prompts
vercel --prod
```

**Option B — Drag and drop**
Go to https://vercel.com/new, and drag this folder in (or just `index.html`).

**Option C — GitHub**
Push this folder to a GitHub repo, then "Import Project" in the Vercel
dashboard and point it at the repo. No build command or framework needed —
Vercel will just serve `index.html` as a static site.

## Updating the sheet logic later

If you ever edit `Code.gs` again, the live `/exec` URL will **not**
automatically pick up the change. You need to:
Deploy → Manage deployments → pencil (edit) icon on the Web app deployment →
Version: **New version** → Deploy.

## Notes

- **Caching**: the API still caches sheet reads for 120 seconds
  (`CacheService`), same as before. The "Sync now" button sends `?refresh=1`
  to force a fresh read.
- **Column layout**: `COL_FLOW` (O), `COL_FMS` (P), `COL_CHECKLIST` (S),
  `COL_FORM` (T) constants at the top of `Code.gs` are unchanged from your
  original file — edit them there if your sheet's columns move.
- **Access control**: "Who has access: Anyone" means anyone with the `/exec`
  URL can read this JSON (no Google login required). If that's a concern for
  a firm-internal system, you can instead set access to "Anyone with a
  Google account", but then the frontend fetch will need the visitor to be
  signed into Google and CORS/auth becomes more involved. For an internal
  tool behind a private Vercel URL, "Anyone" + an unguessable, unshared
  `/exec` URL is the usual practical tradeoff — just don't publish the URL
  publicly.
