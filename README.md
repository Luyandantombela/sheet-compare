# Sheet Compare — Excel Add-in

Compare two tables from any sheet(s) in the same workbook and report what's **missing**, **matching**, and **different** between them.

---

## File Structure

```
sheet-compare/
├── manifest.xml        ← Office add-in manifest (sideload this)
├── taskpane.html       ← Task pane UI
├── taskpane.css        ← Styles
├── taskpane.js         ← All comparison logic (client-side, no backend)
├── commands.html       ← Required stub for the manifest FunctionFile
├── assets/
│   ├── icon-16.png     ← 16×16 icon (add your own, see note below)
│   ├── icon-32.png     ← 32×32 icon
│   └── icon-80.png     ← 80×80 icon
└── README.md
```

> **Icons:** The manifest references icon files in `/assets/`. Add your own PNG icons, or use any placeholder 16×16, 32×32, and 80×80 PNG files. Icon files are not required to sideload and test locally — Excel will fall back gracefully.

---

## How to Run Locally (Sideloading)

### Option 1 — `http-server` (simplest)

1. Install [Node.js](https://nodejs.org) if you don't have it.
2. Install `http-server` globally:
   ```bash
   npm install -g http-server
   ```
3. From the `sheet-compare/` directory run:
   ```bash
   http-server . -p 3000 --cors -S -C cert.pem -K key.pem
   ```
   The `-S` flag enables HTTPS, which Office requires. You need a local self-signed certificate (see below).
4. Sideload `manifest.xml` in Excel (see "Sideloading" section below).

### Option 2 — Office Add-in Development with `office-addin-dev-certs`

1. Install the Yeoman tooling (one-time setup):
   ```bash
   npm install -g yo generator-office office-addin-dev-certs
   ```
2. Generate a trusted local HTTPS certificate:
   ```bash
   npx office-addin-dev-certs install
   ```
   This installs a self-signed cert trusted by your OS so Excel doesn't block the add-in.
3. Copy the generated `cert.pem` and `key.pem` into the `sheet-compare/` folder (or note their path).
4. Serve the add-in:
   ```bash
   http-server . -p 3000 --cors -S -C ~/.office-addin-dev-certs/localhost.crt -K ~/.office-addin-dev-certs/localhost.key
   ```
   (Cert path may vary by OS — check the output of the install step.)
5. Visit `https://localhost:3000/taskpane.html` in your browser and accept the certificate warning (one time).

### Option 3 — GitHub Pages (no local server needed)

1. Push this folder to a GitHub repository.
2. Enable **GitHub Pages** on the repo (Settings → Pages → Branch: `main`, folder: `/` or `/docs`).
3. Wait for the Pages URL (e.g. `https://yourusername.github.io/sheet-compare/`).
4. Find and replace all occurrences of `https://localhost:3000` in `manifest.xml` with your Pages URL.
5. Sideload the updated `manifest.xml`.

> GitHub Pages serves over HTTPS automatically — no certificates needed.

---

## Sideloading the Add-in in Excel

### Excel on Windows
1. Open Excel and create or open a workbook.
2. Go to **Insert → Add-ins → My Add-ins**.
3. Click **Upload My Add-in** (bottom-left link).
4. Browse to and select `manifest.xml`.
5. Click **Upload**. The "Sheet Compare" button will appear on the **Home** ribbon.

### Excel on Mac
1. Copy `manifest.xml` to the add-in manifest folder:
   ```bash
   cp manifest.xml ~/Library/Containers/com.microsoft.Excel/Data/Documents/wef/
   ```
2. Open Excel → **Insert → Add-ins → My Add-ins** → you should see **Sheet Compare**.

### Excel Online (Office 365)
1. Go to **Insert → Add-ins → Upload My Add-in**.
2. Select `manifest.xml`.

> The add-in is only visible in the workbook it was sideloaded in. For development this is fine.

---

## How to Use

1. **Open the task pane** — click the **Open Sheet Compare** button on the Home ribbon (or Insert → My Add-ins if you re-open the workbook).

2. **Select Table A** — click the button, then select a range in your workbook (including the header row), and click back in the task pane. The add-in reads whatever is currently selected.

3. **Select Table B** — same process; can be on a different sheet.

4. **Pick key columns** — choose one or more columns that together uniquely identify a row (composite key supported via Ctrl+click / Cmd+click in the list).

5. **Choose comparison modes:**
   - **Missing** — rows in A but not B, and rows in B but not A
   - **Matching** — rows whose key exists in both tables
   - **Different** — rows with the same key but differing values

6. **For "Different" mode** — optionally select which non-key columns to compare (leave blank to compare all of them).

7. **Options:**
   - *Ignore case & trim whitespace* — `" John@x.com "` and `"john@x.com"` are treated as equal
   - *Numeric tolerance* — floating-point rounding differences within the tolerance are not flagged

8. **Run Comparison** — a new sheet named **"Comparison Report"** is created with a summary block and labeled sections for each mode.

9. **Highlight in Source Sheets** — color-codes the original Table A and B ranges:
   - 🔴 Red = missing row
   - 🟡 Yellow = row with differences

10. **Export as CSV** — downloads the full report as a `.csv` file.

---

## Error Handling

| Situation | Message shown |
|---|---|
| Range has no headers | "Header row contains blank cells" |
| Table A = Table B (same range) | "Table A and Table B are the same range" |
| Key column has blank values | "Table A/B has rows with blank key values" |
| Duplicate keys found | Warning modal — choose to proceed or cancel |
| No common columns between tables | Warning in the key column selector |

---

## Publishing to AppSource (after testing)

1. Update the `<ProviderName>` in `manifest.xml` to your company name.
2. Generate a new unique `<Id>` (any GUID generator works, e.g. [guidgenerator.com](https://www.guidgenerator.com)).
3. Host the files on a public HTTPS server (GitHub Pages, Azure Static Web Apps, etc.).
4. Replace all `https://localhost:3000` URLs in `manifest.xml` with your hosted URL.
5. Submit `manifest.xml` to [Microsoft Partner Center](https://partner.microsoft.com/en-us/dashboard/office/overview).

---

## Extending the Add-in

The comparison logic in `taskpane.js` is structured for easy extension:

- **New comparison mode** — add a checkbox in `taskpane.html`, a key in the `modes` object in `runComparison()`, and a new call to `writeSection()` in `writeReportSheet()`.
- **Custom normalization** — edit the `normalizeVal()` function.
- **Different output format** — replace or supplement `writeReportSheet()`.

---

## License

MIT — use freely, modify as needed.
