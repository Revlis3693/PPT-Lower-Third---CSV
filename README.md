# Lower Third Builder (PowerPoint task pane add-in)

An MVP PowerPoint Office Add-in that builds lower-third slides from a **CSV** by mapping CSV columns to **text boxes** on a template slide.

## Prerequisites

- **Node.js** 18+ (Node 20+ recommended)
- **npm** (ships with Node)
- **Microsoft PowerPoint** (desktop) for sideloading on Windows or Mac

## Setup

In a terminal:

```bash
cd "/Users/marcsilver/Desktop/LowerThirdBuilder"
npm install
```

Install trusted dev certificates for `https://localhost:3000` (one-time):

```bash
npm run dev-certs:install
```

## Run locally (dev)

Start the local dev server:

```bash
npm run dev
```

This serves the task pane at `https://localhost:3000/taskpane.html` and outputs a built manifest at `dist/manifest.xml` (the source manifest is `manifest/manifest.xml`).

## Production deployment (portable — no `npm run dev`)

1. Choose an **HTTPS** URL where you will host the add-in (e.g. Netlify, Azure Static Web Apps, S3 + CloudFront). You need a **valid** certificate (not the dev localhost cert).

2. Set your production origin in **`office-addin.config.cjs`** (`productionOrigin` — no trailing slash), **or** set it only for one build:
   ```bash
   OFFICE_ADDIN_ORIGIN=https://your-site.example.com npm run build
   ```

3. Upload everything in **`dist/`** to your host so these paths work, for example:
   - `https://your-site.example.com/taskpane.html`
   - `https://your-site.example.com/taskpane.js`
   - `https://your-site.example.com/manifest.xml`

4. **Icons**: The manifest references `/assets/icon-16.png` (and 32, 64, 80). Add those files under `dist/assets/` on your host, or update the manifest URLs to real PNGs before building.

5. Sideload **`dist/manifest.xml`** (or the hosted copy of that manifest) in PowerPoint — **not** the dev `manifest/manifest.xml` if you want a production URL.

## Sideload into PowerPoint

### PowerPoint on Windows

- Open PowerPoint.
- Go to **Insert** → **Add-ins** → **My Add-ins**.
- Choose **Add a Custom Add-in** → **Add from file...**
- Select: `LowerThirdBuilder/manifest/manifest.xml`
- The add-in will appear on the ribbon. Click **Lower Third Builder** to open the task pane.

### PowerPoint on Mac

- Open PowerPoint.
- Go to **Insert** → **Add-ins** → **My Add-ins**.
- Choose the option to **Add from file...**
- Select: `LowerThirdBuilder/manifest/manifest.xml`

### PowerPoint for the web

Sideloading behavior varies by tenant and admin policy. If your tenant allows it:

- Open PowerPoint for the web.
- Go to **Insert** → **Add-ins** → **Manage My Add-ins**.
- Upload the manifest: `manifest/manifest.xml`

## How to use

### 1) Create a template slide (in PowerPoint)

- Create a new slide that represents your lower-third layout.
- Add **text boxes** for fields like **Name**, **Title**, **Company**.
- (Optional but recommended) Give text boxes meaningful names in the **Selection Pane**, because duplicates are matched by **shape name** during generation.

### 2) Load CSV data

- In the task pane, under **1) Data Source**, upload a CSV file.
- Example CSV is included at `sample-data/sample.csv`.

### 3) Map CSV columns to template text boxes

- In the thumbnail pane, select **only your template slide** (one slide).
- Under **2) Column Mapping**, click **Refresh shapes** to list text boxes / placeholders on that slide.
- Choose a **CSV column**, choose a **shape** from the list, then click **Add mapping**.
- Repeat for each field (Name, Title, Company, etc.).

### 4) Preview a row on the template slide

- Under **3) Preview**, choose a 0-based row index.
- Click **Preview on current slide**.
- This updates mapped text boxes on the selected template slide (it’s a live edit).

### 5) Generate one slide per CSV row

- Select your template slide (single slide selection).
- Under **4) Generate**, click **Generate slides from template**.
- The add-in:
  - exports the template slide as a 1-slide presentation (base64)
  - inserts that slide once per CSV row **after the template slide**
  - updates mapped text boxes on each inserted slide

## Known limitations (v1)

- **Shape types listed**: mapping lists **TextBox**, **Placeholder**, **GeometricShape**, and **Callout** shapes. Other types (e.g. some SmartArt) are skipped.
- **Text boxes only**: only shapes with text are supported.
- **Slide duplication**: Office.js does not expose a native “duplicate slide” API. This add-in duplicates by **exporting the selected slide** and then **inserting it from base64**.
- **Shape identity across duplicates**: shape IDs can change on inserted slides, so slide generation relies on **shape names** when available. For best results, name your text boxes in PowerPoint’s **Selection Pane**.
- **No persistence**: CSV data + mappings live in memory only (reset on refresh).
- **No styling/layout automation**: no auto-resize, background bar adjustments, animations, or image placeholders.

## Project structure

- `manifest/manifest.xml`: Office Add-in manifest (PowerPoint task pane)
- `src/taskpane/`: React task pane UI
  - `components/App.tsx`: UI + state for CSV, mappings, preview, generate
  - `services/csv.ts`: `parseCsv(file)`
  - `services/pptService.ts`: Office.js PowerPoint service layer:
    - `listMappableShapesOnTemplateSlide` (walks `slide.shapes`; avoids `getSelectedShapes`, which can freeze on Mac)
    - `applyRowToMappingsOnSlide`
    - `duplicateTemplateSlideAndPopulate`
- `sample-data/sample.csv`: example CSV
- `webpack.config.js`: dev server + manifest URL rewrite

## Notes on the mapping approach

- A mapping stores:
  - template slide id
  - shape id
  - shape name (if available)
  - CSV column name
- **Preview** tries to find shapes by **id first**, then by **name**.
- **Generate** inserts copies of the template slide; since copied shapes have new IDs, it updates shapes primarily by **shape name**.

