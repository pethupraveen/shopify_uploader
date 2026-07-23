# Shopify Bulk Product Uploader

Reads your Excel product sheet + per-SKU image folders and creates the products on your Shopify store.

## Folder layout you need

```
products.xlsx
images/
  ABC123/          <- folder name = SKU
    1.jpeg
    2.jpeg
    3.jpeg
  XYZ789/
    1.jpeg
    2.jpeg
```

Excel columns needed (any order, case-insensitive):
`SKU | EAN | MRP | SSP | Model Name | Material | Colour`

See `sample_products.xlsx` and `sample_images/` for a working example.

## VS Code setup

1. **Open the project folder** in VS Code (`File → Open Folder…`, pick this `shopify_uploader` folder).

2. **Install the recommended extensions** — VS Code will prompt you automatically (from `.vscode/extensions.json`), or install manually:
   - Python (`ms-python.python`)
   - Pylance (`ms-python.vscode-pylance`)
   - Python Debugger (`ms-python.debugpy`)

3. **Create a virtual environment** (Terminal → New Terminal, then):
   ```bash
   python3 -m venv .venv
   source .venv/bin/activate        # Windows: .venv\Scripts\activate
   pip install -r requirements.txt
   ```
   VS Code should auto-detect `.venv` as the interpreter (bottom-right status bar). If not: `Ctrl/Cmd+Shift+P → Python: Select Interpreter → .venv`.

4. **Set up your credentials**:
   ```bash
   cp .env.example .env
   ```
   Then open `.env` and fill in your real `SHOPIFY_STORE` and `SHOPIFY_ACCESS_TOKEN`. This file is gitignored so it won't get committed.

5. **Run it** — two ways:
   - **Debug panel** (left sidebar, the play-with-bug icon): pick "Upload Products (dry run)" from the dropdown at top and press F5. Breakpoints work normally.
   - **Terminal**, same as any Python script:
     ```bash
     python upload_products.py --excel sample_products.xlsx --images sample_images --dry-run
     ```

   To point at your real files instead of the samples, either edit the `args` in `.vscode/launch.json`, or just run from the terminal with your own `--excel` / `--images` paths.

## One-time setup

1. **Get a Shopify Admin API token**
   - In your Shopify admin: **Settings → Apps and sales channels → Develop apps → Create an app**
   - Configure Admin API scopes: at minimum `write_products`, `read_products`
   - Click **Install app**, then copy the **Admin API access token** (starts with `shpat_...`)
   - ⚠️ This token is shown only once — save it somewhere safe.

2. **Install Python dependencies**
   ```bash
   pip install pandas openpyxl requests
   ```

3. **Set your credentials as environment variables**
   ```bash
   export SHOPIFY_STORE="your-store-name"        # from your-store-name.myshopify.com
   export SHOPIFY_ACCESS_TOKEN="shpat_xxxxxxxxxxxxxxxxxxxxxxxx"
   ```
   (Or edit the CONFIG block at the top of `upload_products.py` directly — not recommended if you'll share the file.)

## Running it

**Always dry-run first** — this prints what would be created without touching Shopify:
```bash
python upload_products.py --excel products.xlsx --images ./images --dry-run
```

Then run for real:
```bash
python upload_products.py --excel products.xlsx --images ./images
```

If a SKU already exists as a product and you want to update it instead of skipping:
```bash
python upload_products.py --excel products.xlsx --images ./images --update-existing
```

## What it does per row

- **Title**: built from Model Name + Material + Colour
- **Description**: a simple HTML bullet list from your fields (Model, Material, Colour, EAN, SKU) — easy to swap for AI-generated copy later if you want richer descriptions
- **Price**: set to `SSP`
- **Compare-at price** (the strikethrough price): set to `MRP`, but only if MRP > SSP
- **SKU / Barcode**: SKU and EAN mapped to the variant
- **Images**: all `.jpeg/.jpg/.png/.webp` files in the SKU's folder, uploaded in numeric order (1, 2, 3...)
- **Duplicate SKU check**: skips (or updates, with the flag) products that already exist with that SKU

## Notes / things worth knowing

- Shopify's REST API rate limit is roughly 2 requests/second — the script paces itself automatically.
- The duplicate-SKU check currently scans your first 250 products per run (Shopify REST pagination). If your store has more than 250 products, tell me and I'll add cursor-based pagination through your full catalog.
- No variants (like size/color options) are set up here — each row becomes one simple product with one variant. If you actually want color/size as Shopify **variants** on the same product (rather than separate products), that changes the data model — let me know and I'll adjust it.
- Descriptions are template-based, not AI-generated, per your setup. If you'd like richer AI-written descriptions later, I can add an OpenAI/Claude API call to generate `body_html` from the same fields.
