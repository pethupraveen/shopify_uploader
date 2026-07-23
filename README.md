# Shopify Bulk Product Uploader

Reads your Excel product sheet + per-SKU image folders, groups rows into products by **Model Name**, and creates/updates them on your Shopify store via the Admin API — as drafts, ready for you to review before publishing.

## Folder layout you need

```
products.xlsx
images/
  IPHN-17-AIR-TPU-CVR-BLK/     <- folder name = SKU
    1.jpeg
    2.jpeg
    3.jpeg
  IPHN-17-AIR-TPU-CVR-BLU/
    1.jpeg
    2.jpeg
```

Excel columns needed (any order, case-insensitive):
`SKU | EAN | MRP | SSP | Model Name | Material | Colour`

All rows sharing the same **Model Name** become **one Shopify product**, with each row as a **variant** on that product (e.g. different Colour/Material). See `sample_products.xlsx` and `sample_images/` for a working example.

## What gets created, per product

- **Title**: `{Model Name} Wraps (Back Cover Case)`
- **Description**: fixed HTML template (same copy for every product), with `{Model Name}` and `{MRP}` filled in
- **Status**: `draft` — new products are never published live automatically. Review and activate them manually in Shopify admin when ready.
- **Type**: `Mobile Back Covers`
- **Variant options**: **Material** and **Colour** are always both set as options, even if every row shares the same Material
- **Per variant**: price = `SSP`, compare-at price = `MRP` (only if MRP > SSP), SKU, barcode = `EAN`
- **Images**: every image in a SKU's folder is uploaded and linked to that specific variant. The **first image** in the folder (`1.jpeg`, sorted numerically) becomes that variant's main/featured image. Each image's **alt text** is set to `"{Material} - {Colour}"`.
- **Metafields** (namespace `custom`, unless your store uses different ones — see note below):
  - `sub_heading` = "Ultra Sleek Premium Case"
  - `mobile_model` = `{Model Name}`
  - `product_category` = "Mobile Back Covers"
  - `gender` = "unisex"
- **SEO** (via the standard `global` metafields):
  - Page title: `{Model Name} Back Covers by Sprig - India Online`
  - Meta description: `{Model Name} back covers by Sprig with sleek design, shock protection, and perfect fit. Stylish and durable cases for everyday use.`

All of the fixed text above (description template, metafield values, SEO templates, vendor, product type) lives in constants near the top of `upload_products.py` — edit them directly if wording needs to change.

## One-time setup

1. **Get a Shopify Admin API token**
   - In your Shopify admin: **Settings → Apps and sales channels → Develop apps → Create an app**
   - Configure Admin API scopes: at minimum `write_products`, `read_products`
   - Click **Install app**, then copy the **Admin API access token** (starts with `shpat_...`)
   - ⚠️ This token is shown only once — save it somewhere safe.

2. **Install Python dependencies**
   ```bash
   pip install -r requirements.txt
   ```

3. **Set your credentials**
   - Copy `.env.example` to `.env`
   - Fill in your real values:
     ```
     SHOPIFY_STORE=your-store-name        # just the name — NOT the full .myshopify.com URL
     SHOPIFY_ACCESS_TOKEN=shpat_xxxxxxxxxxxxxxxxxxxxxxxx
     SHOPIFY_API_VERSION=2024-10
     SHOPIFY_VENDOR=Default Vendor
     SHOPIFY_PRODUCT_TYPE=Mobile Back Covers
     ```
   - `.env` is git-ignored so your token won't get committed.

## Running it

**Always dry-run first** — prints what would be created, without touching Shopify:
```bash
python upload_products.py --excel products.xlsx --images ./images --dry-run
```

Then run for real:
```bash
python upload_products.py --excel products.xlsx --images ./images
```

If a product with the same title already exists and you want to update it instead of skipping:
```bash
python upload_products.py --excel products.xlsx --images ./images --update-existing
```

## VS Code setup

The project includes `.vscode/launch.json`, `settings.json`, and `extensions.json`.

1. Open the folder in VS Code, install the recommended extensions (Python, Pylance, debugpy) when prompted
2. Create the virtual environment: `python3 -m venv .venv`
3. Select it as the interpreter (`Ctrl+Shift+P` → *Python: Select Interpreter*)
4. `pip install -r requirements.txt`
5. Set up `.env` as above
6. Run/debug via **F5** → pick a configuration, or use the integrated terminal directly

## Notes / things worth knowing

- Shopify's REST API rate limit is roughly 2 requests/second — the script paces itself automatically.
- The duplicate-product check (by title) scans your first 250 products per run. If your store has more than 250 products, tell me and pagination can be added.
- **`--update-existing` limitation**: it updates title/description/metafields and refreshes price/compare-at/barcode for variants that already exist on the matched product. It does **not** add a brand-new variant (e.g. a new colour added later) to an already-existing product — that needs a separate, more involved flow. It also intentionally does **not** touch `status`, so it won't un-publish a product you've already made active in Shopify admin.
- If your store's actual Shopify metafield definitions use a different namespace/key than `custom.sub_heading`, `custom.mobile_model`, `custom.product_category`, `custom.gender`, the values will still get saved but won't populate those specific admin fields. Check **Settings → Custom data → Products** in Shopify admin and let me know the real namespace/keys if they differ.
