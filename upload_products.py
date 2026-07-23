#!/usr/bin/env python3
"""
Shopify Bulk Product Uploader
==============================
Reads product data from an Excel file and images from per-SKU folders,
groups rows by Model Name into ONE Shopify product per model (each row
becomes a variant of that product), and creates/updates it via the
Admin API.

------------------------------------------------------------------
EXPECTED INPUTS
------------------------------------------------------------------
1) Excel file with (at least) these columns - names are matched
   case-insensitively and with spaces/underscores ignored, so
   "Model Name", "model_name", "ModelName" all work:

     SKU | EAN | MRP | SSP | Model Name | Material | Colour

   All rows sharing the same "Model Name" become ONE product, with
   each row as a separate variant (e.g. different Colour/Material).

2) An images root folder containing one sub-folder PER SKU, e.g.:

     images/
       IPHN-17-AIR-TPU-CVR-BLK/
         1.jpeg
         2.jpeg
       IPHN-17-AIR-TPU-CVR-BLU/
         1.jpeg

   The sub-folder name must exactly match the SKU value in Excel.
   Each SKU's images are attached to that specific variant.

------------------------------------------------------------------
SETUP
------------------------------------------------------------------
1. Create a custom app in Shopify Admin:
     Settings -> Apps and sales channels -> Develop apps -> Create an app
   Give it these Admin API scopes (minimum):
     write_products, read_products
   Install the app and copy the "Admin API access token" (starts with shpat_).

2. Set environment variables (or use a .env file, see .env.example):
     SHOPIFY_STORE=your-store-name        # from your-store-name.myshopify.com (no .myshopify.com suffix)
     SHOPIFY_ACCESS_TOKEN=shpat_xxxxxxxx

3. Install dependencies:
     pip install -r requirements.txt

------------------------------------------------------------------
USAGE
------------------------------------------------------------------
  # Dry run first (no data is sent to Shopify, just prints what would happen)
  python upload_products.py --excel products.xlsx --images ./images --dry-run

  # Real run
  python upload_products.py --excel products.xlsx --images ./images

  # Update existing products (matched by title) instead of skipping them
  python upload_products.py --excel products.xlsx --images ./images --update-existing

------------------------------------------------------------------
KNOWN LIMITATION
------------------------------------------------------------------
--update-existing currently updates title/description and refreshes
price/compare-at-price/barcode for variants that already exist on the
matched product. It does NOT add a brand-new variant (e.g. a new
colour added later) to an already-existing product - that requires a
separate, more involved API flow. If you need that, let me know and
I'll extend it.
"""

import argparse
import base64
import os
import re
import sys
import time
import logging
from pathlib import Path
from collections import defaultdict

import pandas as pd
import requests

try:
    from dotenv import load_dotenv
    load_dotenv()  # loads variables from a .env file in the project root, if present
except ImportError:
    pass  # dotenv is optional - env vars can still be set manually

# ------------------------------------------------------------------
# CONFIG (env vars override these if set)
# ------------------------------------------------------------------
SHOPIFY_STORE = os.environ.get("SHOPIFY_STORE", "")          # e.g. "my-store"
SHOPIFY_ACCESS_TOKEN = os.environ.get("SHOPIFY_ACCESS_TOKEN", "")
SHOPIFY_API_VERSION = os.environ.get("SHOPIFY_API_VERSION", "2024-10")
DEFAULT_VENDOR = os.environ.get("SHOPIFY_VENDOR", "Default Vendor")
DEFAULT_PRODUCT_TYPE = os.environ.get("SHOPIFY_PRODUCT_TYPE", "Mobile Back Covers")
REQUEST_DELAY_SECONDS = 0.6   # basic rate-limit courtesy (Shopify allows ~2 req/sec on REST)

# --- Product organization / metafields (same for every back cover product) ---
# If your store's metafield definitions use different namespaces/keys, change them here.
METAFIELD_NAMESPACE = "custom"
SUB_HEADING_VALUE = "Ultra Sleek Premium Case"
PRODUCT_CATEGORY_VALUE = "Mobile Back Covers"
GENDER_VALUE = "unisex"

# --- SEO (page title / meta description) ---
SEO_TITLE_TEMPLATE = "{model_name} Back Covers by Sprig - India Online"
SEO_DESCRIPTION_TEMPLATE = ("{model_name} back covers by Sprig with sleek design, shock protection, "
                            "and perfect fit. Stylish and durable cases for everyday use.")

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%H:%M:%S",
)
log = logging.getLogger("shopify_uploader")


# ------------------------------------------------------------------
# Fixed description template (same for every back cover product)
# ------------------------------------------------------------------
DESCRIPTION_TEMPLATE = """<p>Sprig brings you a Premium quality Back Case for Your Apple <strong>{model_name}</strong> Smart Phone.</p>
<p>Choice of Liquid Silicone, Armor Clip, Stand, TPU Matte, Matte Leather, and MagSafe with great build quality, providing splendid grip and shock resistance.</p>
<p>MagSafe magnets are precisely designed and fitted to support seamless wireless charging in all MagSafe Back Case models.</p>
<p>With a stylish look, the phone case is designed to perfectly fit and keep your <strong>{model_name}</strong> protected from scratches and damage.</p>
<p>Perfect for outdoor activities and strong enough to protect it from bumps and accidental falls.</p>
<p>This mobile cover is 100% compatible with standard chargers and headphones used with <strong>{model_name}</strong>.</p>
<p>This phone cover for <strong>{model_name}</strong> has a thickened design to provide complete access to all buttons and enhanced protection.</p>
<p>Raised screen edges help protect the front display.</p>
<p>Accurately designed to leave access to all sensors, ports, speakers, and microphones.</p>
<p>Soft microfiber lining on Liquid Silicone and Liquid Silicone MagSafe models helps prevent scratches on the back of your <strong>{model_name}</strong>.</p>
<p>Sprig Back Covers are available in multiple colors.</p>
<ol>
<li>Kindly note that this back cover is made of environmentally sustainable material as per ROHS standards with a smooth finish and a PC casing inside for rigidity.</li>
<li>For Liquid Silicone/Liquid Silicone MagSafe material, the sides and bottom are intentionally flexible around buttons, speaker cut-outs, and ports.</li>
<li>The case provides excellent grip while absorbing shocks to protect your device.</li>
<li>All Sprig Back Cases support wireless charging except Armor Clip and Stand cases.</li>
<li>Sprig cases can be cleaned using soap and water and retain their color over time (except Armor Clip and Stand).</li>
<li>Sprig does not use recycled material, solid silicone, or synthetic rubber because of associated health concerns.</li>
</ol>
<p><strong>Name of Product:</strong><br>
Sprig Back Cover / Back Case for Apple <strong>{model_name}</strong></p>
<p><strong>Country of Origin:</strong><br>
China</p>
<p><strong>Net Quantity:</strong><br>
1 Unit</p>
<p><strong>Imported, Packed &amp; Marketed by:</strong><br>
Jamsticks India Private Limited,<br>
First &amp; Second Floor,<br>
No.24/1, Astha Lakshmi Layout,<br>
2nd Main,<br>
Puttenahalli Main Road,<br>
JP Nagar Phase 6,<br>
Bangalore - 560078</p>
<p><strong>MRP:</strong><br>
Rs. {mrp} (Inclusive of all taxes)</p>"""


# ------------------------------------------------------------------
# Helpers
# ------------------------------------------------------------------
def normalize_col(name: str) -> str:
    return re.sub(r"[^a-z0-9]", "", str(name).lower())


COLUMN_ALIASES = {
    "sku": "SKU",
    "ean": "EAN",
    "mrp": "MRP",
    "ssp": "SSP",
    "modelname": "Model Name",
    "material": "Material",
    "colour": "Colour",
    "color": "Colour",
}


def load_excel(path: str) -> pd.DataFrame:
    df = pd.read_excel(path, dtype=str)  # read everything as string first, we'll cast numbers later
    rename_map = {}
    for col in df.columns:
        key = normalize_col(col)
        if key in COLUMN_ALIASES:
            rename_map[col] = COLUMN_ALIASES[key]
    df = df.rename(columns=rename_map)

    required = ["SKU", "EAN", "MRP", "SSP", "Model Name", "Material", "Colour"]
    missing = [c for c in required if c not in df.columns]
    if missing:
        log.error(f"Excel file is missing required column(s): {missing}")
        log.error(f"Columns found: {list(df.columns)}")
        sys.exit(1)

    for col in required:
        df[col] = df[col].astype(str).str.strip()

    df = df[df["SKU"] != ""].reset_index(drop=True)
    return df


def find_images_for_sku(images_root: Path, sku: str):
    """Look for a folder named exactly `sku` inside images_root, and collect
    1.jpeg, 2.jpeg, ... (also accepts .jpg/.png, and non-strictly-numbered names)."""
    folder = images_root / sku
    if not folder.is_dir():
        return []

    files = []
    for f in folder.iterdir():
        if f.suffix.lower() in (".jpeg", ".jpg", ".png", ".webp"):
            files.append(f)

    def sort_key(p: Path):
        m = re.match(r"(\d+)", p.stem)
        return (int(m.group(1)) if m else 999999, p.name)

    files.sort(key=sort_key)
    return files


def to_price(value) -> str:
    """Shopify wants price as a decimal string, e.g. '199.00'."""
    try:
        return f"{float(str(value).replace(',', '').strip()):.2f}"
    except (ValueError, TypeError):
        return "0.00"


def format_mrp_display(value) -> str:
    """Human-friendly MRP for the description text, e.g. 1499 (no forced decimals)."""
    try:
        num = float(str(value).replace(",", "").strip())
        if num == int(num):
            return f"{int(num):,}"
        return f"{num:,.2f}"
    except (ValueError, TypeError):
        return str(value)


def build_title(model_name: str) -> str:
    return f"{model_name} Wraps (Back Cover Case)"


def build_description_html(model_name: str, mrp_value) -> str:
    return DESCRIPTION_TEMPLATE.format(model_name=model_name, mrp=format_mrp_display(mrp_value))


def build_seo_title(model_name: str) -> str:
    return SEO_TITLE_TEMPLATE.format(model_name=model_name)


def build_seo_description(model_name: str) -> str:
    return SEO_DESCRIPTION_TEMPLATE.format(model_name=model_name)


def build_metafields(model_name: str) -> list:
    """Product metafields + SEO (page title / meta description) fields.
    SEO fields use Shopify's standard 'global' namespace so they populate the
    'Search engine listing' title/description in Shopify admin."""
    return [
        {"namespace": METAFIELD_NAMESPACE, "key": "sub_heading", "type": "single_line_text_field", "value": SUB_HEADING_VALUE},
        {"namespace": METAFIELD_NAMESPACE, "key": "mobile_model", "type": "single_line_text_field", "value": model_name},
        {"namespace": METAFIELD_NAMESPACE, "key": "product_category", "type": "single_line_text_field", "value": PRODUCT_CATEGORY_VALUE},
        {"namespace": METAFIELD_NAMESPACE, "key": "gender", "type": "single_line_text_field", "value": GENDER_VALUE},
        {"namespace": "global", "key": "title_tag", "type": "single_line_text_field", "value": build_seo_title(model_name)},
        {"namespace": "global", "key": "description_tag", "type": "single_line_text_field", "value": build_seo_description(model_name)},
    ]


# ------------------------------------------------------------------
# Grouping: rows -> one product group per Model Name
# ------------------------------------------------------------------
def group_rows_by_model(df: pd.DataFrame):
    """Returns a list of dicts: {"model_name": ..., "rows": [row, row, ...]}"""
    groups = []
    for model_name, group_df in df.groupby("Model Name", sort=False):
        groups.append({"model_name": model_name, "rows": list(group_df.to_dict("records"))})
    return groups


def determine_variant_options(rows: list) -> list:
    """Always use Material and Colour as the two variant options, as long as
    at least one row in the group has a non-empty value for that field.
    (Previously this only added an option if values differed across rows;
    now both are always exposed as options even if every row shares the
    same Material, e.g. all TPU with different Colours.)"""
    option_candidates = ["Material", "Colour"]
    options = []
    for field in option_candidates:
        values = {r.get(field, "").strip() for r in rows if r.get(field, "").strip()}
        if values:
            options.append(field)
    return options


# ------------------------------------------------------------------
# Shopify API
# ------------------------------------------------------------------
class ShopifyClient:
    def __init__(self, store: str, token: str, api_version: str):
        if not store or not token:
            log.error("SHOPIFY_STORE and SHOPIFY_ACCESS_TOKEN must be set (env vars or CONFIG block).")
            sys.exit(1)
        self.base_url = f"https://{store}.myshopify.com/admin/api/{api_version}"
        self.session = requests.Session()
        self.session.headers.update({
            "X-Shopify-Access-Token": token,
            "Content-Type": "application/json",
        })

    def _request(self, method: str, path: str, **kwargs) -> requests.Response:
        url = f"{self.base_url}{path}"
        resp = self.session.request(method, url, **kwargs)
        for attempt in range(4):
            if resp.status_code != 429:
                return resp
            wait = float(resp.headers.get("Retry-After", 2))
            log.warning(f"Rate limited, waiting {wait}s...")
            time.sleep(wait)
            resp = self.session.request(method, url, **kwargs)
        return resp

    def find_product_by_title(self, title: str):
        """Scans products (first 250) looking for an exact title match."""
        resp = self._request("GET", "/products.json", params={"limit": 250})
        if resp.status_code != 200:
            log.warning(f"Could not list products to check for existing title: {resp.text[:200]}")
            return None
        for product in resp.json().get("products", []):
            if product.get("title") == title:
                return product
        return None

    def create_product(self, payload: dict):
        return self._request("POST", "/products.json", json={"product": payload})

    def update_product(self, product_id: int, payload: dict):
        return self._request("PUT", f"/products/{product_id}.json", json={"product": payload})

    def update_variant(self, variant_id: int, payload: dict):
        return self._request("PUT", f"/variants/{variant_id}.json", json={"variant": payload})

    def update_image(self, product_id: int, image_id: int, payload: dict):
        return self._request("PUT", f"/products/{product_id}/images/{image_id}.json", json={"image": payload})


# ------------------------------------------------------------------
# Payload builders
# ------------------------------------------------------------------
def build_group_payload(model_name: str, rows: list, images_root: Path):
    """Builds the full product payload (title, description, options, variants,
    images) for one Model Name group. Also returns a sku -> [uploaded filenames]
    map so we can re-link images to the right variant after creation."""
    option_fields = determine_variant_options(rows)  # e.g. ["Colour"] or ["Material", "Colour"]

    # Shopify needs each option's full distinct value list, in a stable order
    option_values = {}
    for field in option_fields:
        seen = []
        for r in rows:
            v = r.get(field, "").strip()
            if v and v not in seen:
                seen.append(v)
        option_values[field] = seen

    options_payload = [{"name": field, "values": option_values[field]} for field in option_fields]

    variants_payload = []
    all_images_payload = []
    sku_to_filenames = {}  # sku -> list of unique filenames uploaded for that sku

    first_mrp = rows[0].get("MRP", "0")

    for row in rows:
        sku = row["SKU"]
        price = to_price(row["SSP"])
        compare_at = to_price(row["MRP"])
        compare_at_val = compare_at if float(compare_at) > float(price) else None

        variant = {
            "sku": sku,
            "price": price,
            "barcode": row.get("EAN", ""),
            "inventory_management": "shopify",
            "fulfillment_service": "manual",
        }
        if compare_at_val:
            variant["compare_at_price"] = compare_at_val

        for idx, field in enumerate(option_fields, start=1):
            variant[f"option{idx}"] = row.get(field, "").strip()

        variants_payload.append(variant)

        image_files = find_images_for_sku(images_root, sku)
        if not image_files:
            log.warning(f"[{sku}] No images found in folder '{images_root / sku}'.")

        material = row.get("Material", "").strip()
        colour = row.get("Colour", "").strip()
        alt_text = " - ".join(p for p in [material, colour] if p)

        filenames_for_sku = []
        for f in image_files:
            unique_filename = f"{sku}__{f.name}"  # prefix guarantees no collisions across SKUs (e.g. every folder has "1.jpeg")
            with open(f, "rb") as fh:
                b64 = base64.b64encode(fh.read()).decode("utf-8")
            image_entry = {"attachment": b64, "filename": unique_filename}
            if alt_text:
                image_entry["alt"] = alt_text
            all_images_payload.append(image_entry)
            filenames_for_sku.append(unique_filename)
        sku_to_filenames[sku] = filenames_for_sku

    payload = {
        "title": build_title(model_name),
        "body_html": build_description_html(model_name, first_mrp),
        "vendor": DEFAULT_VENDOR,
        "product_type": DEFAULT_PRODUCT_TYPE,
        "status": "draft",  # products are created unpublished by default - publish manually in Shopify admin when ready
        "variants": variants_payload,
        "images": all_images_payload,
        "metafields": build_metafields(model_name),
    }
    if options_payload:
        payload["options"] = options_payload

    return payload, sku_to_filenames


def link_images_to_variants(client: "ShopifyClient", product_id: int, product_json: dict, sku_to_filenames: dict, sku_to_variant_id: dict):
    """After a product is created, Shopify returns image objects with URLs
    that include the unique filename we uploaded (e.g. .../SKU__1.jpeg).
    We match on the exact filename (not just a substring) to avoid mixing up
    similarly-named SKUs, then set each variant's main/featured image to the
    FIRST image file in that SKU's folder (1.jpeg before 2.jpeg, etc. - same
    numeric order used when the files were read from disk)."""
    images = product_json.get("images", [])

    filename_to_image_id = {}
    for img in images:
        src = img.get("src", "")
        src_basename = src.split("?")[0].rsplit("/", 1)[-1]  # strip query string, keep just the file part
        for filenames in sku_to_filenames.values():
            for fn in filenames:
                if src_basename == fn or src_basename.endswith(fn):
                    filename_to_image_id[fn] = img["id"]

    for sku, filenames in sku_to_filenames.items():
        if not filenames:
            continue
        variant_id = sku_to_variant_id.get(sku)
        if not variant_id:
            continue

        # `filenames` is already in numeric folder order (1.jpeg, 2.jpeg, ...)
        # from find_images_for_sku(), so image_ids preserves that same order.
        image_ids = [filename_to_image_id[fn] for fn in filenames if fn in filename_to_image_id]
        if not image_ids:
            log.warning(f"[{sku}] Uploaded images could not be matched back to this variant - main image not set.")
            continue

        main_image_id = image_ids[0]  # first image in the SKU's folder becomes the variant's featured image
        resp = client.update_variant(variant_id, {"id": variant_id, "image_id": main_image_id})
        if resp.status_code not in (200, 201):
            log.warning(f"[{sku}] Could not set variant main image: {resp.text[:200]}")

        for image_id in image_ids:
            resp = client.update_image(product_id, image_id, {"id": image_id, "variant_ids": [variant_id]})
            if resp.status_code not in (200, 201):
                log.warning(f"[{sku}] Could not link image {image_id} to variant: {resp.text[:200]}")
            time.sleep(REQUEST_DELAY_SECONDS)


# ------------------------------------------------------------------
# Main
# ------------------------------------------------------------------
def main():
    parser = argparse.ArgumentParser(description="Bulk upload products to Shopify from Excel + image folders (grouped by Model Name).")
    parser.add_argument("--excel", required=True, help="Path to the Excel file with product data.")
    parser.add_argument("--images", required=True, help="Path to the root folder containing per-SKU image subfolders.")
    parser.add_argument("--dry-run", action="store_true", help="Don't call Shopify, just show what would happen.")
    parser.add_argument("--update-existing", action="store_true", help="If a product with the same title exists, update it instead of skipping.")
    args = parser.parse_args()

    excel_path = Path(args.excel)
    images_root = Path(args.images)

    if not excel_path.exists():
        log.error(f"Excel file not found: {excel_path}")
        sys.exit(1)
    if not images_root.exists():
        log.error(f"Images folder not found: {images_root}")
        sys.exit(1)

    df = load_excel(str(excel_path))
    groups = group_rows_by_model(df)
    log.info(f"Loaded {len(df)} rows from {excel_path} -> {len(groups)} product(s) after grouping by Model Name")

    client: ShopifyClient | None = None
    if not args.dry_run:
        client = ShopifyClient(SHOPIFY_STORE, SHOPIFY_ACCESS_TOKEN, SHOPIFY_API_VERSION)

    created, updated, skipped, failed = 0, 0, 0, 0

    for group in groups:
        model_name = group["model_name"]
        rows = group["rows"]
        skus = [r["SKU"] for r in rows]

        payload, sku_to_filenames = build_group_payload(model_name, rows, images_root)

        if args.dry_run:
            opt_summary = payload.get("options", [])
            log.info(f"[DRY RUN] Product '{payload['title']}' | SKUs={skus} | "
                      f"Options={[o['name'] for o in opt_summary]} | "
                      f"Variants={len(payload['variants'])} | "
                      f"Images total={len(payload['images'])}")
            continue

        assert client is not None  # guaranteed above: reached only when not args.dry_run

        existing = client.find_product_by_title(payload["title"])
        if existing and not args.update_existing:
            log.info(f"[{model_name}] Product '{payload['title']}' already exists (id={existing['id']}) — skipping. Use --update-existing to update it.")
            skipped += 1
            continue

        if existing and args.update_existing:
            update_payload = {k: v for k, v in payload.items() if k != "status"}  # don't touch status on updates - avoid un-publishing an already-active product
            resp = client.update_product(existing["id"], update_payload)
            if resp.status_code in (200, 201):
                product_json = resp.json()["product"]
                log.info(f"[{model_name}] Updated product id={existing['id']} (SKUs: {skus})")
                sku_to_variant_id = {v.get("sku"): v.get("id") for v in product_json.get("variants", [])}
                link_images_to_variants(client, existing["id"], product_json, sku_to_filenames, sku_to_variant_id)
                updated += 1
            else:
                log.error(f"[{model_name}] Update failed ({resp.status_code}): {resp.text[:300]}")
                failed += 1
        else:
            resp = client.create_product(payload)
            if resp.status_code == 201:
                product_json = resp.json()["product"]
                new_id = product_json["id"]
                log.info(f"[{model_name}] Created product id={new_id} (SKUs: {skus})")
                sku_to_variant_id = {v.get("sku"): v.get("id") for v in product_json.get("variants", [])}
                link_images_to_variants(client, new_id, product_json, sku_to_filenames, sku_to_variant_id)
                created += 1
            else:
                log.error(f"[{model_name}] Create failed ({resp.status_code}): {resp.text[:300]}")
                failed += 1

        time.sleep(REQUEST_DELAY_SECONDS)

    if not args.dry_run:
        log.info(f"Done. Created={created} Updated={updated} Skipped={skipped} Failed={failed}")
    else:
        log.info("Dry run complete. Re-run without --dry-run to actually upload.")


if __name__ == "__main__":
    main()
