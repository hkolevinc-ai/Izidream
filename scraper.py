from __future__ import annotations

import argparse
import ast
import csv
import html
import json
import logging
import re
import time
from dataclasses import asdict, dataclass
from pathlib import Path
from typing import Any
from urllib.parse import urljoin, urlparse

import requests
from bs4 import BeautifulSoup
from openpyxl import load_workbook

BASE_URL = "https://izidream.bg/"
USER_AGENT = "Mozilla/5.0 (compatible; IzidreamTemuExporter/1.1; +https://izidream.bg/)"

# More-specific rules must be placed before broad rules.
CATEGORY_RULES = [
    (12044, "Home & Kitchen / Home Décor Products / Decorative Pillows / Throw Pillow Covers", ["декоративна калъф", "калъфка за декоратив"]),
    (12042, "Home & Kitchen / Home Décor Products / Decorative Pillows / Throw Pillows", ["декоративни възглав", "декоративна възглав"]),
    (11895, "Home & Kitchen / Bedding / Sheets & Pillowcases / Pillow Protectors", ["протектор за възглав"]),
    (11894, "Home & Kitchen / Bedding / Sheets & Pillowcases / Pillowcases", ["калъфки за възглав", "калъфка за възглав"]),
    (11915, "Home & Kitchen / Bedding / Bed Pillows & Positioners / Bed Pillows", ["възглавница за сън", "възглавници за спане", "анатомична възглав", "мемори възглав", "възглавниц"]),
    (12005, "Home & Kitchen / Bedding / Mattress Protectors & Encasements / Mattress Protectors", ["протектор за матрак", "протектори за матраци"]),
    (12002, "Home & Kitchen / Bedding / Mattress Pads & Toppers / Mattress Pads", ["подматрачна", "матрачен топер", "топ матрак"]),
    (11998, "Home & Kitchen / Bedding / Duvet Covers & Sets / Duvet Cover Sets", ["спален комплект", "спално бельо", "комплект спално"]),
    (11997, "Home & Kitchen / Bedding / Duvet Covers & Sets / Duvet Covers", ["плик за олекотена", "пликове за олекотени"]),
    (11896, "Home & Kitchen / Bedding / Sheets & Pillowcases / Fitted Sheets", ["чаршаф с ластик"]),
    (11897, "Home & Kitchen / Bedding / Sheets & Pillowcases / Flat Sheets", ["долен чаршаф", "чаршафи без ластик", "чаршаф без ластик"]),
    (12010, "Home & Kitchen / Bedding / Bedspreads, Coverlets & Sets / Bedspread & Coverlet Sets", ["комплект шалте", "комплект кувертюра"]),
    (12009, "Home & Kitchen / Bedding / Bedspreads, Coverlets & Sets / Bedspreads & Coverlets", ["шалте", "кувертюр", "покривало за легло"]),
    (53772, "Home & Kitchen / Home Décor Products / Slipcovers / Sofa Slipcovers / Sofa Throw Covers", ["покривало за диван", "шалте за диван"]),
    (12011, "Home & Kitchen / Bedding / Duvets & Down Comforters", ["олекотена завив", "зимни завив", "летни завив", "завивка"]),
    (11810, "Home & Kitchen / Bath / Towels / Towel Sets", ["комплект кърпи", "комплект хавли"]),
    (11812, "Home & Kitchen / Bath / Towels / Bath Towels", ["хавлиена кърпа", "кърпа за баня", "хавлия"]),
    (10492, "Home & Kitchen / Kitchen & Dining / Kitchen & Table Linens / Dish Cloths & Dish Towels", ["кухненска кърпа", "кухненски кърпи"]),
    (10498, "Home & Kitchen / Kitchen & Dining / Kitchen & Table Linens / Tablecloths", ["покривка за маса", "покривки за маса", "тишлайфер"]),
    (12906, "Home & Kitchen / Cleaning Supplies / Household Cleaners / Cleaning Tools / Cleaning Cloths", ["почистваща кърпа", "микрофибърна кърпа", "кърпа за почистване"]),
    (11901, "Home & Kitchen / Bedding / Blankets & Throws / Throws", ["одеяло за диван", "плед", "throw"]),
    (11899, "Home & Kitchen / Bedding / Blankets & Throws / Bed Blankets", ["одеяло", "поларено одеяло", "памучно одеяло"]),
    (26695, "Baby Products / Nursery / Bedding / Toddler Bedding / Bedding Sets", ["бебешки спален комплект", "комплект за кошара"]),
    (26696, "Baby Products / Nursery / Bedding / Toddler Bedding / Sheets, Pillowcases & Sets / Bed Sheets", ["бебешки чаршаф", "чаршаф за кошара"]),
    (11836, "Home & Kitchen / Bath / Kids' Bath / Kids' Bath Towels", ["детска хавлия", "детска кърпа"]),
    (29133, "Clothing, Shoes & Jewelry / Women / Clothing / Lingerie, Sleep & Lounge / Sleepwear / Robes", ["халат", "хавлиен халат"]),
]

PACKAGE_DIMENSIONS = {
    11915: (50.0, 35.0, 18.0), 12042: (45.0, 45.0, 15.0), 12044: (30.0, 25.0, 4.0),
    12011: (55.0, 40.0, 18.0), 11998: (38.0, 30.0, 10.0), 11997: (35.0, 27.0, 6.0),
    11896: (35.0, 27.0, 7.0), 11897: (35.0, 27.0, 7.0), 11894: (28.0, 22.0, 4.0),
    12009: (48.0, 38.0, 15.0), 12010: (50.0, 40.0, 16.0), 53772: (45.0, 35.0, 12.0),
    10498: (32.0, 25.0, 5.0), 10492: (25.0, 20.0, 4.0), 11812: (35.0, 28.0, 8.0),
    11810: (38.0, 30.0, 12.0), 11899: (45.0, 35.0, 14.0), 11901: (42.0, 32.0, 12.0),
    12005: (38.0, 30.0, 10.0), 12002: (55.0, 42.0, 20.0), 11895: (28.0, 22.0, 4.0),
    12906: (25.0, 20.0, 4.0), 26695: (38.0, 30.0, 10.0), 26696: (30.0, 24.0, 6.0),
    11836: (32.0, 25.0, 7.0), 29133: (38.0, 30.0, 8.0),
}
DEFAULT_CATEGORY = (11998, "Home & Kitchen / Bedding / Duvet Covers & Sets / Duvet Cover Sets")
DEFAULT_DIMS = (40.0, 30.0, 10.0)


@dataclass
class Product:
    url: str
    product_id: str
    sku: str
    name: str
    brand: str
    category_source: str
    category_id: int
    category_name: str
    description: str
    bullets: list[str]
    images: list[str]
    quantity: int
    price: float
    old_price: float | None
    weight_g: float
    length_cm: float
    width_cm: float
    height_cm: float
    country_origin: str


def clean_text(value: Any) -> str:
    if value is None:
        return ""
    soup = BeautifulSoup(html.unescape(str(value)), "html.parser")
    return re.sub(r"\s+", " ", soup.get_text(" ", strip=True)).strip()


def normalize_key(value: Any) -> str:
    return re.sub(r"\s+", " ", clean_text(value)).strip().casefold()


def session() -> requests.Session:
    s = requests.Session()
    s.headers.update({"User-Agent": USER_AGENT, "Accept-Language": "bg-BG,bg;q=0.9,en;q=0.7"})
    return s


def get(s: requests.Session, url: str, retries: int = 3) -> requests.Response:
    last = None
    for attempt in range(retries):
        try:
            r = s.get(url, timeout=35)
            r.raise_for_status()
            return r
        except requests.RequestException as exc:
            last = exc
            time.sleep(2 ** attempt)
    raise RuntimeError(f"Failed to download {url}: {last}")


def is_product_url(url: str) -> bool:
    p = urlparse(url)
    if p.netloc and "izidream.bg" not in p.netloc:
        return False
    path = p.path.rstrip("/")
    blocked = ("/blog", "/contacts", "/za-nas", "/user", "/cart", "/wishlist", "/search", "/promocii", "/podaruchni-vaucheri")
    if any(path.startswith(x) for x in blocked):
        return False
    return bool(re.search(r"-\d+$", path))


def discover_from_sitemaps(s: requests.Session) -> set[str]:
    candidates = [urljoin(BASE_URL, x) for x in ("sitemap.xml", "sitemap_index.xml", "sitemap-products.xml")]
    product_urls: set[str] = set()
    seen_maps: set[str] = set()

    def parse_map(url: str) -> None:
        if url in seen_maps:
            return
        seen_maps.add(url)
        try:
            text = get(s, url).text
        except Exception:
            return
        soup = BeautifulSoup(text, "xml")
        for loc in soup.find_all("loc"):
            target = loc.get_text(strip=True)
            if target.endswith(".xml"):
                parse_map(target)
            elif is_product_url(target):
                product_urls.add(target.split("#")[0])

    for candidate in candidates:
        parse_map(candidate)
    return product_urls


def discover_by_crawl(s: requests.Session, max_pages: int = 1800) -> set[str]:
    queue = [urljoin(BASE_URL, "produkti")]
    seen: set[str] = set()
    products: set[str] = set()
    allowed_hints = ("/produkti", "/p", "vazglavnici", "zaviv", "spalno", "charsh", "pokriv", "shalt", "havl", "odeyal", "protektor", "dete", "bebe", "tekstil")
    while queue and len(seen) < max_pages:
        url = queue.pop(0)
        if url in seen:
            continue
        seen.add(url)
        try:
            soup = BeautifulSoup(get(s, url).text, "html.parser")
        except Exception as exc:
            logging.warning("Skip discovery page %s: %s", url, exc)
            continue
        for a in soup.select("a[href]"):
            target = urljoin(url, a.get("href")).split("#")[0]
            if is_product_url(target):
                products.add(target)
            elif "izidream.bg" in urlparse(target).netloc and any(x in target for x in allowed_hints):
                if target not in seen and target not in queue:
                    queue.append(target)
        time.sleep(0.15)
    return products


def extract_event_json(page: str) -> dict[str, Any]:
    """Return the richest JSON payload embedded in data: JSON.parse('...')."""
    candidates: list[dict[str, Any]] = []
    pattern = re.compile(r"data:\s*JSON\.parse\('(.*?)'\)", re.DOTALL)
    for match in pattern.finditer(page):
        raw = match.group(1)
        try:
            decoded = ast.literal_eval("'" + raw.replace("'", "\\'") + "'")
            value = json.loads(decoded)
            if isinstance(value, dict):
                candidates.append(value)
        except Exception:
            continue
    if not candidates:
        return {}

    def score(data: dict[str, Any]) -> int:
        text = json.dumps(data, ensure_ascii=False)
        return sum(token in text for token in ("prod_code", "prod_name", "prod_body_short", "prod_weight", "oldprice", "prod_count"))

    return max(candidates, key=score)


def find_product_object(data: dict[str, Any]) -> dict[str, Any]:
    candidates = []
    if isinstance(data.get("object_data"), dict):
        candidates.append(data["object_data"])
    for key in ("themarketer", "object"):
        if isinstance(data.get(key), dict):
            candidates.append(data[key])
    candidates.sort(key=lambda x: sum(bool(x.get(k)) for k in ("prod_code", "prod_name", "prod_body_short", "prod_weight", "price")), reverse=True)
    for obj in candidates:
        if obj.get("prod_id") or obj.get("prod_code"):
            return obj
    return {}


def choose_category(name: str, source_cat: str, description: str) -> tuple[int, str]:
    hay = f"{name} {source_cat} {description}".casefold()
    for category_id, category_name, needles in CATEGORY_RULES:
        if any(n.casefold() in hay for n in needles):
            return category_id, category_name
    logging.warning("No category rule matched: %s | %s. Using fallback %s", name, source_cat, DEFAULT_CATEGORY[0])
    return DEFAULT_CATEGORY


def extract_origin(description: str) -> str:
    match = re.search(r"страна\s+на\s+произход\s*[:\-]?\s*([А-ЯA-Z][А-Яа-яA-Za-z ]{2,30})", description, re.IGNORECASE)
    if match:
        value = match.group(1).strip().split(" Доставка")[0].strip(" .;,")
        normalized = {"българия": "Bulgaria", "турция": "Turkey", "китай": "China", "пакистан": "Pakistan", "индия": "India"}
        return normalized.get(value.casefold(), value)
    return "Bulgaria"


def parse_product_html(page: str, url: str) -> Product | None:
    soup = BeautifulSoup(page, "html.parser")

    ld_product: dict[str, Any] = {}
    for node in soup.select('script[type="application/ld+json"]'):
        try:
            value = json.loads(node.string or "{}")
        except Exception:
            continue
        if isinstance(value, dict) and value.get("@type") == "Product":
            ld_product = value
            break
    if not ld_product:
        logging.warning("Not a product page: %s", url)
        return None

    event = extract_event_json(page)
    obj = find_product_object(event)
    offer = ld_product.get("offers") if isinstance(ld_product.get("offers"), dict) else {}

    name = clean_text(obj.get("prod_name") or ld_product.get("name"))
    if not name or "ваучер" in name.casefold():
        return None

    id_match = re.search(r"-(\d+)$", urlparse(url).path.rstrip("/"))
    product_id = str(obj.get("prod_id") or (id_match.group(1) if id_match else ""))
    sku = str(obj.get("prod_code") or ld_product.get("sku") or product_id).strip()

    brand_value = ld_product.get("brand", {})
    brand = clean_text(obj.get("brand", {}).get("title") if isinstance(obj.get("brand"), dict) else "")
    if not brand and isinstance(brand_value, dict):
        brand = clean_text(brand_value.get("name"))
    brand = brand or "Izidream"

    category_source = ""
    if isinstance(event.get("category"), dict):
        category_source = clean_text(event["category"].get("hierarchy") or event["category"].get("title"))
    if not category_source and isinstance(obj.get("category"), dict):
        category_source = clean_text(obj["category"].get("hierarchy") or obj["category"].get("title"))
    if not category_source:
        crumbs = [clean_text(x) for x in soup.select(".breadcrumb__link")]
        category_source = " | ".join(crumbs[1:-1])

    description_html = obj.get("prod_body_short") or ""
    desc_soup = BeautifulSoup(html.unescape(str(description_html)), "html.parser")
    paragraph_bullets = [clean_text(node) for node in desc_soup.select("p, li")]
    paragraph_bullets = [x for x in paragraph_bullets if len(x) >= 8]
    description = clean_text(description_html) or clean_text(ld_product.get("description"))
    bullets = paragraph_bullets[:6]
    if not bullets:
        bullets = [x.strip() for x in re.split(r"(?<=[.!?])\s+|\s*\|\s*", description) if len(x.strip()) >= 15][:6]

    images: list[str] = []
    for el in soup.select(".gallery__thumb[data-zoom], .product-details__gallery-wrapper [data-zoom]"):
        src = el.get("data-zoom")
        if src:
            full = urljoin(BASE_URL, src)
            if full not in images:
                images.append(full)
    if not images and ld_product.get("image"):
        values = ld_product["image"] if isinstance(ld_product["image"], list) else [ld_product["image"]]
        images = [urljoin(BASE_URL, str(x)) for x in values]

    price = float(str(obj.get("price") or offer.get("price") or 0).replace(",", "."))
    old_price_raw = obj.get("oldprice") or obj.get("base_price")
    old_price = float(str(old_price_raw).replace(",", ".")) if old_price_raw and float(str(old_price_raw).replace(",", ".")) > price else None
    quantity = int(float(obj.get("prod_count") or (1 if "InStock" in str(offer.get("availability")) else 0)))

    weight_kg = float(str(obj.get("prod_weight") or 0).replace(",", "."))
    if not weight_kg:
        for prop in ld_product.get("additionalProperty", []) or []:
            if "тегло" in str(prop.get("name", "")).casefold():
                weight_kg = float(str(prop.get("value", "0")).replace(",", "."))
                break
    weight_g = max(1.0, round(weight_kg * 1000, 1))

    category_id, category_name = choose_category(name, category_source, description)
    length_cm, width_cm, height_cm = PACKAGE_DIMENSIONS.get(category_id, DEFAULT_DIMS)

    return Product(
        url=url,
        product_id=product_id,
        sku=sku,
        name=name,
        brand=brand,
        category_source=category_source,
        category_id=category_id,
        category_name=category_name,
        description=description[:2000],
        bullets=bullets,
        images=images[:95],
        quantity=max(0, quantity),
        price=round(price, 2),
        old_price=round(old_price, 2) if old_price else None,
        weight_g=weight_g,
        length_cm=length_cm,
        width_cm=width_cm,
        height_cm=height_cm,
        country_origin=extract_origin(description),
    )


def parse_product(s: requests.Session, url: str) -> Product | None:
    return parse_product_html(get(s, url).text, url)


def header_map(ws) -> dict[str, list[int]]:
    """Temu display headers are on row 2; internal field names are on row 4."""
    result: dict[str, list[int]] = {}
    for col in range(1, ws.max_column + 1):
        for value in (ws.cell(2, col).value, ws.cell(4, col).value):
            key = normalize_key(value)
            if key:
                result.setdefault(key, []).append(col)
    return result


def first_col(mapping: dict[str, list[int]], *keys: str) -> int | None:
    # Keys are checked in the supplied order, so exact internal field names should come first.
    for key in keys:
        cols = mapping.get(normalize_key(key))
        if cols:
            return cols[0]
    return None


def all_cols(mapping: dict[str, list[int]], *keys: str) -> list[int]:
    out: list[int] = []
    for key in keys:
        out.extend(mapping.get(normalize_key(key), []))
    return sorted(set(out))


def set_value(ws, row: int, mapping: dict[str, list[int]], value: Any, *keys: str) -> None:
    col = first_col(mapping, *keys)
    if col is None:
        raise KeyError(f"Template column not found for any of: {keys}")
    ws.cell(row, col).value = value


def fill_template(template: Path, output: Path, products: list[Product]) -> None:
    book = load_workbook(template)
    ws = book["Template"]
    mapping = header_map(ws)
    start_row = 5  # Rows 1-4 are Temu headers / reserved technical rows.

    if normalize_key(ws.cell(4, 20).value) != normalize_key("t_2_Product Description"):
        raise RuntimeError("Unexpected Temu template structure: technical row 4 is missing or modified.")

    for offset, p in enumerate(products):
        row = start_row + offset
        set_value(ws, row, mapping, p.category_id, "t_1_Category", "Category")
        set_value(ws, row, mapping, p.category_name, "t_1_Category Name", "Category Name")
        set_value(ws, row, mapping, p.name, "t_1_Product Name", "Product Name")
        set_value(ws, row, mapping, p.sku, "t_1_Contribution Goods", "Contribution Goods")
        set_value(ws, row, mapping, p.sku, "t_1_Contribution SKU", "Contribution SKU")
        set_value(ws, row, mapping, "Add", "t_1_Update or Add", "Update or Add")
        set_value(ws, row, mapping, p.brand, "t_1_Brand", "Brand")
        set_value(ws, row, mapping, p.brand, "t_1_Trademark", "Trademark")
        set_value(ws, row, mapping, p.description, "t_2_Product Description", "Product Description")

        bullet_cols = all_cols(mapping, "t_2_Bullet Point", "Bullet Point")
        for i, col in enumerate(bullet_cols[:6]):
            ws.cell(row, col).value = p.bullets[i] if i < len(p.bullets) else None

        detail_cols = all_cols(mapping, "t_2_Detail Images URL", "Detail Images URL")
        for i, col in enumerate(detail_cols):
            ws.cell(row, col).value = p.images[i] if i < len(p.images) else None

        sku_image_cols = all_cols(mapping, "t_6_SKU Images URL", "SKU Images URL")
        if sku_image_cols:
            ws.cell(row, sku_image_cols[0]).value = p.images[0] if p.images else ""

        set_value(ws, row, mapping, p.quantity, "t_6_Quantity")
        set_value(ws, row, mapping, p.price, "t_6_Base Price - EUR", "Base Price - EUR")
        set_value(ws, row, mapping, p.url, "t_6_Reference Link", "Reference Link")
        if p.old_price:
            set_value(ws, row, mapping, p.old_price, "t_6_List Price - EUR", "List Price - EUR")
        else:
            set_value(ws, row, mapping, "N/A", "t_6_Not available for List price", "Not available for List price")
        set_value(ws, row, mapping, p.weight_g, "t_6_Weight - g", "Weight - g")
        set_value(ws, row, mapping, p.length_cm, "t_6_Length - cm", "Length - cm")
        set_value(ws, row, mapping, p.width_cm, "t_6_Width - cm", "Width - cm")
        set_value(ws, row, mapping, p.height_cm, "t_6_Height - cm", "Height - cm")
        set_value(ws, row, mapping, "1 Day", "t_7_Handling Time", "Handling Time")
        set_value(ws, row, mapping, "GEN STANDARD", "t_7_Item Tax Code", "Item Tax Code")
        set_value(ws, row, mapping, p.country_origin, "t_8_Country/Region of Origin", "Country/Region of Origin")
        # Shipping Template, Fulfillment Channel, Manufacturer, EU responsible person and Product Identification intentionally blank.

    book.save(output)
    validate_output(output, len(products))


def validate_output(path: Path, expected_rows: int) -> None:
    book = load_workbook(path, read_only=True, data_only=False)
    ws = book["Template"]
    mapping = header_map(ws)
    start_row = 5

    if normalize_key(ws.cell(4, 20).value) != normalize_key("t_2_Product Description"):
        raise RuntimeError("Validation failed: Temu technical row 4 was overwritten.")

    required_keys = {
        "Category": ("t_1_Category",),
        "Product Name": ("t_1_Product Name",),
        "Contribution SKU": ("t_1_Contribution SKU",),
        "Quantity": ("t_6_Quantity",),
        "Base Price": ("t_6_Base Price - EUR",),
        "Weight": ("t_6_Weight - g",),
        "Length": ("t_6_Length - cm",),
        "Width": ("t_6_Width - cm",),
        "Height": ("t_6_Height - cm",),
    }
    cols = {label: first_col(mapping, *keys) for label, keys in required_keys.items()}
    missing_columns = [label for label, col in cols.items() if col is None]
    if missing_columns:
        raise RuntimeError(f"Validation failed: template columns not found: {missing_columns}")

    errors: list[str] = []
    for row in range(start_row, start_row + expected_rows):
        for label, col in cols.items():
            value = ws.cell(row, col).value
            if value in (None, ""):
                errors.append(f"row {row}: {label} is blank")
        if len(errors) >= 20:
            break
    if errors:
        raise RuntimeError("Validation failed:\n" + "\n".join(errors))
    logging.info("Output validation passed: %d populated product rows beginning at row %d", expected_rows, start_row)


def export_raw(path: Path, products: list[Product]) -> None:
    fields = list(asdict(products[0]).keys()) if products else ["url"]
    with path.open("w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=fields)
        writer.writeheader()
        for p in products:
            row = asdict(p)
            row["bullets"] = " | ".join(p.bullets)
            row["images"] = " | ".join(p.images)
            writer.writerow(row)


def main() -> None:
    parser = argparse.ArgumentParser(description="Scrape Izidream products and fill the supplied Temu template.")
    parser.add_argument("--template", default="template.xlsx")
    parser.add_argument("--output", default="TEMU_IZIDREAM_UPLOAD.xlsx")
    parser.add_argument("--raw", default="izidream_raw_export.csv")
    parser.add_argument("--limit", type=int, default=20, help="0 means all products")
    parser.add_argument("--urls-file", help="Optional text file containing one product URL per line")
    parser.add_argument("--delay", type=float, default=0.35)
    args = parser.parse_args()

    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
    s = session()
    if args.urls_file:
        urls = {x.strip() for x in Path(args.urls_file).read_text(encoding="utf-8").splitlines() if x.strip()}
    else:
        urls = discover_from_sitemaps(s)
        if not urls:
            logging.info("No usable sitemap found; using category crawl fallback")
            urls = discover_by_crawl(s)
    urls = sorted(urls)
    if args.limit > 0:
        urls = urls[:args.limit]
    logging.info("Product URLs selected: %d", len(urls))

    products: list[Product] = []
    for index, url in enumerate(urls, 1):
        try:
            product = parse_product(s, url)
            if product:
                products.append(product)
                logging.info("[%d/%d] SKU=%s | %s | Category=%s", index, len(urls), product.sku, product.name, product.category_id)
        except Exception as exc:
            logging.exception("Failed product %s: %s", url, exc)
        time.sleep(max(0, args.delay))

    if not products:
        raise RuntimeError("No product data was extracted. Output file was not created.")

    fill_template(Path(args.template), Path(args.output), products)
    export_raw(Path(args.raw), products)
    logging.info("Done: %s products -> %s", len(products), args.output)


if __name__ == "__main__":
    main()
