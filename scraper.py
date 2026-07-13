from __future__ import annotations

import argparse
import ast
import csv
import html
import json
import logging
import re
import time
from dataclasses import dataclass, asdict
from pathlib import Path
from typing import Any, Iterable
from urllib.parse import urljoin, urlparse

import requests
from bs4 import BeautifulSoup
from openpyxl import load_workbook

BASE_URL = "https://izidream.bg/"
USER_AGENT = "Mozilla/5.0 (compatible; IzidreamTemuExporter/1.0; +https://izidream.bg/)"

# Temu category IDs available in the supplied template.
CATEGORY_RULES = [
    (11915, "Home & Kitchen / Bedding / Bed Pillows & Positioners / Bed Pillows", ["възглавниц", "pillow"]),
    (12044, "Home & Kitchen / Home Décor Products / Decorative Pillows / Throw Pillow Covers", ["декоративна калъф", "калъфка за декоратив"]),
    (12042, "Home & Kitchen / Home Décor Products / Decorative Pillows / Throw Pillows", ["декоративни възглав", "декоративна възглав"]),
    (12011, "Home & Kitchen / Bedding / Duvets & Down Comforters", ["олекотена завив", "зимни завив", "летни завив", "завивка"]),
    (11998, "Home & Kitchen / Bedding / Duvet Covers & Sets / Duvet Cover Sets", ["спален комплект", "спално бельо", "комплект спално"]),
    (11997, "Home & Kitchen / Bedding / Duvet Covers & Sets / Duvet Covers", ["плик за олекотена", "пликове за олекотени"]),
    (11896, "Home & Kitchen / Bedding / Sheets & Pillowcases / Fitted Sheets", ["чаршаф с ластик"]),
    (11897, "Home & Kitchen / Bedding / Sheets & Pillowcases / Flat Sheets", ["долен чаршаф", "чаршафи без ластик", "чаршаф без ластик"]),
    (11894, "Home & Kitchen / Bedding / Sheets & Pillowcases / Pillowcases", ["калъфки за възглав", "калъфка за възглав"]),
    (12009, "Home & Kitchen / Bedding / Bedspreads, Coverlets & Sets / Bedspreads & Coverlets", ["шалте", "кувертюр", "покривало за легло"]),
    (53772, "Home & Kitchen / Home Décor Products / Slipcovers / Sofa Slipcovers / Sofa Throw Covers", ["покривало за диван", "шалте за диван"]),
    (10498, "Home & Kitchen / Kitchen & Dining / Kitchen & Table Linens / Tablecloths", ["покривка за маса", "покривки за маса", "тишлайфер"]),
    (10492, "Home & Kitchen / Kitchen & Dining / Kitchen & Table Linens / Dish Cloths & Dish Towels", ["кухненска кърпа", "кухненски кърпи"]),
    (11812, "Home & Kitchen / Bath / Towels / Bath Towels", ["хавлиена кърпа", "кърпа за баня", "хавлия"]),
    (11810, "Home & Kitchen / Bath / Towels / Towel Sets", ["комплект кърпи", "комплект хавли"]),
    (11899, "Home & Kitchen / Bedding / Blankets & Throws / Bed Blankets", ["одеяло", "поларено одеяло", "памучно одеяло"]),
    (12005, "Home & Kitchen / Bedding / Mattress Protectors & Encasements / Mattress Protectors", ["протектор за матрак", "протектори за матраци"]),
    (11895, "Home & Kitchen / Bedding / Sheets & Pillowcases / Pillow Protectors", ["протектор за възглав"]),
    (12906, "Home & Kitchen / Cleaning Supplies / Household Cleaners / Cleaning Tools / Cleaning Cloths", ["почистваща кърпа", "микрофибърна кърпа"]),
]

PACKAGE_DIMENSIONS = {
    11915: (50.0, 35.0, 18.0), 12042: (45.0, 45.0, 15.0), 12044: (30.0, 25.0, 4.0),
    12011: (55.0, 40.0, 18.0), 11998: (38.0, 30.0, 10.0), 11997: (35.0, 27.0, 6.0),
    11896: (35.0, 27.0, 7.0), 11897: (35.0, 27.0, 7.0), 11894: (28.0, 22.0, 4.0),
    12009: (48.0, 38.0, 15.0), 53772: (45.0, 35.0, 12.0), 10498: (32.0, 25.0, 5.0),
    10492: (25.0, 20.0, 4.0), 11812: (35.0, 28.0, 8.0), 11810: (38.0, 30.0, 12.0),
    11899: (45.0, 35.0, 14.0), 12005: (38.0, 30.0, 10.0), 11895: (28.0, 22.0, 4.0),
    12906: (25.0, 20.0, 4.0),
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
    country_origin: str = "Bulgaria"


def clean_text(value: Any) -> str:
    if value is None:
        return ""
    soup = BeautifulSoup(html.unescape(str(value)), "html.parser")
    return re.sub(r"\s+", " ", soup.get_text(" ", strip=True)).strip()


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
                product_urls.add(target)

    for candidate in candidates:
        parse_map(candidate)
    return product_urls


def is_product_url(url: str) -> bool:
    p = urlparse(url)
    if p.netloc and "izidream.bg" not in p.netloc:
        return False
    path = p.path.rstrip("/")
    blocked = ("/blog", "/contacts", "/za-nas", "/user", "/cart", "/wishlist", "/search", "/promocii", "/podaruchni-vaucheri")
    if any(path.startswith(x) for x in blocked):
        return False
    return bool(re.search(r"-\d+$", path))


def discover_by_crawl(s: requests.Session, max_pages: int = 1500) -> set[str]:
    queue = [urljoin(BASE_URL, "produkti")]
    seen: set[str] = set()
    products: set[str] = set()
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
            target = urljoin(url, a.get("href"))
            if is_product_url(target):
                products.add(target.split("#")[0])
            elif "izidream.bg" in urlparse(target).netloc and any(x in target for x in ("/produkti", "/p", "vazglavnici", "zaviv", "spalno", "charsh", "pokriv", "shalt", "havl", "odeyal", "protektor")):
                if target not in seen and target not in queue:
                    queue.append(target.split("#")[0])
        time.sleep(0.15)
    return products


def extract_event_json(page: str) -> dict[str, Any]:
    marker = "data: JSON.parse('"
    start = page.find(marker)
    if start < 0:
        return {}
    start += len(marker)
    end = page.find("')", start)
    if end < 0:
        return {}
    raw = page[start:end]
    try:
        # The argument is a JavaScript single-quoted string containing JSON.
        # ast.literal_eval safely resolves escaped quotes and \uXXXX sequences
        # without corrupting native Cyrillic characters.
        decoded = ast.literal_eval("'" + raw.replace("'", "\\'") + "'")
        return json.loads(decoded)
    except Exception:
        return {}


def find_product_object(data: dict[str, Any]) -> dict[str, Any]:
    candidates = []
    if isinstance(data.get("object_data"), dict):
        candidates.append(data["object_data"])
    for key in ("themarketer", "object"):
        if isinstance(data.get(key), dict):
            candidates.append(data[key])
    for obj in candidates:
        if obj.get("prod_id") or obj.get("prod_code"):
            return obj
    return {}


def choose_category(name: str, source_cat: str, description: str) -> tuple[int, str]:
    hay = f"{name} {source_cat} {description}".lower()
    for category_id, category_name, needles in CATEGORY_RULES:
        if any(n in hay for n in needles):
            return category_id, category_name
    return DEFAULT_CATEGORY


def parse_product(s: requests.Session, url: str) -> Product | None:
    page = get(s, url).text
    soup = BeautifulSoup(page, "html.parser")

    ld_product = {}
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
    product_id = str(obj.get("prod_id") or re.search(r"-(\d+)$", urlparse(url).path.rstrip("/")).group(1))
    sku = str(obj.get("prod_code") or ld_product.get("sku") or product_id)
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
    description = clean_text(description_html)
    if not description:
        description = clean_text(ld_product.get("description"))
    bullets = [x.strip() for x in re.split(r"(?<=[.!?])\s+|\s*\|\s*", description) if len(x.strip()) >= 15][:6]

    images = []
    for el in soup.select(".gallery__thumb[data-zoom], .product-details__gallery-wrapper [data-zoom]"):
        src = el.get("data-zoom")
        if src:
            full = urljoin(BASE_URL, src)
            if full not in images:
                images.append(full)
    if not images and ld_product.get("image"):
        values = ld_product["image"] if isinstance(ld_product["image"], list) else [ld_product["image"]]
        images = [urljoin(BASE_URL, x) for x in values]

    price = float(obj.get("price") or offer.get("price") or 0)
    old_price_raw = obj.get("oldprice") or obj.get("base_price")
    old_price = float(old_price_raw) if old_price_raw and float(old_price_raw) > price else None
    quantity = int(float(obj.get("prod_count") or (1 if "InStock" in str(offer.get("availability")) else 0)))
    weight_kg = float(obj.get("prod_weight") or 0)
    if not weight_kg:
        for prop in ld_product.get("additionalProperty", []) or []:
            if "тегло" in str(prop.get("name", "")).lower():
                weight_kg = float(str(prop.get("value", "0")).replace(",", "."))
                break
    weight_g = max(1.0, round(weight_kg * 1000, 1))

    category_id, category_name = choose_category(name, category_source, description)
    length_cm, width_cm, height_cm = PACKAGE_DIMENSIONS.get(category_id, DEFAULT_DIMS)

    return Product(
        url=url, product_id=product_id, sku=sku, name=name, brand=brand,
        category_source=category_source, category_id=category_id, category_name=category_name,
        description=description[:2000], bullets=bullets, images=images[:50], quantity=max(0, quantity),
        price=round(price, 2), old_price=round(old_price, 2) if old_price else None,
        weight_g=weight_g, length_cm=length_cm, width_cm=width_cm, height_cm=height_cm,
    )


def header_map(ws) -> dict[str, list[int]]:
    result: dict[str, list[int]] = {}
    for col in range(1, ws.max_column + 1):
        header = clean_text(ws.cell(1, col).value)
        internal = clean_text(ws.cell(3, col).value)
        for key in {header.lower(), internal.lower()}:
            if key:
                result.setdefault(key, []).append(col)
    return result


def first_col(mapping: dict[str, list[int]], *keys: str) -> int | None:
    for key in keys:
        cols = mapping.get(key.lower())
        if cols:
            return cols[0]
    return None


def all_cols(mapping: dict[str, list[int]], *keys: str) -> list[int]:
    out = []
    for key in keys:
        out.extend(mapping.get(key.lower(), []))
    return sorted(set(out))


def set_value(ws, row: int, mapping: dict[str, list[int]], value: Any, *keys: str) -> None:
    col = first_col(mapping, *keys)
    if col:
        ws.cell(row, col).value = value


def fill_template(template: Path, output: Path, products: list[Product]) -> None:
    book = load_workbook(template)
    ws = book["Template"]
    mapping = header_map(ws)
    start_row = 4

    for offset, p in enumerate(products):
        row = start_row + offset
        set_value(ws, row, mapping, p.category_id, "Category", "t_1_category")
        set_value(ws, row, mapping, p.category_name, "Category Name", "t_1_category name")
        set_value(ws, row, mapping, p.name, "Product Name", "t_1_product name")
        set_value(ws, row, mapping, p.sku, "Contribution Goods", "t_1_contribution goods")
        set_value(ws, row, mapping, p.sku, "Contribution SKU", "t_1_contribution sku")
        set_value(ws, row, mapping, "Add", "Update or Add", "t_1_update or add")
        set_value(ws, row, mapping, p.brand, "Brand", "t_1_brand")
        set_value(ws, row, mapping, p.brand, "Trademark", "t_1_trademark")
        set_value(ws, row, mapping, p.description, "Product Description", "t_2_product description")

        bullet_cols = all_cols(mapping, "Bullet Point", "t_2_bullet point")
        for i, col in enumerate(bullet_cols[:6]):
            ws.cell(row, col).value = p.bullets[i] if i < len(p.bullets) else None
        detail_cols = all_cols(mapping, "Detail Images URL", "t_2_detail images url")
        for i, col in enumerate(detail_cols):
            ws.cell(row, col).value = p.images[i] if i < len(p.images) else None

        # No artificial variants: every Izidream product page remains a separate product/SKU.
        set_value(ws, row, mapping, p.images[0] if p.images else "", "SKU Images URL", "t_5_sku images url")
        set_value(ws, row, mapping, p.quantity, "Quantity", "t_5_quantity")
        set_value(ws, row, mapping, p.price, "Base Price - EUR", "t_5_base price - eur")
        if p.old_price:
            set_value(ws, row, mapping, p.old_price, "Recommended retail price - EUR", "t_5_recommended retail price - eur")
        else:
            set_value(ws, row, mapping, "N/A", "Not available for List price", "t_5_not available for list price")
        set_value(ws, row, mapping, p.weight_g, "Weight - g", "t_5_weight - g")
        set_value(ws, row, mapping, p.length_cm, "Length - cm", "t_5_length - cm")
        set_value(ws, row, mapping, p.width_cm, "Width - cm", "t_5_width - cm")
        set_value(ws, row, mapping, p.height_cm, "Height - cm", "t_5_height - cm")
        set_value(ws, row, mapping, "1 Day", "Handling Time", "t_6_handling time")
        set_value(ws, row, mapping, "GEN STANDARD", "Item Tax Code", "t_6_item tax code")
        set_value(ws, row, mapping, p.country_origin, "Country/Region of Origin", "t_8_country/region of origin")
        # Shipping template, fulfillment channel, manufacturer, EU responsible person and product identification intentionally blank.

    book.save(output)


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
                logging.info("[%d/%d] %s", index, len(urls), product.name)
        except Exception as exc:
            logging.exception("Failed product %s: %s", url, exc)
        time.sleep(max(0, args.delay))

    fill_template(Path(args.template), Path(args.output), products)
    export_raw(Path(args.raw), products)
    logging.info("Done: %s products -> %s", len(products), args.output)


if __name__ == "__main__":
    main()
