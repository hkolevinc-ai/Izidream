# Izidream → Temu scraper

## Files
- `scraper.py` – discovers and extracts Izidream product pages.
- `template.xlsx` – supplied Temu template.
- `.github/workflows/scrape.yml` – manual GitHub Actions workflow.
- `TEMU_IZIDREAM_UPLOAD.xlsx` – generated Temu file.
- `izidream_raw_export.csv` – audit/export file.

## First test in GitHub
1. Upload the entire folder contents to a GitHub repository.
2. Open **Actions → Scrape Izidream → Run workflow**.
3. Leave `limit` as `20` for the first test.
4. Download the `izidream-results` artifact.
5. After checking the result, run again with `limit = 0` for all discovered products.

## Applied rules
- Each Izidream product page remains a separate Temu product. Similar sizes/colors are not merged.
- Current price → `Base Price - EUR`.
- Higher previous price → `Recommended retail price - EUR`; otherwise `N/A`.
- Brand/Trademark → `Izidream`.
- Original gallery images only (`/resources/...`).
- Weight comes from the product page and is converted from kg to g.
- Package L/W/H are estimated by product category.
- Shipping Template, Fulfillment Channel, Manufacturer, EU responsible person and Product Identification remain blank.
- Country of origin defaults to Bulgaria, consistent with the supplied product description/example.

## Local run
```bash
pip install -r requirements.txt
python scraper.py --limit 20
python scraper.py --limit 0
```
