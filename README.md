# Izidream → Temu scraper

## Files
- `scraper.py` – discovers and extracts Izidream product pages.
- `template.xlsx` – supplied Temu template.
- `.github/workflows/scrape.yml` – manual GitHub Actions workflow.
- `TEMU_IZIDREAM_UPLOAD.xlsx` – generated Temu file.
- `izidream_raw_export.csv` – audit/export file.

## First test in GitHub
1. Replace the old repository files with the contents of this package.
2. Open **Actions → Scrape Izidream → Run workflow**.
3. Leave `limit` as `20` for the first test.
4. Download the `izidream-results` artifact.
5. Check rows beginning at **row 5** in the `Template` sheet.
6. After checking the result, run again with `limit = 0` for all discovered products.

## Applied rules
- Each Izidream product page remains a separate Temu product. Similar sizes/colors are not merged.
- Category ID and Category Name are selected from the supplied Temu category list.
- Product name, Contribution Goods and Contribution SKU are populated for every row.
- Current price → `Base Price - EUR`.
- Higher previous price → `List Price - EUR`; otherwise `N/A`.
- Product URL → `Reference Link`.
- Brand/Trademark → `Izidream` or the brand found on the product page.
- Original gallery images only (`/resources/...`).
- Weight comes from the product page and is converted from kg to g.
- Package L/W/H are estimated by product category.
- Country of origin is read from the description when present; otherwise Bulgaria.
- Shipping Template, Fulfillment Channel, Manufacturer, EU responsible person and Product Identification remain blank.

## Built-in safety check
The script validates the generated workbook before the GitHub job finishes. It fails instead of uploading a bad result when Category, Product Name, SKU, Quantity, Price, Weight or package dimensions are missing, or if Temu's technical row 4 is overwritten.

## Local run
```bash
pip install -r requirements.txt
python scraper.py --limit 20
python scraper.py --limit 0
```
