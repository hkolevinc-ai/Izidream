# Izidream → Temu scraper

## Files
- `scraper.py` – discovers and extracts Izidream product pages and fills the supplied Temu workbook.
- `template.xlsx` – the supplied Temu template.
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

## Mandatory red fields
The red cells are driven by Temu conditional formatting and change with the Category ID. The scraper reads the `GoodsLevelMode` and `Dropdown Lists` sheets dynamically and fills the required fields for each selected category, including when applicable:

- Category ID and Category Name
- Product Name, Contribution Goods and Contribution SKU
- cover/material composition and filling material
- Variation Theme, Size Family, Size and Color
- category-specific product dimensions and filling weight
- SKU image, Quantity, Base Price and List Price/N/A
- package Weight, Length, Width and Height
- Individually packed, Total packaging quantity and Packaging unit
- Shipping Template = `Bulgaria`
- Fulfillment Channel = `I will ship this item myself`
- Handling Time, Item Tax Code and Country/Region of Origin

Manufacturer, EU responsible person and Product Identification are not invented because they must match records configured and approved in the seller account.

## Applied scraping rules
- Every Izidream product page remains a separate Temu product. Similar sizes/colors are not merged.
- Category selection prioritizes the main product title, so components mentioned in descriptions do not change the category.
- Bedding sets are not classified as pillowcases merely because a pillowcase is included.
- Towel sets are not classified as robes merely because the source breadcrumb says “Кърпи и халати”.
- Current price → `Base Price - EUR`.
- Higher previous price → `List Price - EUR`; otherwise `N/A`.
- Product URL → `Reference Link`.
- Brand/Trademark → `Izidream` or the brand found on the product page.
- Original gallery images only (`/resources/...`).
- Weight comes from the product page and is converted from kg to g.
- Package L/W/H are estimated by product category.
- Country of origin is read from the description when present; otherwise Bulgaria.

## Built-in safety check
The script validates the finished workbook before the GitHub job uploads it. It fails when a red/category-required field or a core commercial field is blank, when a material option group has no selected value, when Color/Size is missing for the selected Variation Theme, or when Temu’s protected technical row 4 is overwritten.

## Local run
```bash
pip install -r requirements.txt
python scraper.py --limit 20
python scraper.py --limit 0
```

## Red percentage fields fix (v1.2)
- Every category-required material/composition option is written as a numeric value.
- Unselected red percentage fields are written as `0` instead of being left blank.
- Each required material group is normalized to exactly `100%`.
- Material detection uses the labelled Material/Face Fabric/Filling/Core sections to avoid false matches such as `лен` inside `неизбелван` or `пера` inside `пералня`.
- Validation fails if any red percentage cell is blank, non-numeric, or the group total differs from `100`.
