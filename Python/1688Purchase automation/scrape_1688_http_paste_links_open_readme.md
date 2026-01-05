# 1688 HTTP SKU Scraper (Paste-Links Workflow)

Unified **HTTP-based** workflow for scraping **1688 商品ID / SKU ID / Spec ID**  
No Selenium · No Browser Automation · Stable & Auditable

![Python](https://img.shields.io/badge/Python-3.10+-blue.svg)
![Platform](https://img.shields.io/badge/Platform-Windows-yellow.svg)
![Mode](https://img.shields.io/badge/Mode-Production-green.svg)
![License](https://img.shields.io/badge/License-Private-red.svg)

---

This script provides a **fast, low-fragility** way to extract SKU-level identifiers from 1688 product pages using **pure HTTP requests**.

It is designed for **cross-border e-commerce operators** who need reliable SKU / SpecID data for:
- ERP mapping
- Add-to-cart automation
- DXM / procurement exports
- Manual auditing & reconciliation

---

## Core Capabilities

1. **1688 SKU & SpecID Scraping**
2. **Interactive Paste-Links Workflow**
3. **Automatic Excel Output**
4. **Auto-Open Result Workbook**
5. **Failure-Safe with Debug Artifacts**

---

## Table of Contents

- Overview  
- Folder Structure  
- Global Configuration  
- Workflow: Paste-Links Mode  
- Output Format  
- Example Usage  
- Error Handling & Warnings  
- Troubleshooting  
- File Citations  

---

## Overview

This script replaces browser-based scraping (Selenium / Playwright) with a **deterministic HTTP pipeline**:

- No iframe issues  
- No login pop-ups  
- No browser crashes  
- Minimal maintenance cost  

It directly parses the `skuModel` JSON embedded in 1688 product pages and outputs structured Excel data.

---

## Folder Structure

```
ID_Scrape/
│
├─ scrape_1688_http_paste_links_open.py
├─ config.py
│
├─ debug_html/
│   ├─ 1234567890.html
│   └─ ...
│
├─ pasted_links_YYYYMMDD-HHMMSS(done).xlsx
└─ README.md
```

---

## Global Configuration

All global settings are centralized in `config.py`:

- `SCRAPE_FOLDER`
- `ALI_COOKIE_PATH`
- `USER_AGENT`
- `TIMEOUT`

A **valid 1688 login cookie** is required.

---

## Workflow: Paste-Links Mode (Default)

1. Run the script  
   ```bash
   python scrape_1688_http_paste_links_open.py
   ```

2. Paste one or multiple 1688 product links (one per line)

3. Press **Enter on an empty line** to start scraping

4. Excel output is generated and **opened automatically**

---

## Output Format

| Column Name | Description |
|------------|-------------|
| 商品链接 | Product URL |
| 商品ID | 1688 offerId |
| 属性SKU | SKU attributes |
| SKU ID | SKU identifier |
| Spec ID | Spec identifier |
| 店铺名称 | Shop name |

---

## Example Usage (Excel Mode – Optional)

```bash
python scrape_1688_http_paste_links_open.py --excel
```

Or:

```bash
python scrape_1688_http_paste_links_open.py --excel input.xlsx
```

---

## Error Handling & Warnings

- No valid links → warning, no output  
- All links fail → warning, no Excel  
- Partial failure → Excel still generated  
- Raw HTML always saved to `debug_html/`

---

## Troubleshooting

Common causes:
- Expired cookie
- Login-required product
- SKU data not exposed
- Network issues

Check `debug_html/{offerId}.html` for diagnostics.

---

## File Citations

- `scrape_1688_http_paste_links_open.py`
- `config.py`
- `debug_html/*.html`
- `README.md`

---

**Status:** Production-ready  
**Design Philosophy:** Simple · Deterministic · Low-Maintenance  
**Audience:** 1688 procurement & ERP automation
