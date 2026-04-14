# 1688 Browser SKU Scraper (Paste-Links Workflow)

Unified **browser-session-based** workflow for scraping **1688 商品ID / SKU ID / Spec ID**  
Playwright · Manual Login Support · Persistent Chrome Profile

![Python](https://img.shields.io/badge/Python-3.10+-blue.svg)
![Platform](https://img.shields.io/badge/Platform-Windows-yellow.svg)
![Engine](https://img.shields.io/badge/Engine-Playwright-green.svg)
![Mode](https://img.shields.io/badge/Mode-Browser%20Session-orange.svg)
![License](https://img.shields.io/badge/License-Private-red.svg)

---

This project provides a **browser-based** way to extract SKU-level identifiers from 1688 product pages.

It is designed for cases where:

- pure HTTP requests are fast but unstable
- cookie-only scraping can hit 1688 anti-bot checks
- a **real logged-in browser session** is required for reliable scraping

This version uses **Playwright + a persistent Chrome profile**.  
You log in manually when needed, and the script reuses that login state on later runs.

---

## Core Capabilities

1. **1688 SKU / Spec ID Scraping**
2. **Interactive Paste-Links Workflow**
3. **Persistent Browser Login Session**
4. **Automatic Excel Output**
5. **Failure Logging with Debug HTML**

---

## Table of Contents

- Overview
- How This Version Works
- Folder Structure
- Requirements
- Installation
- Workflow: Paste-Links Mode
- Output Format
- Example Usage
- Error Handling & Warnings
- Troubleshooting
- Notes on Performance
- File List

---

## Overview

This script scrapes product data from 1688 by using a **real browser page navigation flow**.

Compared with pure HTTP mode:

- **more reliable**
- better at surviving anti-bot checks
- can reuse a real Chrome login state

Trade-off:

- **slower than pure HTTP requests**
- because each product page is opened through the browser session

This is the currently working version when cookie-only HTTP scraping is blocked.

---

## How This Version Works

High-level flow:

1. Launch Chrome through Playwright using a **persistent browser profile**
2. Open the 1688 homepage
3. Let you **log in manually** if needed
4. Paste one or more product links into the terminal
5. Navigate to each product page in the browser
6. Parse `skuModel` from the page HTML
7. Export results to Excel
8. Save raw HTML to `debug_html/` for troubleshooting

---

## Folder Structure

```text
ID_Scrape/
│
├─ scrape_1688_http_paste_links_open.py
├─ config.py
│
├─ debug_html/
│   ├─ 1234567890.html
│   └─ ...
│
├─ playwright_1688_profile/
│   └─ ... persistent Chrome profile data ...
│
├─ pasted_links_YYYYMMDD-HHMMSS(done).xlsx
├─ failed_links_YYYYMMDD-HHMMSS.xlsx
└─ README.md
```

---

## Requirements

- Windows
- Python 3.10 or newer
- Google Chrome installed
- Python packages:
  - `playwright`
  - `pandas`
  - `openpyxl`

---

## Installation

Install dependencies:

```bash
pip install playwright pandas openpyxl
playwright install chromium
```

Make sure your `config.py` contains at least:

- `SCRAPE_FOLDER`

This browser-based version does **not** require a manually maintained cookie text file.

---

## Workflow: Paste-Links Mode

Run:

```bash
python scrape_1688_http_paste_links_open.py
```

Then:

1. Paste one or more 1688 product links, one per line
2. Press **Enter on an empty line** to finish input
3. Chrome opens through the saved Playwright profile
4. Log in to 1688 manually if required
5. Return to the terminal and press **Enter**
6. The script processes links one by one
7. Excel output is generated at the end

---

## Output Format

| Column Name | Description |
|---|---|
| 商品链接 | Product URL |
| 商品ID | 1688 offerId |
| 属性SKU | SKU attribute combination |
| SKU ID | SKU identifier |
| Spec ID | Spec identifier |
| 店铺名称 | Shop name |

---

## Example Usage

### Normal run

```bash
python scrape_1688_http_paste_links_open.py
```

### Typical terminal flow

```text
请粘贴 1688 商品链接（每行一个）。
粘贴完成后，请再输入一个空行并回车结束。

https://detail.1688.com/offer/123456789.html
https://detail.1688.com/offer/987654321.html

=== 浏览器已打开 ===
请在浏览器里手动登录 1688。
登录完成并确认能正常打开商品页后，回到终端按回车继续...
```

---

## Error Handling & Warnings

The script is designed to fail safely.

### Possible outcomes

- **No valid links**
  - warning only
  - no output file generated

- **Some links fail**
  - successful rows still exported
  - failed links exported to a separate Excel file

- **All links fail**
  - no output workbook
  - terminal shows failure summary

- **Blocked / verification page detected**
  - the page HTML is still saved into `debug_html/`
  - script reports `命中风控验证页`

---

## Troubleshooting

### 1. Script asks for login every time
Possible causes:

- the Chrome profile was not preserved correctly
- 1688 login expired
- Chrome or Playwright opened a different profile than expected

Check whether this folder is being reused:

```text
playwright_1688_profile/
```

### 2. It opens pages but gets no SKU
Possible causes:

- anti-bot verification page
- product page structure changed
- page did not fully become usable before parsing

Check the saved file in:

```text
debug_html/{offerId}.html
```

### 3. It is slower than the old HTTP version
That is expected.

Reason:

- this version uses **real browser navigation**
- each link is loaded as a real page
- this is more reliable, but slower than direct `requests`

### 4. Chrome opens but script fails immediately
Check:

- Chrome is installed
- Playwright is installed correctly
- `playwright install chromium` has been run
- `config.py` points to a valid folder path

---

## Notes on Performance

This version prioritizes **reliability over speed**.

Why it is slower:

- it uses `page.goto(...)` for product links
- it waits for browser-side navigation and challenge transitions
- it works in cases where:
  - cookie-only HTTP scraping fails
  - browser `fetch()` also gets blocked
  - plain `requests` is redirected to a verification page

In practice:

- **HTTP version** = faster, but fragile
- **browser-session version** = slower, but currently workable

---

## File List

Main runtime files:

- `scrape_1688_http_paste_links_open.py`
- `config.py`

Runtime-generated files:

- `debug_html/*.html`
- `playwright_1688_profile/*`
- `pasted_links_YYYYMMDD-HHMMSS(done).xlsx`
- `failed_links_YYYYMMDD-HHMMSS.xlsx`

---

## Recommended Usage

Use this version when:

- HTTP-only scraping is being blocked
- you want a reusable logged-in 1688 browser session
- reliability matters more than raw speed

Do **not** expect this version to match the speed of the original pure HTTP scraper.

---

**Status:** Working, browser-based fallback version  
**Design Philosophy:** Reliable session first, speed second  
**Audience:** 1688 procurement / SKU mapping / ERP data preparation
