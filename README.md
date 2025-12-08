# BigSeller Scrape Cleanup (Tampermonkey Script)

A lightweight one-click cleanup tool for BigSeller’s **抓取页面**.  
This script standardizes scraped product data so it is clean and ready for editing.

## Features
- **Convert 产品名称 简体 → 繁体（台湾）**  
  Preserves English brand names and formats Chinese text with readable spacing.
- **Clear 短描述**  
  Removes all short-description text.
- **Limit 长描述 to first 12 images**  
  Deletes extra images to meet platform requirements (e.g., Shopee limits).

## How to Use
1. Install Tampermonkey.  
2. Install this script from GitHub (open `.user.js` → click **Raw** → Tampermonkey prompt appears).  
3. On BigSeller Scrape Page, use:
   - Floating button  
   - **Alt + T**  
   - Tampermonkey menu command  

The script will immediately process the title, short description, and long description.

## Why Use This Script?
Scraped products often contain:
- Long unformatted titles  
- Useless short descriptions  
- Too many detail images  

This script gives you a clean, consistent starting point before moving to the Edit Page.

## Compatibility
- BigSeller Scrape Page:  
  `https://www.bigseller.pro/web/crawl/index.htm*`
- Browser: Chrome / Edge / Brave (with Tampermonkey)
