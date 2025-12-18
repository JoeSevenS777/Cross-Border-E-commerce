# 1688 Image Scraper (Main / SKU / Details)

A Chrome (Manifest V3) extension for 1688 product detail pages that reliably downloads product images and videos into a clean, structured folder layout.  
This version is built strictly on actual 1688 page behavior, not assumptions.

---

## Features

- Base folder named after the product title (never supplier / shop name)
- Image folders:
  - main/ – main gallery images (SKU images excluded)
  - sku/ – SKU / variant images
  - details/ – description (图文详情) images
- Product video download (if present) into the base folder
- Minimum image size filter (default ≥ 500 × 500)
- No duplication between main/ and sku/
- SKU images named using real SKU names from 1688
- Sequential SKU prefixes (01_, 02_, …) matching page order
- Supports .jpg, .png, .webp
- Stable DOM + embedded-data parsing (no fragile hacks)

---

## Output Structure

Product Title/
  ├─ main/
  │   ├─ 001.jpg
  │   ├─ 002.webp
  │   └─ ...
  ├─ sku/
  │   ├─ 01_01#自然黑.jpg
  │   ├─ 02_02#自然棕.jpg
  │   └─ ...
  ├─ details/
  │   ├─ 001.jpg
  │   └─ ...
  └─ video.mp4   (if available)

---

## Installation

1. Unzip the extension package
2. Open Chrome and go to chrome://extensions
3. Enable Developer mode
4. Click Load unpacked
5. Select the extension folder
6. Open a detail.1688.com/offer/... product page
7. Click Scrape & Download

---

## Core Logic

### Product Title (Folder Name)

The base folder name is resolved in priority order:

1. Embedded JSON in inline script blocks (offerTitle, subject)
2. Meta tag: property="og:title"
3. H1 title
4. document.title (with “- 阿里巴巴” stripped)

Supplier / shop keywords such as 公司, 商行, 旗舰店, 选品中心 are filtered out.

---

### Details Images (图文详情)

1688 does not render detail images directly in the DOM.

Instead, pages embed a URL similar to:

  "detailUrl": "https://itemcdn.tmall.com/1688offer/xxxx.html"

Process:

1. Scan inline script blocks
2. Extract detailUrl
3. Fetch that HTML directly
4. Parse img tags
5. Filter by actual image dimensions

This is the only reliable way to get description images on 1688.

---

### Main Images

- Collected from the product image gallery
- Thumbnail URLs are normalized to original size
- Images smaller than minSize are discarded
- Any image used by a SKU is removed from main/

Result:
main/ contains only true non-SKU gallery images.

---

### SKU Images (Named, Ordered, De-duplicated)

SKU name source:

SKU data is extracted from embedded JSON (skuProps), not from DOM labels, for example:

  name: "01#自然黑"
  imageUrl: "https://..."

This avoids polluted names like inventory or price text.

Ordering:

- SKU images are processed in the same order as skuProps
- Each file gets a two-digit prefix: 01_, 02_, …

Deduplication:

- If multiple SKUs reference the same image URL, the image is downloaded once
- SKU images are removed from main/ automatically

---

### Product Video

The extension detects video via:

- video / source elements
- Embedded JSON fields (videoUrl, wirelessVideo)
- Direct media URLs (.mp4, .m3u8, .webm, .mov)

If found:

- Saved as video.*
- Placed in the base product folder

---

### Image Size Filtering

- Images are loaded via JavaScript Image()
- Checked using naturalWidth and naturalHeight
- Both dimensions must meet minSize
- Prevents thumbnails, icons, and placeholders

---

## Configuration

- Min size (px): default 500
- Applies to width and height
- Adjustable in the popup UI

---

## Design Principles

- Use embedded data first, DOM second
- Match real 1688 data flow
- No XHR / fetch interception
- No iframe brute-force crawling
- No SKU auto-click simulation

This keeps the extension stable, predictable, and maintainable.

---

## Notes

- Some products genuinely have no SKU images
- Some products genuinely have no detail images
- Some products have no video
- The extension reflects exactly what the seller provides

---

## Usage

For personal or internal business use.  
Respect 1688 / Alibaba terms and local regulations.

---

## Summary

This extension was built by iterating against real 1688 pages, not theory.  
It prioritizes correctness and stability over shortcuts.
