# Mapping Data Updater (1688 SKU Mapping Builder)
[![Python](https://img.shields.io/badge/Python-3.10+-blue)]()
[![Mapping Engine](https://img.shields.io/badge/Module-Mapping%20Updater-purple)]()
[![Status](https://img.shields.io/badge/Status-Production-green)]()

This module processes output from the 1688 ID Scraper and automatically updates the global Mapping_Data.xlsx file used by downstream automation (including add_to_cart and DXM workflows).

It reads the latest `(done).xlsx` file produced by the scraper, fills missing SKU option codes, and appends new mapping entries into Mapping_Data.xlsx without overwriting existing ones.

This README appears in one fenced block, with no nested fences and indentation-only code examples.

---

## Purpose

This script solves three core problems:

1. Automatically filling 商品選項貨號 (SKU option code) based on pattern inference  
2. Appending newly discovered mapping rows into Mapping_Data.xlsx  
3. Ensuring Mapping_Data opens with the cursor positioned at the newest appended rows

It functions as the authoritative builder and maintainer of your 1688 supplier–SKU mapping database.

---

## Features

- Automatically detects the most recent `(done).xlsx` from the scraper output directory
- Infers option code prefixes using group-specific sample rows
- Fills missing 商品選項貨號 by combining prefix + 属性SKU
- Appends new rows into Mapping_Data.xlsx while:
  - Avoiding duplicates
  - Preserving existing mappings
  - Maintaining consistent column order
- Adjusts Excel view so new data is visible immediately when opened
- Human-friendly terminal output with warnings and data previews

---

## Directory Structure

    project_root/
        update_mapping_from_scrape.py
        config.py
        ID_Scrape/
            A(done).xlsx (newest scraper output)
        Mapping_Data/
            Mapping_Data.xlsx (global mapping master)

---

## Configuration

The script reads central paths from config.py:

| Config Field | Description |
|--------------|-------------|
| SCRAPE_FOLDER | Folder containing scraper results |
| MAPPING_PATH | Location of Mapping_Data.xlsx |

These values are referenced internally via:

    from config import SCRAPE_FOLDER, MAPPING_PATH

---

## Input File Requirements

The `(done).xlsx` file must include these columns:

| Column | Meaning |
|--------|---------|
| 商品選項貨號 | SKU option code (may be empty for some rows) |
| 商品链接 | 1688 product URL |
| 商品ID | 1688 product ID |
| 属性SKU | Variant attribute (e.g., color-size) |
| SKU ID | 1688 SKU identifier |
| Spec ID | Internal spec identifier |
| 店铺名称 | Shop name scraped from 1688 |

Rows missing 商品選項貨號 or 属性SKU are excluded from Mapping_Data update.

---

## How Option Codes Are Auto-Filled

For each 商品ID group:

1. Identify the first row where 商品選項貨號 and 属性SKU are both non-empty  
2. Treat this row as the sample:  
   prefix = sample_code minus sample_attr (if suffix matches)  
3. Apply prefix + 属性SKU to other rows missing 商品選項貨號  

Example:

    商品選項貨號: ABC-红色  
    属性SKU: 红色  

Prefix becomes:

    ABC-

Output example:

| 商品ID | 属性SKU | 填充后商品選項貨號 |
|--------|----------|----------------------|
| 1001 | 红色 | ABC-红色 |
| 1001 | 蓝色 | ABC-蓝色 |

---

## Mapping Update Logic

New rows are appended only if:

- 商品選項貨號 is not empty  
- 属性SKU is not empty  
- 商品選項貨號 does NOT already exist in Mapping_Data.xlsx  

The script preserves all existing mapping entries.

A preview of skipped rows is shown to the user for transparency.

---

## Excel View Adjustment

After updating Mapping_Data.xlsx, the script moves the sheet’s view to the bottom:

- `topLeftCell = A(last_row - 20)`  
- Active cell set to the last mapping row  

This ensures the latest appended rows are immediately visible upon opening the file.

---

## Execution

Run the script normally:

    python update_mapping_from_scrape.py

It will automatically:

1. Locate the newest `(done).xlsx`  
2. Fill missing 商品選項貨號  
3. Append new rows to Mapping_Data.xlsx  
4. Write updated files  
5. Open both files for review  
6. Pause if warnings occurred

---

## Terminal Output Example

    ====================================================
      从 B(done).xlsx 填充 商品選項貨號 并追加到 Mapping_Data.xlsx
    ====================================================
    工作目录: ID_Scrape/
    Mapping: Mapping_Data/Mapping_Data.xlsx
    [INFO] 将处理最新的 (done) 工作簿: A(done).xlsx
    [INFO] B(done) 总行数: 60
    [INFO] 可用于 Mapping 的行数: 55
    [WARN] 跳过 5 行：商品選項貨號 或 属性SKU 为空
    [INFO] 去掉已存在的商品選項貨號后，新追加行数: 22
    [INFO] 已更新 Mapping_Data.xlsx: Mapping_Data/Mapping_Data.xlsx
    全部处理完成。

---

## Error Handling

- Missing `(done).xlsx` → Clear FileNotFoundError message  
- Missing required columns → KeyError with column name  
- save / open errors → Non-fatal warnings  

If warnings occurred, script waits for the user before exiting.

---

## Integration in the Global Pipeline

This module is **Stage 2** in the automation sequence:

    Stage 1: 1688 ID Scraper
        ↓ produces A(done).xlsx
    Stage 2: Mapping Updater (this script)
        ↓ updates Mapping_Data.xlsx
    Stage 3: Add-to-Cart Automation
        ↓ uses Mapping_Data to build cart requests
    Stage 4: DXM picklist export + SKU reconciliation

---

## License

Internal use only for automated 1688 SKU mapping.  
Not affiliated with Alibaba or Dianxiaomi.

