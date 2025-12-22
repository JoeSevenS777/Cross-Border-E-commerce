# DXM Export & Audit Automation
[![Python Version](https://img.shields.io/badge/Python-3.10+-blue)]()
[![DXM Automation](https://img.shields.io/badge/Dianxiaomi-Export%20%26%20Audit-orange)]()
[![Status](https://img.shields.io/badge/Status-Production-green)]()

Automated Dianxiaomi (DXM) workflow for exporting picklists, resolving package IDs, generating Excel outputs, and performing batch audit operations via official DXM APIs.  
Designed for cross-border e-commerce operations requiring repeatable, reliable, and fast processing of large volumes of orders.

This README is fully GitHub-safe, inside one fenced Markdown block, with no nested fences.

---

## Overview

This script automates:

- Extracting pending orders from DXM  
- Mapping order IDs to package IDs  
- Exporting picklists using DXM’s export API  
- Polling export job status  
- Downloading the final Excel file  
- Summarizing by SKU  
- Auto-auditing the exported packages  

It replaces time-consuming, error-prone manual operations with a deterministic, fast workflow.

---

## Key Features

- **Mode 1:** Export *all* pending (待审核) orders  
- **Mode 2:** Export only orders listed in a user-provided Excel workbook  
- Automatic package-ID resolution  
- Automatic polling and file retrieval  
- Automatic SKU summarization (aggregates identical SKUs)  
- Automatic batch audit (if enabled via config.py)  
- Friendly error messages for missing workbooks or cookies  
- 100% pipeline-ready (can be chained with downstream scripts)

---

## Directory Structure

    project_root/
        dxm_export_and_audit.py
        config.py
        Batch_added_to_cart/
            Locate&Audit_UnprocessedOrders_InDXM/
                your_order_workbook.xlsx
            export_files_here.xlsx

---

## Configuration

All global settings are defined in `config.py`:

| Setting | Description |
|--------|-------------|
| DXM_COOKIE_PATH | Path to dxm_cookie.txt |
| PICKLIST_FOLDER | Output folder for DXM exports |
| DRY_RUN | If True, disables audit requests |
| ENABLE_AUDIT | Controls batchAudit execution |
| USER_AGENT | Browser identity used in headers |

You may also define the cookie via environment variable:

    set DXM_COOKIE=your_cookie_here

---

## Modes

### Mode 1 — Export ALL Pending Orders
Exports every order currently in the DXM "paid / 待审核" state.

Usage:

    python dxm_export_and_audit.py
    选择: 1

---

### Mode 2 — Export ONLY Orders in Workbook
The script searches for the newest `.xlsx` file under:

    Batch_added_to_cart/Locate&Audit_UnprocessedOrders_InDXM/

**Post-processing cleanup (updated):**  
After the Mode 2 workbook has been processed **successfully**, the script will **delete the workbook** from `Locate&Audit_UnprocessedOrders_InDXM/` to prevent accidental re-processing and to keep the folder clean. If processing fails (exception / HTTP error / file read error), the workbook is **not** deleted.


If the folder is empty or missing, the script prints a friendly message and exits:

    [提示] 未检测到工作簿，请把订单工作簿放入文件夹后重新运行 Mode 2。

Usage:

    python dxm_export_and_audit.py
    选择: 2

---

## DXM API Workflow

The script interacts with DXM via authenticated POST requests:

| API | Purpose |
|-----|---------|
| list.json | Retrieve pending orders or match orderId → packageId |
| exportPickData.json | Create a picklist export job |
| checkProcess.json | Poll export job status to retrieve download link |
| batchAudit.json | Audit packages in bulk |

Workflow diagram:

    User Input → Select Mode
        ↓
    list.json (order lookup)
        ↓
    exportPickData.json (generate UUID)
        ↓
    checkProcess.json (poll until link is ready)
        ↓
    Download Excel
        ↓
    Summarize SKU quantities
        ↓
    batchAudit.json (if ENABLE_AUDIT=True)
        ↓
    Done

---

## Output Files

Downloaded picklists are stored in:

    Batch_added_to_cart/

Files are automatically SKU-summarized:

- Identical SKUs merged  
- Quantities summed  
- First occurrence row used for metadata  

---

## Automatic SKU Summarization

Example transformation:

| SKU | Quantity |
|-----|----------|
| A001 | 3 |
| A001 | 4 |
| B009 | 2 |

After summarization:

| SKU | Quantity |
|-----|----------|
| A001 | 7 |
| B009 | 2 |

---

## Example Usage

### Run With Menu
    python dxm_export_and_audit.py

### Run Mode 1 Directly (for pipelines)
    python dxm_export_and_audit.py 1

### Run Mode 2 Directly
    python dxm_export_and_audit.py 2

---

## Integration With Downstream Pipelines

Often used together with mapping and 1688 add-to-cart automation.

Recommended pipeline runner:

- `ADD ALL DXM to CART (AUTO).bat` — runs DXM → verifies a *new* picklist was exported → then runs `add_to_cart_http_1688.py` with auto-confirm enabled.

If you prefer a pure Python chain:

    python dxm_export_and_audit.py 2
    python add_to_cart_http_1688.py wholesale


---

## Error Handling

The script prints readable diagnostic messages:

- Cookie missing → clear failure message  
- Workbook missing (Mode 2) → friendly exit  
- HTTP errors → status + snippet  
- JSON parsing failures → warnings  
- Duplicate SKU detection → correct summarization  

No partial failures are silently ignored.

---

## Requirements

- Python 3.10+  
- Packages:
  - requests  
  - pandas  
  - openpyxl (if you later integrate with other tools)

---

## License

Internal operational tool for private DXM automation.  
No distribution of proprietary DXM functionality is included.

