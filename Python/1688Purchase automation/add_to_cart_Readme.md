# 1688 Add-to-Cart Script – DXM Picklist → Mapping_Data → 1688 Cart

![Python](https://img.shields.io/badge/Python-3.10+-blue)
![Status](https://img.shields.io/badge/Status-Production-green)
![Platform](https://img.shields.io/badge/Platform-Windows-lightgrey)
![Automation](https://img.shields.io/badge/Automation-DXM→1688-success)
![Mode](https://img.shields.io/badge/Mode-Wholesale%20%7C%20Consign-orange)
![Output](https://img.shields.io/badge/Output-Done%20Workbook-blueviolet)
![License](https://img.shields.io/badge/License-Private-important)

A focused Python automation script for converting a Dianxiaomi (DXM) picklist workbook into 1688 add-to-cart operations.  
It reads the latest pending `.xlsx` workbook, applies `Mapping_Data.xlsx`, validates required fields, calls the 1688 add-to-cart endpoint, writes a sorted `(done).xlsx` result workbook, highlights important `拣货备注` rows, and archives processed files.

---

## 🚀 Key Features

- Automatically detects the latest unprocessed `.xlsx` picklist
- Supports DXM raw export → `Mapping_Data.xlsx` field mapping
- Case-insensitive SKU matching
- Supports both 批发 / wholesale mode and 代发 / consign mode
- Validates `商品链接`, `Spec ID`, `数量`, and `商品ID`
- Skips or marks special `备货 / 備貨` rows
- Sends authenticated HTTP requests to 1688 cart API
- Supports DRY-RUN mode through `config.py`
- Generates a clean `(done).xlsx` result workbook
- Sorts result rows by operational priority
- Highlights non-empty `拣货备注` cells and their header in red
- Moves source and result files to `Finished_added_to_cart`

---

## 🧩 Script Role in the Workflow

```text
DXM Picklist Workbook
        │
        ▼
add_to_cart_http_1688.py
        │
        ├── Load latest .xlsx from PICKLIST_FOLDER
        ├── Apply Mapping_Data.xlsx if needed
        ├── Validate 商品链接 / Spec ID / 数量
        ├── Detect 备货 rows
        ├── Submit add-to-cart HTTP requests
        ├── Generate sorted (done).xlsx
        ├── Highlight 拣货备注 when needed
        └── Move files to Finished_added_to_cart
```

---

## 📁 Required Files and Folders

The script depends on paths defined in `config.py`.

Recommended structure:

```text
AutomationRoot/
├── config.py
├── add_to_cart_http_1688.py
├── Mapping_Data/
│   └── Mapping_Data.xlsx
├── Batch_added_to_cart/
│   ├── pending_picklist.xlsx
│   └── Finished_added_to_cart/
└── Cookies/
    └── ali_cookie.txt
```

---

## ⚙️ Configuration

The script imports these values from `config.py`:

| Config Name | Purpose |
|---|---|
| `PICKLIST_FOLDER` | Folder containing DXM picklist workbooks |
| `MAPPING_PATH` | Path to `Mapping_Data.xlsx` |
| `ALI_COOKIE_PATH` | Path to 1688 cookie file |
| `USER_AGENT` | Browser user-agent string |
| `ENABLE_ADD_TO_CART` | Enables real add-to-cart requests |
| `TIMEOUT` | HTTP request timeout |

---

## 🍪 1688 Cookie Priority

The script searches for 1688 cookies in this order:

1. Environment variable `ALI_COOKIE`
2. `ali_cookie.txt` in the script directory
3. `ali_cookie.txt` in `PICKLIST_FOLDER`
4. `ALI_COOKIE_PATH` from `config.py`
5. The internal `COOKIE` constant in the script

> Directly pasting cookies into the script is supported but not recommended.

---

## 📥 Input Workbook Requirements

The input workbook should contain at least:

| Column | Required | Notes |
|---|---:|---|
| `SKU` | Yes | Used for Mapping_Data matching |
| `数量` | Yes | Must be a positive integer |
| `商品链接` | If already mapped | Used to extract offer ID |
| `Spec ID` | If already mapped | Required by 1688 add-to-cart API |
| `拣货备注` | Optional | Non-empty rows are highlighted and sorted into an attention group |

If the workbook does **not** already contain `商品链接` and `Spec ID`, the script treats it as a DXM raw export and applies `Mapping_Data.xlsx`.

---

## 🧬 Mapping_Data.xlsx Requirements

`Mapping_Data.xlsx` should contain these columns:

| Column | Purpose |
|---|---|
| `商品選項貨號` / `商品选项货号` | SKU key |
| `商品链接` | 1688 product link |
| `商品ID` | 1688 offer ID |
| `属性SKU` | Attribute SKU |
| `SKU ID` | SKU ID |
| `Spec ID` | 1688 Spec ID |
| `主供应商` | Main supplier |

The script normalizes the key column into `SKU` internally and matches SKUs case-insensitively.

---

## 🛒 Add-to-Cart Behavior

For each valid row, the script builds and submits a POST request to:

```text
https://cart.1688.com/ajax/safe/add_to_cart_list_new.jsx
```

The request includes:

- `cargoIdentity`
- `specId`
- `amount`
- `purchaseType`
- timestamp
- cookie-authenticated headers

If `ENABLE_ADD_TO_CART = False`, the script does not send the request and marks rows as `DRY_RUN`.

---

## 🧾 Result Statuses

The output workbook includes:

| Status | Meaning |
|---|---|
| `SUCCESS` | Added to cart successfully |
| `FAILED` | Validation failed or request failed |
| `DRY_RUN` | Real add-to-cart request was disabled |

The `备注` column explains the reason, such as:

- `加入购物车成功`
- `Spec ID 为空`
- `数量错误`
- `无法从商品链接解析商品ID`
- `备货`
- HTTP response snippet or request exception

---

## 🔀 Result Sorting Order

The result workbook is sorted by operational priority:

| Sort Key | Group |
|---:|---|
| `0` | `FAILED` + `Spec ID 为空` |
| `1` | `FAILED` other errors, excluding `备货` |
| `2` | `UNKNOWN / DRY_RUN` without `拣货备注` |
| `3` | Any row with non-empty `拣货备注` |
| `4` | `SUCCESS` without `拣货备注` |
| `5` | `FAILED` + `备货` without `拣货备注` |

Key rule:

```text
Any row with non-empty 拣货备注 forms its own attention group,
regardless of whether it is SUCCESS, FAILED, or 备货.
```

---

## 🔴 Red Highlighting Rule

The script highlights `拣货备注` only when needed.

```text
If 拣货备注 has at least one non-empty cell:
    fill the 拣货备注 header in red
    fill each non-empty 拣货备注 cell in red

If 拣货备注 is completely empty:
    do not fill the header
    do not fill any cell
```

This makes special picking notes visually obvious without adding unnecessary red formatting to ordinary files.

---

## ▶️ Usage

Default wholesale mode:

```bash
python add_to_cart_http_1688.py
```

Consign mode:

```bash
python add_to_cart_http_1688.py consign
```

Alternative consign arguments:

```bash
python add_to_cart_http_1688.py consign_purchase_type
python add_to_cart_http_1688.py daifa
python add_to_cart_http_1688.py 代发
```

---

## ✅ Safety Confirmation

Before real add-to-cart execution, the script asks for confirmation:

```text
Press Y/y to continue.
Press any other key to cancel.
```

When triggered by the full pipeline `.bat`, the script can auto-confirm only if the newest workbook was freshly exported during the current pipeline run.

Supported environment variables:

| Variable | Purpose |
|---|---|
| `PIPELINE_START_EPOCH` | Confirms the file was created after the pipeline started |
| `AUTO_CONFIRM_LATEST` | Enables fallback auto-confirm for newest files |
| `AUTO_CONFIRM_WINDOW_SEC` | Freshness window in seconds, default `900` |

---

## 📤 Output

For an input file:

```text
jianhuodan.xlsx
```

The script creates:

```text
jianhuodan(done).xlsx
```

Then it moves both files to:

```text
Finished_added_to_cart/
```

If a file with the same name already exists, the script appends a timestamp to avoid overwriting.

---

## 📦 Dependencies

Install required Python packages:

```bash
pip install pandas requests openpyxl
```

Recommended Python version:

```text
Python 3.10+
```

---

## 🧠 Operational Logic Summary

```text
Find latest workbook
→ Apply mapping if needed
→ Validate each row
→ Mark 备货 rows
→ Add valid rows to 1688 cart
→ Write result workbook
→ Sort by attention priority
→ Highlight 拣货备注 when non-empty
→ Move files to Finished_added_to_cart
→ Ask whether to open result
```

---

## ⚠️ Important Notes

- Use only on accounts and data you are authorized to operate.
- Keep cookies private.
- Test with `ENABLE_ADD_TO_CART = False` before running real add-to-cart operations.
- The script processes only the latest unprocessed `.xlsx` workbook in `PICKLIST_FOLDER`.
- Files containing `(done)` in the filename are ignored as input.

---

## 🔒 License

Private internal business automation script.  
Not intended for public distribution.
