# 1688 Package Arrival Manager (Order ⇄ Tracking Sync & Arrival Updates)

![Python](https://img.shields.io/badge/Python-3.8%2B-blue)
![OpenPyXL](https://img.shields.io/badge/openpyxl-Excel%20I%2FO-orange)
![Workflow](https://img.shields.io/badge/Workflow-1688%20Logistics-informational)
![Status](https://img.shields.io/badge/Status-Production-brightgreen)

A **production‑ready logistics ledger tool** for 1688 exports.

It turns noisy, inconsistent 1688 Excel exports into a **stable, auditable arrival‑tracking workbook**, while preserving your manual notes and arrival status.

---

## Core Design (Do Not Break)

* **Each row = one (订单编号, 运单号)** relationship
* **No loss of tracking numbers** under any circumstance
* One order → multiple packages → multiple rows (duplicate 订单编号 highlighted)
* One package → multiple orders → multiple rows (duplicate 运单号 highlighted)
* Manual fields are **never overwritten** during sync

This structure is the only one that remains correct under real‑world logistics.

---

## What This Tool Does

### Mode 1 — Sync from ALL new 1688 exports

* Scans the base folder for **all unparsed `.xlsx` exports** from 1688
* Ignores:

  * `Order_Package_arrival_Confirmation.xlsx`
  * Excel temp files (`~$*`)
  * Already‑parsed exports in `./parsed files/`

Before syncing, it **asks for confirmation**:

```
发现 3 个1688导出文件
待收货订单总有 41 ，
确认请按Y/y同步，否则按其他任意键退出:
```

Only `Y / y` proceeds. Anything else exits safely.

#### During sync

* Extracts only required columns:

  * 订单编号 / 卖家公司名 / 订单状态
  * 订单创建时间 / 订单付款时间
  * 物流公司 / 运单号
* Normalizes data (numbers, floats, scientific notation)
* Builds a clean ledger where:

  * Header shows **unique counts**:

    * `订单编号（N）`
    * `运单号（M）`
  * Duplicate 订单编号 highlighted (multi‑package orders)
  * Duplicate 运单号 highlighted (shared packages)
  * Column widths auto‑adjust to content

#### After sync

* Parsed exports are moved to `./parsed files/`
* Workbook is saved safely
* **Excel opens automatically** on success

---

### Mode 2 — Paste arrival info (WeChat‑friendly)

Paste raw arrival messages such as:

```
【承运公司】圆通速递(YTO)
【运单号】YT7598163190233
恒志达，齐
```

Rules:

* Paste freely
* Press **Enter twice** to finish

#### What happens

* Text is parsed into arrival blocks
* Matching is done **by 运单号 only**
* If a 运单号 maps to multiple orders, **all rows are updated**

#### Arrival status inference (conservative)

| Note contains | 到达情况 |
| ------------- | ---- |
| 退款            | 退款   |
| 退货            | 退货   |
| 异常 / 丢        | 异常   |
| otherwise     | 已到达  |

#### Notes handling

* New notes are **appended**, never overwritten
* Existing notes are preserved

#### Completion summary

After processing, the script prints:

```
Arrival update summary:
  Parsed packages: 18
  Rows updated: 21
  Unmatched 运单号: 0
```

* Unmatched 运单号 (if any) are listed explicitly
* Workbook opens automatically on success

---

## Folder Structure

```
Package_Arrival_Confirmation/
│
├─ Package_Arrival_Management.py
├─ Package_Arrival_Management.bat
├─ Order_Package_arrival_Confirmation.xlsx
├─ <one or more 1688 exports>.xlsx
└─ parsed files/
```

---

## Requirements

* Python **3.8+**
* `openpyxl`

Install:

```
pip install openpyxl
```

---

## Usage (Windows)

Double‑click:

```
Package_Arrival_Management.bat
```

Choose:

* `1` → Sync from ALL new 1688 exports
* `2` → Paste arrival info

---

## Safety Guarantees

* Excel file locking is respected
* If the workbook is open, the script exits with a clear error
* No partial writes
* No silent data loss

---

## Troubleshooting

### “No new export workbook found”

Cause:

* All exports already moved to `./parsed files/`, or
* No new 1688 export placed in base folder

Fix:

* Export again from 1688
* Place `.xlsx` next to the script
* Run Mode 1

---

### “Permission denied”

Cause:

* `Order_Package_arrival_Confirmation.xlsx` is open in Excel

Fix:

* Close Excel
* Run again

---

## Philosophy

This tool is built for **operational correctness**, not cosmetic simplicity.

If a design choice prevents losing a 运单号, it wins.

---

## License

Internal workflow tool. Modify freely to fit your operation.
