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

Paste the complete raw arrival text, even when it contains dates, image filenames, unrelated numbers, blank lines, or inconsistent formatting:

```
乐购易
2026年09月03日 18:18
韵达快递 435337210043705蓝蕴，齐

乐购易
2026年09月03日 18:57
中通快递(ZTO) 79030031408885
美妆，已申请退款
```

Rules:

* Choose Mode `2`
* Paste all arrival text into the terminal
* Blank lines are accepted as part of the pasted text
* After the pasted text, press **Enter** to start a new line
* Type `END` on that separate line and press **Enter**
* The workbook is updated and saved automatically—no confirmation is required

Example:

```
Enter 1 or 2: 2
Paste arrival text. Type END on a separate line to save automatically:

乐购易
2026年09月03日 18:18
韵达快递 435337210043705蓝蕴，齐

END
```

#### What happens

* Reads only rows whose `运单号` is present and `到达情况` is blank
* Uses those pending tracking numbers as a whitelist
* Searches for each pending tracking number anywhere in the pasted text
* Ignores dates, timestamps, image filenames, sender names, and unrelated numbers
* Supports numeric and letter‑number tracking numbers, full‑width characters, mixed case, and accidental spaces
* Marks every matched pending row as `已到达`
* If one 运单号 belongs to multiple rows, every blank‑status row is updated
* Rows with an existing 到达情况 are left unchanged

#### Notes handling

* Text after the tracking number on the same line is saved in `备注`
* If no text follows on that line, the next meaningful line is used as the note
* If there is no note, `备注` stays blank
* Notes such as `蓝蕴，齐` and `美妆，已申请退款` do not change the result: the package is still marked `已到达`
* Existing notes are preserved; a distinct new note is appended once
* Repeating the same import does not duplicate notes or update completed rows again

#### Saved update record

After saving, the script prints every changed row and every pending tracking number that was not found:

```
Arrival update record (saved):
  Pending rows checked: 10
  Pending tracking numbers checked: 10
  Tracking numbers found: 8
  Rows updated: 8
  Row 5 | 435316139221920 | 已到达 | 备注: 蓝蕴,齐
  Row 6 | 79026301712135 | 已到达 | 备注: 美妆,已申请退款
  Still not found: 2
  - JT3174236048220
  - 800214560799
```

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
* Existing nonblank arrival statuses are never changed by Mode 2
* Arrival matching is limited to pending tracking numbers already in the workbook

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

### Mode 2 finds zero tracking numbers

Cause:

* `END` was entered before the arrival text was pasted, or
* None of the pending tracking numbers appears in the pasted text

Fix:

* Choose Mode `2`
* Paste the complete arrival text first
* Type `END` on a new line after the pasted text

---

## Philosophy

This tool is built for **operational correctness**, not cosmetic simplicity.

If a design choice prevents losing a 运单号, it wins.

---

## License

Internal workflow tool. Modify freely to fit your operation.
