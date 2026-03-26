# Hyperlink Label Print VBA

Automated Taiwan shelf-label generation with hyperlink navigation and print preview in Excel.

![Excel](https://img.shields.io/badge/EXCEL-AUTOMATION-5B5B5B?style=for-the-badge)
![Language](https://img.shields.io/badge/LANGUAGE-VBA-C9A227?style=for-the-badge)
![Feature](https://img.shields.io/badge/FEATURE-HYPERLINK%20LABEL%20PRINT-1E88E5?style=for-the-badge)
![Workflow](https://img.shields.io/badge/WORKFLOW-TAIWAN%20SHELF%20LABELS-43A047?style=for-the-badge)
![Selection](https://img.shields.io/badge/SELECTION-ACTIVE%20SHEET-616161?style=for-the-badge)
![Output](https://img.shields.io/badge/OUTPUT-PRINT%20PREVIEW-F57C00?style=for-the-badge)
![Status](https://img.shields.io/badge/STATUS-STABLE-52C41A?style=for-the-badge)

A VBA macro that builds a **标签打印** sheet from the active data sheet, generates **one shelf label per page**, adds **hyperlinks from each shelf number to its matching label**, and opens **Print Preview**.

---

## ✨ Key Features

### 1. Hyperlink-Driven Navigation

Each valid **台湾货架位** cell in the source sheet becomes a hyperlink that jumps to the matching label position in **标签打印**.

### 2. One Label per Page

Each valid row generates one print label block intended for one page in preview.

### 3. Header-Driven Mapping

The macro locates columns by header names rather than fixed positions.

Required headers:

| Header |
| ------ |
| 台湾货架位  |
| SKU编号  |
| GTIN   |

### 4. Automatic Print Sheet Management

The macro creates **标签打印** if it does not exist, or refreshes it if it already exists.

### 5. Print Preview Workflow

After generating all labels, the macro opens Excel Print Preview automatically.

---

## 🧭 Workflow

1. Use the **active sheet** as the source data sheet.
2. Find the required headers.
3. Scan all used rows under the headers.
4. Keep rows where both **台湾货架位** and **SKU编号** are non-empty.
5. Create or refresh **标签打印**.
6. Write one label block per valid row.
7. Add a hyperlink on each shelf number in the source sheet.
8. Open **Print Preview**.

---

## 🧱 Label Layout

Each label prints these lines:

```text
A-02-02

MJ-Flortte-娇憨憨

GTIN: 6926309394395
```

Layout characteristics:

* square-oriented label composition
* horizontally centered text
* enlarged shelf number
* enlarged SKU line
* GTIN shown only when present
* extra blank lines between text blocks for readability

---

## 📌 Macro Entry Point

Run this macro:

```vb
BuildTaiwanShelfLabelPrintSheet
```

---

## 📂 Sheet Behavior

### Source sheet

The currently active worksheet must contain the header row and label data.

### Output sheet

The macro writes output to:

```text
标签打印
```

If the sheet already exists, it is cleared and rebuilt.

---

## ✅ Valid Row Rules

A row is processed only when:

* **台湾货架位** is not empty
* **SKU编号** is not empty

If **GTIN** is empty, the label is still created, but the GTIN line is omitted.

---

## 🔗 Hyperlink Behavior

After the macro runs:

* each source-sheet shelf cell links to its corresponding label block in **标签打印**
* old hyperlinks in the shelf column are removed before rebuilding

---

## 🖨️ Print Notes

The macro prepares the print layout in Excel and attempts to use a custom paper size through VBA. Actual enforcement of **100mm × 100mm** depends on the installed printer driver.

If Excel preview does not perfectly match a square label page, set the custom paper size manually in the printer properties.

---

## 🚀 Setup

### 1. Save workbook as macro-enabled

Save the workbook as:

```text
.xlsm
```

### 2. Open the VBA editor

```text
Alt + F11
```

### 3. Insert a standard module

```text
Insert → Module
```

### 4. Paste the VBA code

Paste the macro code into the standard module.

### 5. Run the macro

Return to Excel, activate the data sheet, and run:

```vb
BuildTaiwanShelfLabelPrintSheet
```

---

## 🛠️ Adjustable Settings

You can tune the layout directly in the code.

### Sheet name

```vb
Public Const PRINT_SHEET_NAME As String = "标签打印"
```

### Shelf number font size

```vb
With labelRange.Characters(1, shelfLen).Font
    .Name = "Arial"
    .Bold = True
    .Size = 48
End With
```

### SKU font size

```vb
With labelRange.Characters(skuStart, skuLen).Font
    .Name = "Microsoft JhengHei"
    .Bold = True
    .Size = 24
End With
```

### GTIN font size

```vb
With labelRange.Characters(gtinStart, gtinLen).Font
    .Name = "Arial"
    .Bold = False
    .Size = 18
End With
```

### Row height

```vb
.Rows(r).RowHeight = 28
```

### Square label width

```vb
.Columns("A").ColumnWidth = 3.2
.Columns("B").ColumnWidth = 10.2
.Columns("C").ColumnWidth = 10.2
.Columns("D").ColumnWidth = 10.2
.Columns("E").ColumnWidth = 3.2
```

---

## ⚠️ Limitations

* VBA must be stored in an **`.xlsm`** workbook, not `.xlsx`
* exact custom paper sizing depends on the printer driver
* layout in Print Preview may vary slightly across printers and Windows configurations
* hyperlinks jump to the label block position in the print sheet, not directly to a printer dialog target

---

## 📄 License

Use, adapt, and maintain internally as needed for Taiwan warehouse label-print workflows.
