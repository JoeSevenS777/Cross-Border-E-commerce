# Excel Protected Sheet Sorting System  
Automated Sorting Workflow for Excel Tables on Protected Sheets

![Excel](https://img.shields.io/badge/Excel-Automation-217346?style=for-the-badge&logo=microsoft-excel&logoColor=white)
![VBA](https://img.shields.io/badge/Language-VBA-yellow?style=for-the-badge)
![Sorting](https://img.shields.io/badge/Feature-Header%20Sorting-blue?style=for-the-badge)
![ActionSort](https://img.shields.io/badge/Action%20Column-Colour%20Sort-orange?style=for-the-badge)
![Protected](https://img.shields.io/badge/Protected%20Sheet-Supported-success?style=for-the-badge)
![Status](https://img.shields.io/badge/Status-Stable-brightgreen?style=for-the-badge)

A VBA module that brings modern, reliable sorting to Excel protected sheets.  
It includes header sort buttons, full-table sorting logic, and colour-like sorting for conditional formatting in the **action** column.

---

## ✨ Features

### **1. Clickable Header Sort Buttons (▼)**
- Automatically created with `SetupHeaderSortButtons`
- One button per header cell  
- Sorts on protected sheets  
- Alternates ASC / DESC for normal columns  

### **2. Full-Range Sorting Engine**
The macro automatically detects:
- The last used header column  
- The deepest used row **across ALL columns**  

This ensures sorting always covers the entire dataset — never partial.

### **3. Colour-Like Sorting for the `action` Column**
Excel cannot natively sort by conditional-format colours.  
This system simulates colour sorting:

1. User selects a coloured cell in the **action** column  
2. User clicks the ▼ button  
3. Macro reads `DisplayFormat.Interior.Color`  
4. A temporary key assigns:
   - Matching colour → **1**  
   - Other colours → **2**  
5. Table is sorted by this key  
6. Temporary helper is removed  

Used for:
- Red → *place order*  
- Green → *Check12days / check*  
- Neutral statuses  

### **4. Buttons Always Sort the Correct Column**
Buttons map to their column using:

`TopLeftCell.Column`

This makes the system **immune to column moves, insertions, and deletions**.

### **5. Protected-Sheet Compatible**
Sheets are protected with:
- Sorting allowed  
- Filtering allowed  
- UserInterfaceOnly = True  

All formulas and locked cells stay protected.

---

## 🛠 Installation

1. Open Excel → press **ALT + F11**  
2. Insert a **Standard Module**  
3. Paste the VBA source code  
4. Adjust settings if needed:
   - `HEADER_ROW`
   - `DATA_FIRST_COL`
   - `PWD`  

5. Run: `SetupHeaderSortButtons`  
   → Automatically installs new ▼ header buttons  

---

## 🚀 Usage

### **Normal Sorting (any column except action)**
Click the ▼:
- First click → ascending  
- Second click → descending  

### **Colour-Like Sorting (action column)**
1. Select a coloured cell in the action column  
2. Click the ▼ on the action header  
3. Rows with the same displayed colour move to the top  

---

## 🔁 When Table Structure Changes

If you:
- Add/remove columns  
- Edit header titles  
- Rearrange the table  

Re-run: `SetupHeaderSortButtons`  
→ Old ▼ buttons are removed and rebuilt correctly.

---

## ✔ Why This System Is Better

| Excel Limitation | This Module Solves It |
|------------------|------------------------|
| Cannot sort protected sheets | Auto-unprotect → sort → re-protect |
| Cannot sort by CF colour | Uses simulated colour ranking |
| Sorting sometimes ignores lower rows | Deepest-row detection across all columns |
| Button breaks after column changes | Uses physical button position, not name |
| Duplicate/misaligned buttons | Auto-cleanup inside SetupHeaderSortButtons |

---

## 📄 License
MIT License — free for personal and commercial automation projects.

