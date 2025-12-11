Excel Protected Sheet Header Sorting System
Reliable Sorting + Colour-Like Sorting for the “action” Column

This project provides an Excel VBA module that enables:

Sorting on protected sheets

Clickable ▼ sort buttons in every header

Full-table sorting that always detects the correct data range

Colour-like sorting for the action column (works even with conditional formatting colours)

✨ Features Overview
1. Header Sort Buttons (▼)

Automatically created using SetupHeaderSortButtons

One button per header cell

Runs on protected sheets

Toggles ascending/descending sorting for normal columns

2. Full-Range Table Sorting

The system dynamically detects the table’s true dimensions:

Finds the last used column in the header row

Scans all table columns to find the deepest used row

Always sorts the entire dataset, even if some columns have blank areas

3. Colour-Like Sorting for the “action” Column

Excel cannot sort by conditional-format colours — this module solves it.

How it works:

User clicks a coloured cell (red/green) in the action column

User clicks the ▼ button on the action header

VBA reads the displayed cell colour

A temporary key column is created

Rows matching the selected colour → key = 1

All other rows → key = 2

The entire table is sorted with key = 1 rows first

Temporary key column is cleared

This produces intuitive colour-group sorting without relying on Excel's unreliable colour sort engine.

4. Buttons Always Map to the Correct Column

Instead of depending on button names (which break when columns move), the macro determines the column by:

Reading the button’s TopLeftCell.Column

This ensures correct behaviour even after column insertions or deletions.

5. Sheet Protection Fully Supported

The sheet is automatically protected using settings that allow:

Sorting

Filtering

VBA operations

while keeping user edits restricted.

🛠 Installation

Press ALT + F11 to open the VBA editor

Insert a new Standard Module

Paste the VBA code from this repository

Adjust configuration constants if needed:

HEADER_ROW = row number of header
DATA_FIRST_COL = first column of the table
PWD = optional protection password

Run SetupHeaderSortButtons to generate the ▼ buttons

🚀 How to Use
Normal Sorting (any column except action)

Click the ▼ in the header.

First click → ascending

Second click → descending

Colour-Like Sorting (action column only)

Step 1: Click a coloured action cell (red/green)
Step 2: Click the ▼ in the action header

Rows with the selected colour move to the top.

🔁 Updating Headers

If you add, remove, or edit header cells:

Run SetupHeaderSortButtons again

Old buttons are auto-deleted

New buttons are auto-created and aligned

✔ Why This System Is Better
Problem in Native Excel	How This Module Fixes It
Cannot sort protected sheets	VBA unprotects → sorts → re-protects automatically
Cannot sort by conditional formatting colours	Reads actual displayed colour and uses a helper key
Sorting only partially affects the table	Scans all columns to find the deepest used row
Buttons break when columns move	Buttons detect true column by position, not name
Duplicate/misaligned buttons	Setup automatically deletes all old buttons
📄 License

MIT License
Free for personal and commercial Excel automation projects.
