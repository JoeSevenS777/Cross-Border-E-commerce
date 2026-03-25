# Inventory Formula Logic

![EXCEL](https://img.shields.io/badge/EXCEL-INVENTORY_MODEL-4d4d4d?style=for-the-badge)
![LOGIC](https://img.shields.io/badge/LOGIC-REPLENISHMENT-d4af37?style=for-the-badge)
![SCOPE](https://img.shields.io/badge/SCOPE-NETWORK_%2B_DEFAULT-1e88e5?style=for-the-badge)
![FORMULA](https://img.shields.io/badge/FORMULA-LET_STRUCTURE-43a047?style=for-the-badge)
![OUTPUT](https://img.shields.io/badge/OUTPUT-PO_DECISION-ff8f00?style=for-the-badge)
![STATUS](https://img.shields.io/badge/STATUS-STABLE-52c41a?style=for-the-badge)

## Overview

This README documents the inventory-formula framework used to control replenishment decisions for SKUs stored in **Taiwan Warehouse** and **Default Warehouse**.

It focuses on four linked parts:

* **ROP** → when the network should reorder
* **Action** → whether the shortage is network-level or Default-only
* **PO Quantity** → how much to order after packaging rules are applied
* **σD** → how demand-variability level is assigned from daily sales

In short:

```text
ROP defines WHEN to order,
Action defines WHY to order,
PO Quantity defines HOW MUCH to order,
σD adjusts the demand buffer.
```

This file is mainly for **future maintenance and recall**. When formulas need to be reviewed later, this README should make it easy to remember:

* what each formula is for
* which parameters can be changed safely
* which logic should stay stable

---

## Purpose

This document explains the **logic and maintenance rules** for three key formulas:

* **ROP (Reorder Point)**
* **Action**
* **PO Quantity**

It is designed as a quick reference to help recall the logic later.

---

## Business Context (Simple)

* One SKU exists in **Default Warehouse** and **Taiwan Warehouse**
* Stock flow: **Supplier → Default → Taiwan**
* All purchase decisions are made at the **network level**, but Default must also stay operational

---

# 1. ROP (Network ROP)

### Formula

```excel
=LET(
    sku,A2,
    r,ROW(A2),
    otherSales,IFERROR(INDEX(FILTER($I$1:$I$147,($A$1:$A$147=sku)*(ROW($A$1:$A$147)<>r)),1),0),
    otherP,IFERROR(INDEX(FILTER($P$1:$P$147,($A$1:$A$147=sku)*(ROW($A$1:$A$147)<>r)),1),0),
    otherQ,IFERROR(INDEX(FILTER($Q$1:$Q$147,($A$1:$A$147=sku)*(ROW($A$1:$A$147)<>r)),1),0),
    otherSafety,IFERROR(INDEX(FILTER($O$1:$O$147,($A$1:$A$147=sku)*(ROW($A$1:$A$147)<>r)),1),0),
    (Q2+P2)*I2 + otherSales*(otherP+otherQ+Q2) + O2 + otherSafety
)
```

### Logic (remember this)

```text
ROP = Default demand
    + Taiwan demand
    + Default safety
    + Taiwan safety
```

* Default demand = `(Q + P) × Default sales`
* Taiwan demand includes **extra waiting time (Default stage)**

👉 This is a **network-level ROP**, not local ROP

---

# 2. Action

### Formula

```excel
=LET(
    sku,A2,
    r,ROW(A2),
    otherStock,IFERROR(INDEX(FILTER($H$1:$H$147,($A$1:$A$147=sku)*(ROW($A$1:$A$147)<>r)),1),0),
    otherSales,IFERROR(INDEX(FILTER($I$1:$I$147,($A$1:$A$147=sku)*(ROW($A$1:$A$147)<>r)),1),0),
    wholeStock,H2+otherStock,
    wholeSales,I2+otherSales,
    networkROP,L2,
    defaultROP,(Q2+P2)*I2+O2,
    IF(
        wholeStock<=networkROP,
        "Place Order - Network",
        IF(
            H2<=defaultROP,
            "Place Order - Default",
            IF(
                wholeStock-networkROP<wholeSales,
                "Check",
                "No Action Needed"
            )
        )
    )
)
```

### Logic (very important)

Order of decisions:

```text
1. If whole stock ≤ network ROP → Place Order - Network
2. Else if Default stock ≤ Default ROP → Place Order - Default
3. Else if close to shortage → Check
4. Else → No Action Needed
```

👉 Key idea:

* **Network first**, then **Default local**

---

# 3. PO Quantity

### Formula

```excel
=LET(
    minQty,20,
    multiple,10,
    trigger,3,

    sku,A2,
    r,ROW(A2),
    action,N2,

    otherStock,IFERROR(INDEX(FILTER($H$1:$H$147,($A$1:$A$147=sku)*(ROW($A$1:$A$147)<>r)),1),0),
    wholeStock,H2+otherStock,
    networkROP,L2,
    defaultROP,(Q2+P2)*I2+O2,

    networkGap,MAX(networkROP-wholeStock,0),
    defaultGap,MAX(defaultROP-H2,0),

    rawQty,
        IF(
            action="Place Order - Network",
            MAX(networkGap,defaultGap),
            IF(
                action="Place Order - Default",
                defaultGap,
                0
            )
        ),

    IF(
        rawQty<=0,
        0,
        LET(
            q,MAX(rawQty,minQty),
            base,QUOTIENT(q,multiple)*multiple,
            IF(q-base>trigger,base+multiple,base)
        )
    )
)
```

---

## Logic (remember this clearly)

### Step 1 — decide base quantity

```text
Network PO → use MAX(network gap, default gap)
Default PO → use default gap
Else → 0
```

---

### Step 2 — apply packaging rule

```text
q = max(rawQty, minQty)
base = floor(q to multiple)

if remainder > trigger → round up
else → stay
```

---

## Parameters (only change these)

At the top of the formula:

```text
minQty   → minimum order
multiple → package size
trigger  → when to round up
```

### Examples

| Rule       | minQty | multiple | trigger |
| ---------- | ------ | -------- | ------- |
| Normal     | 10     | 10       | 0       |
| Current    | 20     | 10       | 3       |
| 300 pack   | 300    | 300      | 100     |
| 1000 + 500 | 1000   | 500      | 200     |

---

## Key Understanding (Most Important)

### 1. Two types of "Place Order"

```text
Network → whole system shortage
Default → only Default is low
```

### 2. ROP is network-based

* Do NOT use it alone for Default decisions

### 3. Packaging rule is independent

* Always applied **after rawQty**
* Controlled only by 3 parameters

---

## Maintenance Notes

### If something looks wrong:

Check in order:

1. Action result (Network / Default / Check)
2. rawQty (gap calculation)
3. minQty / multiple / trigger

---

### If adding new rules

You do NOT change formula structure.

Only change:

```text
minQty / multiple / trigger
```

---

---

# 4. σD (Demand Variability)

### Formula

```excel
=IFS(
    I2>=22.5,8,
    I2>=17.5,7,
    I2>=12.5,5,
    I2>=7.5,4,
    I2>=2.5,3,
    TRUE,1
)
```

---

## Logic (important)

σD is determined **only by daily sales (column I)**.

It is divided into **sales bands**:

| Daily Sales | σD |
| ----------- | -- |
| ≥ 22.5      | 8  |
| ≥ 17.5      | 7  |
| ≥ 12.5      | 5  |
| ≥ 7.5       | 4  |
| ≥ 2.5       | 3  |
| < 2.5       | 1  |

---

## Why this is done

Instead of manually tuning σD for each SKU:

```text
σD now automatically adjusts based on sales level
```

Meaning:

* High sales → higher σD → more safety
* Low sales → lower σD → leaner stock

---

## Key idea

```text
σD represents demand uncertainty, not target stock level
```

Stock level is controlled by:

* ROP
* PO logic
* packaging rules

σD only adjusts how much **buffer** is needed.

---

## Maintenance

If adjustment is needed:

* Change only the threshold numbers (22.5, 17.5, etc.)
* Do NOT change formula structure

---

## Summary (one sentence)

```text
ROP defines WHEN to order, Action defines WHY to order, PO Quantity defines HOW MUCH to order (with packaging rules).
```
