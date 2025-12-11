# BigSeller Shopee Title Prefix Helper
Automated Listing Workflow for Shopee Sellers on BigSeller

![Userscript](https://img.shields.io/badge/Userscript-Tampermonkey-blue?logo=tampermonkey)
![Platform](https://img.shields.io/badge/Platform-BigSeller-orange)
![Shopee](https://img.shields.io/badge/Marketplace-Shopee-ff5722?logo=shopee)
![Status](https://img.shields.io/badge/Status-Active-brightgreen)
![Language](https://img.shields.io/badge/Language-JavaScript-yellow)

A userscript that automates Shopee listing workflows inside BigSeller 编辑产品 pages.  
It standardizes title prefixes, description templates, SKUs, variant names, and more.

---

## ✨ Features

### 1. Store-Aware Title & Description Automation

**Title Prefix**
- Inserts the correct prefix by store  
- Removes old prefixes  
- Prevents duplicate or mixed prefixes  

**Description Templates**
- Inserts store-matching header + footer  
- Cleans old templates before replacing  
- Works for textarea / CKEditor / Quill / contenteditable  

**MD5 Auto Refresh**
- Automatically triggers `.sell_md5` after applying templates  

---

### 2. SKU Automation

#### A. 合成 SKU（Parent + Child）

- Detects parent SKU  
- Converts parent SKU: Traditional → Simplified  
- Normalizes variant SKUs into:

    父SKU-子SKU

- Strips weight suffixes:

    5g, 10g, 30ml, 100ML, -8G, -5ml …

- Avoids touching unrelated inputs  

#### B. SKU 转繁体（Variant Name Conversion）

- Locates variant name edit buttons (DOM + iframe + shadowRoot)  
- Opens popup → reads text → converts Simplified → Traditional  
- Example:

    CP365-01#蔷薇烟  →  01#薔薇煙

- Saves automatically and closes popup  
- Uses OpenCC with fallback dictionary  

---

### 3. Title Fine-Tuning Tools
Available via dropdown:

- 尾词调换  
- 學生黨平價  
- 美妝化妝品  
- 新品上市  

---

### 4. Floating Helper Panel

A UI panel appears at bottom-right:

- Shows detected store  
- Store override selector  
- Buttons:

  - 应用前缀+描述+MD5  
  - 標題微調選項  
  - 合成SKU  
  - SKU转繁体  

---

## 🧠 Technical Highlights

### DOM Targeting
- Product name: autoid="product_name_text"  
- SKU fields intelligently detected  
- Variant edit popups found across:
  - Normal DOM  
  - Ifames  
  - Shadow DOM  

### Chinese Conversion
- Uses OpenCC full build  
- Full 简 ↔ 繁 support  
- Graceful fallback  

### Stability
- Retry logic for dynamic BigSeller UI  
- Safe DOM operations  
- Ignores unrelated fields  

---

## 📦 Installation

Requirements:
- Chrome/Edge  
- Tampermonkey  

Steps:
1. Install Tampermonkey  
2. Create new userscript  
3. Paste code from this repo  
4. Save & refresh BigSeller 编辑产品 page  

---

## 🧭 Usage

1. Open BigSeller · Shopee 编辑产品页面  
2. Floating panel appears  
3. Select store  
4. Use tools:
   - 应用前缀+描述+MD5  
   - 合成SKU  
   - SKU转繁体  
   - Title fine-tuning options  

---

## 📝 Version History

### v0.95
- Fixed Title detection via autoid  
- Improved SKU合成 logic  
- Full fix for SKU转繁体 (shadow DOM + iframe)  
- Better OpenCC fallback  
- Enhanced DOM resilience  

---

## 🤝 Contributing

Suggestions and pull requests are welcome.

