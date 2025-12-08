 (cd "$(git rev-parse --show-toplevel)" && git apply --3way <<'EOF' 
diff --git a/README.md b/README.md
index fb01db1ecf44fdeb8cae146173cda28186638ef8..3e1f75704a9c8a0b0c2195ec5de7e272d09f1d44 100644
--- a/README.md
+++ b/README.md
@@ -1,2 +1,27 @@
-# bigseller-scrape-cleanup
-One-click cleanup for BigSeller scraped products: CN→TW title conversion, clear short description, limit long-description images to 12.
+# BigSeller Scrape Cleanup Userscript
+
+This Tampermonkey userscript streamlines BigSeller product clean up while you review scraped items on the crawl page. It converts product titles from Simplified Chinese to Traditional Chinese (Taiwan), inserts readable spacing, clears the short description field, and trims long descriptions down to the first 12 images with a single action.
+
+## Features
+- Convert product titles from CN→TW without changing English parts and add smart spacing for readability.
+- Clear the "短描述" field automatically.
+- Keep only the first 12 images in the "长描述" rich text section, removing the rest.
+- Trigger via floating button, Alt+T keyboard shortcut, or Tampermonkey menu entry.
+
+## Installation
+1. Install the [Tampermonkey](https://www.tampermonkey.net/) browser extension.
+2. Create a new userscript and paste the contents of [`bigseller-scrape-cleanup.user.js`](./bigseller-scrape-cleanup.user.js).
+3. Save the script. It will run on `https://www.bigseller.pro/web/crawl/index.htm*`.
+
+## Usage
+1. Open the BigSeller crawl page where your scraped products are listed.
+2. Use one of the following to process the current page:
+   - Click the floating button labelled `标题繁体 + 清短描 + 长描限12图` (pinned near the top-right).
+   - Press `Alt + T`.
+   - Open the Tampermonkey menu and choose `BigSeller 一键：标题繁体 + 清空短描述 + 长描述限12图`.
+3. The script will update the current product form fields in place and log actions to the browser console.
+
+## Notes
+- The script relies on the `opencc-js` CDN to handle Simplified→Traditional conversion.
+- If the page uses iframes for the long description editor, only the first 12 images inside the frame are kept.
+- The userscript runs at `document-end` and will reattach its floating button if the page reloads.
 
EOF
)
