import os
import re
import json
import time
import sys
from pathlib import Path
from datetime import datetime

import pandas as pd
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError

from config import SCRAPE_FOLDER


SCRIPT_VERSION = "scrape_1688_browser_login v2026-04-13-01"


# ======================================================================
# Directories
# ======================================================================
BASE_DIR = Path(SCRAPE_FOLDER)
BASE_DIR.mkdir(parents=True, exist_ok=True)

DEBUG_DIR = BASE_DIR / "debug_html"
DEBUG_DIR.mkdir(parents=True, exist_ok=True)

PROFILE_DIR = BASE_DIR / "playwright_1688_profile"
PROFILE_DIR.mkdir(parents=True, exist_ok=True)


# ======================================================================
# Utilities
# ======================================================================
def extract_offer_id(url: str) -> str:
    if not url:
        return ""
    s = str(url).strip()
    m = re.search(r"/offer/(\d+)\.html", s)
    if m:
        return m.group(1)
    m = re.search(r"offerId=(\d+)", s)
    if m:
        return m.group(1)
    return ""


def save_debug_html(offer_id: str, html: str) -> None:
    target = DEBUG_DIR / f"{offer_id or 'unknown'}.html"
    target.write_text(html, encoding="utf-8")


def read_links_from_stdin() -> list[str]:
    print("\n请粘贴 1688 商品链接（每行一个）。")
    print("粘贴完成后，请再输入一个空行并回车结束。\n")

    urls: list[str] = []
    seen: set[str] = set()

    while True:
        line = sys.stdin.readline()
        if not line:
            break
        line = line.strip()
        if not line:
            break

        parts = [p.strip() for p in re.split(r"\s+", line) if p.strip()]
        for p in parts:
            u = p.strip().strip('"').strip("'")
            if not u:
                continue
            if not (u.startswith("http://") or u.startswith("https://")):
                continue
            if u not in seen:
                seen.add(u)
                urls.append(u)

    return urls


def build_output_path() -> Path:
    ts = datetime.now().strftime("%Y%m%d-%H%M%S")
    return BASE_DIR / f"pasted_links_{ts}(done).xlsx"


def normalize_text(x) -> str:
    if x is None:
        return ""
    return str(x).strip()


# ======================================================================
# Page parsing
# ======================================================================
def extract_json_object_from_text(text: str, key: str) -> str | None:
    idx = text.find(key)
    if idx == -1:
        return None

    start = text.find("{", idx)
    if start == -1:
        return None

    brace_level = 0
    in_str = False
    esc = False

    for i in range(start, len(text)):
        ch = text[i]
        if in_str:
            if esc:
                esc = False
            elif ch == "\\":
                esc = True
            elif ch == '"':
                in_str = False
        else:
            if ch == '"':
                in_str = True
            elif ch == "{":
                brace_level += 1
            elif ch == "}":
                brace_level -= 1
                if brace_level == 0:
                    return text[start:i + 1]
    return None


def parse_sku_data_from_html(html: str) -> tuple[list[dict], str]:
    json_str = extract_json_object_from_text(html, "skuModel")
    if not json_str:
        return [], ""

    try:
        data = json.loads(json_str)
    except Exception:
        return [], ""

    sku_map = data.get("skuInfoMap") or {}
    records: list[dict] = []

    for _, v in sku_map.items():
        sku_id = v.get("skuId") or v.get("id") or ""
        spec_id = v.get("specId") or v.get("specIdStr") or ""
        attrs = v.get("specAttrs") or []

        if isinstance(attrs, list):
            attr_values = [str(a.get("value", "")).strip() for a in attrs if a]
            attr = "-".join([x for x in attr_values if x])
        else:
            attr = str(attrs).strip() if attrs else ""

        records.append(
            {
                "SKU ID": str(sku_id),
                "Spec ID": str(spec_id),
                "属性SKU": attr,
            }
        )

    shop_name = ""
    m = re.search(r'"companyName"\s*:\s*"([^"\\]+)"', html)
    if m:
        shop_name = m.group(1)

    return records, shop_name


def detect_blocked_page(html: str) -> bool:
    signals = [
        "_____tmd_____/punish",
        '"action":"captcha"',
        "x5secdata=",
        "window._config_",
    ]
    lowered = html.lower()
    return any(s.lower() in lowered for s in signals)


# ======================================================================
# Browser helpers
# ======================================================================
def wait_for_manual_login(page) -> None:
    print("\n=== 浏览器已打开 ===")
    print("请在浏览器里手动登录 1688。")
    print("登录完成并确认能正常打开商品页后，回到终端按回车继续...\n")
    input()


def ensure_home_ready(page) -> None:
    print("[INFO] 打开 1688 首页...")
    page.goto("https://www.1688.com/", wait_until="domcontentloaded", timeout=120000)
    time.sleep(3)


def try_pass_challenge(page, wait_seconds: int = 20) -> None:
    print(f"[INFO] 等待页面验证/跳转，最多 {wait_seconds} 秒...")
    end_time = time.time() + wait_seconds
    last_url = ""

    while time.time() < end_time:
        try:
            current = page.url
            if current != last_url:
                print(f"       当前页面: {current}")
                last_url = current
        except Exception:
            pass
        time.sleep(1)


def scrape_one_product(page, url: str) -> tuple[list[dict], str]:
    url = normalize_text(url)
    if not url:
        return [], "空链接"

    offer_id = extract_offer_id(url)
    if not offer_id:
        return [], f"无法解析 offerId: {url}"

    print(f"\n-> Scraping {url}")
    print(f"   offerId={offer_id}")

    try:
        page.goto(url, wait_until="domcontentloaded", timeout=120000)
    except PlaywrightTimeoutError:
        print("   [WARN] 首次打开页面超时，继续尝试读取当前页面内容")
    except Exception as e:
        return [], f"打开页面失败: {e}"

    time.sleep(3)
    try_pass_challenge(page, wait_seconds=15)

    try:
        html = page.content()
    except Exception as e:
        return [], f"读取页面 HTML 失败: {e}"

    save_debug_html(offer_id, html)

    if detect_blocked_page(html):
        print("   [WARN] 页面仍然是风控/验证页")
        print("   [INFO] 请查看浏览器是否需要手动点击验证或重新登录")
        return [], "命中风控验证页"

    sku_records, shop_name = parse_sku_data_from_html(html)
    print(f"   [DEBUG] 解析到 SKU 数量: {len(sku_records)}")

    if not sku_records:
        return [], "未从页面 HTML 中解析到 skuModel"

    rows: list[dict] = []
    for rec in sku_records:
        rows.append(
            {
                "商品链接": url,
                "商品ID": offer_id,
                "属性SKU": rec.get("属性SKU", ""),
                "SKU ID": rec.get("SKU ID", ""),
                "Spec ID": rec.get("Spec ID", ""),
                "店铺名称": shop_name,
            }
        )

    return rows, ""


# ======================================================================
# Main
# ======================================================================
def main() -> None:
    print("=== 1688 Browser ID Scrape ===")
    print("版本:", SCRIPT_VERSION)
    print("工作目录:", BASE_DIR)
    print("浏览器用户目录:", PROFILE_DIR)

    urls = read_links_from_stdin()
    if not urls:
        print("[WARN] 未提供任何有效商品链接。")
        return

    all_rows: list[dict] = []
    failed: list[dict] = []

    with sync_playwright() as p:
        context = p.chromium.launch_persistent_context(
            user_data_dir=str(PROFILE_DIR),
            headless=False,
            channel="chrome",
            args=[
                "--start-maximized",
                "--disable-blink-features=AutomationControlled",
            ],
            viewport={"width": 1440, "height": 900},
        )

        page = context.new_page()
        ensure_home_ready(page)
        wait_for_manual_login(page)

        for i, url in enumerate(urls, start=1):
            print(f"\n===== 进度 {i}/{len(urls)} =====")
            rows, error = scrape_one_product(page, url)
            if rows:
                all_rows.extend(rows)
            else:
                failed.append({"商品链接": url, "失败原因": error})
                print(f"   [WARN] 失败原因: {error}")

        context.close()

    if not all_rows:
        print("\n[WARN] 未获得任何 SKU 数据，不输出文件。")
        if failed:
            print("[WARN] 失败链接如下：")
            for item in failed[:20]:
                print(f"  - {item['商品链接']} | {item['失败原因']}")
            if len(failed) > 20:
                print(f"  ... 以及另外 {len(failed) - 20} 条")
        return

    out_df = pd.DataFrame(all_rows)
    out_path = build_output_path()
    out_df.to_excel(out_path, index=False)

    print(f"\n全部处理完成。输出文件：{out_path}")

    if failed:
        fail_path = BASE_DIR / f"failed_links_{datetime.now().strftime('%Y%m%d-%H%M%S')}.xlsx"
        pd.DataFrame(failed).to_excel(fail_path, index=False)
        print(f"[WARN] 有 {len(failed)} 个链接失败，失败清单：{fail_path}")

    print("\n提示：")
    print("1. 这个版本依赖真实浏览器登录状态，而不是单独 cookie 字符串。")
    print("2. 浏览器用户目录会保存在 playwright_1688_profile，下次可复用登录状态。")
    print("3. 如再次遇到验证页，请在浏览器里先手动完成验证，再继续运行。")


if __name__ == "__main__":
    main()
