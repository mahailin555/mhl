#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
CNAS 授权项目导出工具
======================
将中联品检（佛山）检验技术有限公司在 CNAS 官网的认可授权检测项目
下载并保存为 Excel 文件。

表头：序号 | 检测对象 | 项目-参数 | 检测标准（方法）

依赖安装（选其一）：
  方式A（推荐，自动处理JS渲染）：
      pip install playwright openpyxl
      python -m playwright install chromium

  方式B（轻量，若页面无JS渲染）：
      pip install requests beautifulsoup4 lxml openpyxl

用法：
  python cnas_export.py
"""

import re
import sys
import time
import logging
from pathlib import Path

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# ─────────────────────────────── 配置 ────────────────────────────────────────
TARGET_ORG      = "中联品检（佛山）检验技术有限公司"
SEARCH_KEYWORD  = "中联品检"           # 搜索关键词
OUTPUT_FILE     = "中联品检佛山_CNAS授权项目.xlsx"

# CNAS 查询系统入口（检测/校准实验室 + 检验机构均尝试）
QUERY_URLS = [
    "https://las.cnas.org.cn/LAS_FQ/publish/externalQueryL1.jsp",  # 检测/校准实验室
    "https://las.cnas.org.cn/LAS_FQ/publish/externalQueryIB.jsp",  # 检验机构
]

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/120.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "zh-CN,zh;q=0.9",
}
REQUEST_DELAY = 1.5   # 请求间隔（秒）

# ─────────────────────────────── 日志 ────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%H:%M:%S",
)
log = logging.getLogger(__name__)


# ══════════════════════════════════════════════════════════════════════════════
# 工具函数：解析 HTML
# ══════════════════════════════════════════════════════════════════════════════
def _parse_org_list(html: str) -> list[dict]:
    """
    从搜索结果页解析机构列表。
    返回: [{"orgName": ..., "certNo": ..., "detailUrl": ...}, ...]
    """
    from bs4 import BeautifulSoup
    from urllib.parse import urljoin
    BASE = "https://las.cnas.org.cn"

    soup = BeautifulSoup(html, "lxml")
    results = []

    for table in soup.find_all("table"):
        rows = table.find_all("tr")
        for row in rows[1:]:
            cells = row.find_all(["td", "th"])
            texts = [c.get_text(strip=True) for c in cells]
            if not texts:
                continue
            # 找含 CNAS L/I 证书号的单元格
            cert_no = next(
                (t for t in texts if re.search(r"CNAS\s*[LI]\d+", t, re.I)), None
            )
            if cert_no is None:
                continue
            cert_no = re.search(r"CNAS\s*[LI]\d+", cert_no, re.I).group()
            org_name = texts[0] if texts else ""
            link = row.find("a", href=True)
            detail_url = urljoin(BASE, link["href"]) if link else None
            results.append(
                {"orgName": org_name, "certNo": cert_no, "detailUrl": detail_url}
            )

    return results


def _parse_items_table(html: str) -> list[dict]:
    """
    从机构详情页解析授权检测项目表格。
    返回: [{"testObject": ..., "parameter": ..., "standard": ...}, ...]
    """
    from bs4 import BeautifulSoup

    soup = BeautifulSoup(html, "lxml")
    items = []

    HEADER_ALIAS = {
        "检测对象": "testObject",
        "检测项目": "parameter",
        "项目":     "parameter",
        "参数":     "parameter",
        "项目/参数": "parameter",
        "项目-参数": "parameter",
        "检测标准": "standard",
        "标准":     "standard",
        "检测标准（方法）": "standard",
        "检测标准(方法)":  "standard",
        "检测方法": "standard",
    }

    for table in soup.find_all("table"):
        headers = [th.get_text(strip=True) for th in table.find_all("th")]
        # 必须同时含"检测对象"和"标准"相关列才是目标表
        has_obj = any("检测对象" in h for h in headers)
        has_std = any(k in h for k in ("标准", "方法") for h in headers)
        if not (has_obj and has_std):
            continue

        # 映射列索引
        col_map: dict[str, int] = {}
        for i, h in enumerate(headers):
            for alias, field in HEADER_ALIAS.items():
                if alias in h and field not in col_map:
                    col_map[field] = i

        for row in table.find_all("tr")[1:]:
            cells = [td.get_text(" ", strip=True) for td in row.find_all("td")]
            if len(cells) < 2:
                continue
            item: dict[str, str] = {
                "testObject": cells[col_map["testObject"]] if "testObject" in col_map and col_map["testObject"] < len(cells) else "",
                "parameter":  cells[col_map["parameter"]]  if "parameter"  in col_map and col_map["parameter"]  < len(cells) else "",
                "standard":   cells[col_map["standard"]]   if "standard"   in col_map and col_map["standard"]   < len(cells) else "",
            }
            # 过滤全空行
            if any(item.values()):
                items.append(item)

    return items


def _total_pages(html: str) -> int:
    """从分页文字提取总页数，例如"共10页"。"""
    m = re.search(r"共\s*(\d+)\s*页", html)
    return int(m.group(1)) if m else 1


# ══════════════════════════════════════════════════════════════════════════════
# 方式 A：Playwright（推荐）
# ══════════════════════════════════════════════════════════════════════════════
def run_with_playwright() -> tuple[dict, list[dict]]:
    """使用 Playwright 无头浏览器抓取数据。"""
    from playwright.sync_api import sync_playwright, TimeoutError as PwTimeout

    def _search(page, query_url: str) -> list[dict]:
        log.info("Playwright 访问: %s", query_url)
        page.goto(query_url, wait_until="networkidle", timeout=30000)

        # 填写机构名称输入框（常见 name/id 关键字）
        selectors = [
            "input[name*='orgName']",
            "input[name*='org_name']",
            "input[name*='name']",
            "input[id*='orgName']",
            "input[placeholder*='机构']",
            "input[placeholder*='名称']",
        ]
        filled = False
        for sel in selectors:
            try:
                inp = page.locator(sel).first
                if inp.is_visible(timeout=2000):
                    inp.fill(SEARCH_KEYWORD)
                    filled = True
                    log.info("输入关键词到: %s", sel)
                    break
            except Exception:
                continue

        if not filled:
            log.warning("未找到名称输入框，跳过 %s", query_url)
            return []

        # 点击查询按钮
        btn_selectors = [
            "button:has-text('查询')",
            "input[type='submit']",
            "input[value*='查询']",
            "a:has-text('查询')",
            "button[type='submit']",
        ]
        for bsel in btn_selectors:
            try:
                btn = page.locator(bsel).first
                if btn.is_visible(timeout=2000):
                    btn.click()
                    page.wait_for_load_state("networkidle", timeout=15000)
                    log.info("点击查询按钮: %s", bsel)
                    break
            except Exception:
                continue

        time.sleep(REQUEST_DELAY)
        return _parse_org_list(page.content())

    def _fetch_detail(page, detail_url: str) -> list[dict]:
        log.info("获取详情页: %s", detail_url)
        page.goto(detail_url, wait_until="networkidle", timeout=30000)
        time.sleep(REQUEST_DELAY)
        html = page.content()
        items = _parse_items_table(html)
        total = _total_pages(html)

        for pg in range(2, total + 1):
            log.info("  第 %d/%d 页...", pg, total)
            # 尝试点击"下一页"
            try:
                page.locator("a:has-text('下一页')").first.click()
                page.wait_for_load_state("networkidle", timeout=15000)
            except Exception:
                # 备用：URL 分页
                sep = "&" if "?" in detail_url else "?"
                page.goto(f"{detail_url}{sep}currentPage={pg}",
                          wait_until="networkidle", timeout=30000)
            time.sleep(REQUEST_DELAY)
            items.extend(_parse_items_table(page.content()))

        return items

    with sync_playwright() as pw:
        browser = pw.chromium.launch(headless=True)
        ctx = browser.new_context(
            locale="zh-CN",
            extra_http_headers={"Accept-Language": "zh-CN,zh;q=0.9"},
        )
        page = ctx.new_page()
        page.set_extra_http_headers(HEADERS)

        target: dict | None = None
        for qurl in QUERY_URLS:
            try:
                orgs = _search(page, qurl)
                if orgs:
                    target = next(
                        (o for o in orgs if TARGET_ORG in o["orgName"]), orgs[0]
                    )
                    log.info("找到机构: %s  证书号: %s", target["orgName"], target["certNo"])
                    break
            except PwTimeout as e:
                log.warning("超时: %s", e)

        if not target:
            browser.close()
            raise RuntimeError(f"未找到机构: {TARGET_ORG}")

        if not target.get("detailUrl"):
            browser.close()
            raise RuntimeError("未获取到详情页URL")

        items = _fetch_detail(page, target["detailUrl"])
        browser.close()

    return target, items


# ══════════════════════════════════════════════════════════════════════════════
# 方式 B：requests + BeautifulSoup
# ══════════════════════════════════════════════════════════════════════════════
def run_with_requests() -> tuple[dict, list[dict]]:
    """使用 requests + BeautifulSoup 抓取（适合静态/服务端渲染页面）。"""
    import requests
    from urllib.parse import urljoin
    from bs4 import BeautifulSoup

    BASE = "https://las.cnas.org.cn"
    session = requests.Session()
    session.headers.update(HEADERS)
    session.verify = True

    def _get(url, params=None):
        for i in range(1, 4):
            try:
                r = session.get(url, params=params, timeout=30)
                r.raise_for_status()
                r.encoding = r.apparent_encoding or "utf-8"
                time.sleep(REQUEST_DELAY)
                return r
            except requests.RequestException as e:
                log.warning("GET %s 第%d次失败: %s", url, i, e)
                time.sleep(2 ** i)
        raise RuntimeError(f"请求失败: {url}")

    def _post(url, data):
        for i in range(1, 4):
            try:
                r = session.post(url, data=data, timeout=30)
                r.raise_for_status()
                r.encoding = r.apparent_encoding or "utf-8"
                time.sleep(REQUEST_DELAY)
                return r
            except requests.RequestException as e:
                log.warning("POST %s 第%d次失败: %s", url, i, e)
                time.sleep(2 ** i)
        raise RuntimeError(f"请求失败: {url}")

    target: dict | None = None
    for qurl in QUERY_URLS:
        log.info("访问查询页: %s", qurl)
        resp = _get(qurl)
        soup = BeautifulSoup(resp.text, "lxml")

        # 收集所有 form 隐藏字段
        form_data: dict = {}
        for inp in soup.find_all("input"):
            n = inp.get("name")
            if n:
                form_data[n] = inp.get("value", "")

        # 注入搜索参数（尝试常见字段名）
        for field in ("orgName", "org_name", "searchName", "queryName", "name"):
            form_data[field] = SEARCH_KEYWORD

        # 确定 form action
        form = soup.find("form")
        action = urljoin(qurl, form["action"]) if (form and form.get("action")) else qurl

        log.info("提交搜索: %s -> %s", SEARCH_KEYWORD, action)
        resp = _post(action, form_data)
        orgs = _parse_org_list(resp.text)
        if orgs:
            target = next(
                (o for o in orgs if TARGET_ORG in o["orgName"]), orgs[0]
            )
            log.info("找到机构: %s  证书号: %s", target["orgName"], target["certNo"])
            break

    if not target:
        raise RuntimeError(f"未找到机构: {TARGET_ORG}，请手动确认搜索关键词。")

    if not target.get("detailUrl"):
        raise RuntimeError("未获取到详情页URL")

    # 获取详情页（含分页）
    resp = _get(target["detailUrl"])
    items = _parse_items_table(resp.text)
    total = _total_pages(resp.text)

    for pg in range(2, total + 1):
        log.info("  获取第 %d/%d 页...", pg, total)
        sep = "&" if "?" in target["detailUrl"] else "?"
        resp = _get(f"{target['detailUrl']}{sep}currentPage={pg}")
        items.extend(_parse_items_table(resp.text))

    return target, items


# ══════════════════════════════════════════════════════════════════════════════
# 导出 Excel
# ══════════════════════════════════════════════════════════════════════════════
def export_excel(items: list[dict], org_name: str, cert_no: str,
                 out_file: str = OUTPUT_FILE) -> None:
    """将授权项目列表导出为格式化 Excel 文件。"""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "CNAS授权项目"

    # 样式
    H_FONT   = Font(name="微软雅黑", bold=True, size=11, color="FFFFFF")
    H_FILL   = PatternFill("solid", fgColor="1F6BB7")
    I_FONT   = Font(name="微软雅黑", bold=True, size=11)
    D_FONT   = Font(name="微软雅黑", size=10)
    STRIPE   = PatternFill("solid", fgColor="EEF4FC")
    CENTER   = Alignment(horizontal="center", vertical="center", wrap_text=True)
    LEFT     = Alignment(horizontal="left",   vertical="center", wrap_text=True)
    THIN     = Side(style="thin", color="BBBBBB")
    BORDER   = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

    # 第1行：机构信息
    ws.merge_cells("A1:D1")
    ws["A1"] = f"机构名称：{org_name}    认可证书号：{cert_no}"
    ws["A1"].font = I_FONT
    ws["A1"].alignment = LEFT
    ws.row_dimensions[1].height = 24

    # 第2行：表头
    col_titles = ["序号", "检测对象", "项目-参数", "检测标准（方法）"]
    col_widths  = [7, 28, 42, 52]
    for ci, (title, width) in enumerate(zip(col_titles, col_widths), start=1):
        c = ws.cell(row=2, column=ci, value=title)
        c.font = H_FONT; c.fill = H_FILL
        c.alignment = CENTER; c.border = BORDER
        ws.column_dimensions[c.column_letter].width = width
    ws.row_dimensions[2].height = 22

    # 数据行
    for ri, item in enumerate(items, start=1):
        row = ri + 2
        vals = [ri, item.get("testObject", ""),
                item.get("parameter", ""), item.get("standard", "")]
        for ci, v in enumerate(vals, start=1):
            c = ws.cell(row=row, column=ci, value=v)
            c.font = D_FONT
            c.alignment = CENTER if ci == 1 else LEFT
            c.border = BORDER
            if ri % 2 == 0:
                c.fill = STRIPE
        ws.row_dimensions[row].height = 18

    ws.freeze_panes = "A3"
    wb.save(out_file)
    log.info("Excel 已保存: %s（%d 条）", Path(out_file).resolve(), len(items))


# ══════════════════════════════════════════════════════════════════════════════
# 主入口
# ══════════════════════════════════════════════════════════════════════════════
def main() -> None:
    log.info("目标机构: %s", TARGET_ORG)

    # 优先使用 Playwright（处理JS渲染）；若未安装则回退到 requests
    target: dict
    items: list[dict]
    try:
        import playwright  # noqa: F401
        log.info("使用 Playwright 模式")
        target, items = run_with_playwright()
    except ImportError:
        log.info("Playwright 未安装，使用 requests 模式")
        from bs4 import BeautifulSoup  # 触发早期错误提示
        target, items = run_with_requests()

    if not items:
        log.error("未解析到任何授权项目，请检查网页结构或网络连接。")
        sys.exit(1)

    export_excel(items, target["orgName"], target["certNo"])
    print(f"\n✓ 完成！共 {len(items)} 条授权项目，已保存到：{Path(OUTPUT_FILE).resolve()}")


if __name__ == "__main__":
    main()
