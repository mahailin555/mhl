#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
CNAS 授权项目导出工具
======================
将中联品检（佛山）检验技术有限公司在 CNAS 官网的认可授权检测项目
下载并保存为 Excel 文件。

表头：序号 | 检测对象 | 项目-参数 | 检测标准（方法）

依赖安装:
    pip install playwright beautifulsoup4 openpyxl lxml
    python -m playwright install chromium

若无法使用 Playwright，仅用 requests（需能跳过验证码时）：
    pip install requests beautifulsoup4 lxml openpyxl

用法:
    python cnas_export.py

注意：CNAS 搜索页面含验证码（CAPTCHA），脚本使用 Playwright 时
      会在需要输入验证码时暂停并提示用户手动操作后按回车继续。
"""

import re
import sys
import time
import logging
from pathlib import Path

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# ──────────────────────────────── 配置 ─────────────────────────────────────
TARGET_ORG     = "中联品检（佛山）检验技术有限公司"
SEARCH_KEYWORD = "中联品检"
OUTPUT_FILE    = "中联品检佛山_CNAS授权项目.xlsx"

# CNAS LAS 系统 URL（实际后端，不经过 www.cnas.org.cn 的 iframe 包装）
BASE_LAS       = "https://las.cnas.org.cn"
SEARCH_URL     = f"{BASE_LAS}/LAS_FQ/publish/externalQueryL1.jsp"   # 检测/校准实验室
SEARCH_URL_IB  = f"{BASE_LAS}/LAS_FQ/publish/externalQueryIB.jsp"  # 检验机构

# 授权检测项目详情页（需要 baseInfoId）
ABILITY_URL    = f"{BASE_LAS}/LAS_FQ/publish/lab/checkLabObjListView.jsp"

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/120.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "zh-CN,zh;q=0.9",
}
REQUEST_DELAY = 1.5

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%H:%M:%S",
)
log = logging.getLogger(__name__)


# ══════════════════════════════════════════════════════════════════════════════
# HTML 解析工具
# ══════════════════════════════════════════════════════════════════════════════
def _extract_base_info_id(html: str) -> str | None:
    """
    从搜索结果页或详情页中提取机构的 baseInfoId（UUID 格式）。
    该 ID 用于构造检测能力列表页面 URL。
    """
    # 常见于链接 href 中，例如 baseInfoId=2794652fb3de4fc3a4f11a6f3a22d0a0
    m = re.search(r"baseInfoId=([a-f0-9]{32})", html, re.I)
    if m:
        return m.group(1)
    # 部分页面用 UUID 带连字符格式
    m = re.search(
        r"baseInfoId=([0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12})",
        html, re.I,
    )
    return m.group(1) if m else None


def _extract_cert_dates(html: str) -> tuple[str, str]:
    """从详情页提取证书更新时间和有效期，用于构造能力列表 URL。"""
    update = re.search(r"certUpdateTs[='\"][\s:]*(\d{4}-\d{2}-\d{2})", html)
    valid  = re.search(r"validate[='\"][\s:]*(\d{4}-\d{2}-\d{2})", html)
    update_ts = update.group(1) if update else ""
    valid_ts  = valid.group(1)  if valid  else ""
    return update_ts, valid_ts


def _parse_org_list(html: str) -> list[dict]:
    """从搜索结果页解析机构列表。"""
    from bs4 import BeautifulSoup
    from urllib.parse import urljoin

    soup = BeautifulSoup(html, "lxml")
    results = []

    for table in soup.find_all("table"):
        rows = table.find_all("tr")
        for row in rows[1:]:
            cells = row.find_all(["td", "th"])
            texts = [c.get_text(strip=True) for c in cells]
            if not texts:
                continue
            # 行中须含 CNAS L/I 证书号
            cert_m = next(
                (re.search(r"CNAS\s*[LI]\d+", t, re.I) for t in texts
                 if re.search(r"CNAS\s*[LI]\d+", t, re.I)), None
            )
            if cert_m is None:
                continue
            cert_no = cert_m.group()
            org_name = texts[0]
            link = row.find("a", href=True)
            detail_url = urljoin(BASE_LAS, link["href"]) if link else None
            # 尝试直接从链接 href 提取 baseInfoId
            base_info_id = _extract_base_info_id(
                link["href"] if link else ""
            )
            results.append({
                "orgName":     org_name,
                "certNo":      cert_no,
                "detailUrl":   detail_url,
                "baseInfoId":  base_info_id,
            })

    return results


def _build_ability_url(base_info_id: str,
                       cert_update_ts: str = "",
                       validate: str = "") -> str:
    """
    构造检测能力列表页 URL。
    https://las.cnas.org.cn/LAS_FQ/publish/lab/checkLabObjListView.jsp
      ?baseInfoId=<UUID>
      &enstart=0
      &blueTooth=0
      &type=abilityL1        # abilityL1=检测, abilityL2=校准
      &orgEnOrCh=Ch
      &certUpdateTs=YYYY-MM-DD
      &validate=YYYY-MM-DD
      &attactdate=
    """
    params = (
        f"?baseInfoId={base_info_id}"
        f"&enstart=0&blueTooth=0&type=abilityL1&orgEnOrCh=Ch"
        f"&certUpdateTs={cert_update_ts}&validate={validate}&attactdate="
    )
    return ABILITY_URL + params


def _parse_items_table(html: str) -> list[dict]:
    """
    从检测能力列表页解析授权项目表格。

    目标列（CNAS 检测能力列表页的典型表头）：
      检测对象 | 检测项目/参数 | 检测标准(方法)及编号 | 限制范围 | 说明
    """
    from bs4 import BeautifulSoup

    soup = BeautifulSoup(html, "lxml")
    items = []

    # 字段别名映射
    ALIAS: dict[str, str] = {
        "检测对象":          "testObject",
        "检测项目":          "parameter",
        "项目":              "parameter",
        "参数":              "parameter",
        "项目/参数":         "parameter",
        "项目-参数":         "parameter",
        "检测项目/参数":     "parameter",
        "检测标准":          "standard",
        "标准":              "standard",
        "检测标准（方法）":  "standard",
        "检测标准(方法)":    "standard",
        "检测标准(方法)及编号": "standard",
        "检测标准（方法）及编号": "standard",
        "检测方法":          "standard",
    }

    for table in soup.find_all("table"):
        # 同时支持 <th> 和首行 <td> 作为表头
        header_row = table.find("tr")
        if not header_row:
            continue
        headers = [
            cell.get_text(strip=True)
            for cell in header_row.find_all(["th", "td"])
        ]

        # 必须同时含"检测对象"和"标准/方法"相关列
        has_obj = any("检测对象" in h for h in headers)
        has_std = any(any(k in h for k in ("标准", "方法")) for h in headers)
        if not (has_obj and has_std):
            continue

        # 映射列索引（保留找到的第一个）
        col_map: dict[str, int] = {}
        for idx, h in enumerate(headers):
            for alias, field in ALIAS.items():
                if alias in h and field not in col_map:
                    col_map[field] = idx

        log.info("  找到能力表，列映射: %s", col_map)

        for row in table.find_all("tr")[1:]:
            cells = [td.get_text(" ", strip=True) for td in row.find_all("td")]
            if len(cells) < 2:
                continue

            def _cell(field: str) -> str:
                idx = col_map.get(field)
                return cells[idx].strip() if idx is not None and idx < len(cells) else ""

            item = {
                "testObject": _cell("testObject"),
                "parameter":  _cell("parameter"),
                "standard":   _cell("standard"),
            }
            if any(item.values()):
                items.append(item)

    return items


def _total_pages(html: str) -> int:
    """从分页控件提取总页数。"""
    m = re.search(r"共\s*(\d+)\s*页", html)
    if m:
        return int(m.group(1))
    # 备用：找 totalPage 变量
    m = re.search(r"totalPage\s*[=:]\s*(\d+)", html)
    return int(m.group(1)) if m else 1


# ══════════════════════════════════════════════════════════════════════════════
# 方式 A：Playwright（推荐，可处理 JS 渲染和验证码交互）
# ══════════════════════════════════════════════════════════════════════════════
def run_with_playwright() -> tuple[dict, list[dict]]:
    from playwright.sync_api import sync_playwright, TimeoutError as PwTimeout

    def _fill_and_search(page, url: str) -> list[dict]:
        log.info("访问搜索页: %s", url)
        page.goto(url, wait_until="networkidle", timeout=30_000)
        time.sleep(1)

        # 填写机构名称输入框
        name_selectors = [
            "input[name*='orgName']", "input[name*='org_name']",
            "input[id*='orgName']",   "input[name*='name']",
            "input[placeholder*='机构']", "input[placeholder*='名称']",
            "input[type='text']:first-of-type",
        ]
        filled = False
        for sel in name_selectors:
            try:
                loc = page.locator(sel).first
                if loc.is_visible(timeout=2_000):
                    loc.fill(SEARCH_KEYWORD)
                    filled = True
                    log.info("  已填写关键词到 %s", sel)
                    break
            except Exception:
                continue

        if not filled:
            log.warning("  未找到名称输入框，跳过 %s", url)
            return []

        # 检测是否有验证码，若有则提示用户手动操作
        captcha_found = False
        for cap_sel in ["img[src*='captcha']", "img[src*='code']",
                        "#captcha", ".captcha", "img[alt*='验证码']"]:
            try:
                if page.locator(cap_sel).is_visible(timeout=1_000):
                    captcha_found = True
                    break
            except Exception:
                continue

        if captcha_found:
            print("\n" + "=" * 60)
            print("检测到验证码！请在浏览器窗口中手动填写验证码并点击查询。")
            print("完成后在此处按回车键继续...")
            print("=" * 60)
            input()
        else:
            # 自动点击查询按钮
            for bsel in [
                "button:has-text('查询')", "input[value*='查询']",
                "input[type='submit']",    "button[type='submit']",
                "a:has-text('查询')",
            ]:
                try:
                    btn = page.locator(bsel).first
                    if btn.is_visible(timeout=2_000):
                        btn.click()
                        log.info("  点击查询按钮: %s", bsel)
                        break
                except Exception:
                    continue

        page.wait_for_load_state("networkidle", timeout=20_000)
        time.sleep(REQUEST_DELAY)
        return _parse_org_list(page.content())

    def _get_base_info(page, org: dict) -> tuple[str, str, str]:
        """
        进入机构详情页，提取 baseInfoId、certUpdateTs、validate。
        """
        if org.get("baseInfoId"):
            # 已从搜索结果链接中直接提取到
            html = page.content()
            u, v = _extract_cert_dates(html)
            # 若还没进入详情页则访问
            if org.get("detailUrl"):
                page.goto(org["detailUrl"], wait_until="networkidle", timeout=30_000)
                time.sleep(REQUEST_DELAY)
                u, v = _extract_cert_dates(page.content())
            return org["baseInfoId"], u, v

        if not org.get("detailUrl"):
            raise RuntimeError("无详情页URL且无baseInfoId，无法继续。")

        log.info("访问详情页: %s", org["detailUrl"])
        page.goto(org["detailUrl"], wait_until="networkidle", timeout=30_000)
        time.sleep(REQUEST_DELAY)
        html = page.content()
        base_info_id = _extract_base_info_id(html)
        if not base_info_id:
            raise RuntimeError(
                "未能从详情页提取 baseInfoId，请手动检查页面结构。\n"
                f"详情页URL: {org['detailUrl']}"
            )
        u, v = _extract_cert_dates(html)
        return base_info_id, u, v

    def _fetch_items(page, ability_url: str) -> list[dict]:
        log.info("获取检测能力列表: %s", ability_url)
        page.goto(ability_url, wait_until="networkidle", timeout=30_000)
        time.sleep(REQUEST_DELAY)
        html = page.content()
        items = _parse_items_table(html)
        total = _total_pages(html)

        for pg in range(2, total + 1):
            log.info("  第 %d/%d 页...", pg, total)
            try:
                page.locator("a:has-text('下一页')").first.click()
                page.wait_for_load_state("networkidle", timeout=15_000)
            except Exception:
                sep = "&" if "?" in ability_url else "?"
                page.goto(f"{ability_url}{sep}currentPage={pg}",
                          wait_until="networkidle", timeout=30_000)
            time.sleep(REQUEST_DELAY)
            items.extend(_parse_items_table(page.content()))

        return items

    with sync_playwright() as pw:
        # headless=False 让用户能手动处理验证码
        browser = pw.chromium.launch(headless=False)
        ctx = browser.new_context(locale="zh-CN")
        page = ctx.new_page()
        page.set_extra_http_headers(HEADERS)

        target: dict | None = None
        for qurl in [SEARCH_URL, SEARCH_URL_IB]:
            try:
                orgs = _fill_and_search(page, qurl)
                if orgs:
                    target = next(
                        (o for o in orgs if TARGET_ORG in o["orgName"]), orgs[0]
                    )
                    log.info("找到机构: %s  证书号: %s",
                             target["orgName"], target["certNo"])
                    break
            except PwTimeout as e:
                log.warning("超时: %s", e)

        if not target:
            browser.close()
            raise RuntimeError(f"未找到机构: {TARGET_ORG}")

        base_info_id, cert_update_ts, validate = _get_base_info(page, target)
        log.info("baseInfoId: %s  certUpdateTs: %s  validate: %s",
                 base_info_id, cert_update_ts, validate)

        ability_url = _build_ability_url(base_info_id, cert_update_ts, validate)
        items = _fetch_items(page, ability_url)
        browser.close()

    return target, items


# ══════════════════════════════════════════════════════════════════════════════
# 方式 B：requests + BeautifulSoup（静态页面 / 已知 baseInfoId 时使用）
# ══════════════════════════════════════════════════════════════════════════════
def run_with_requests(base_info_id: str = "",
                      cert_update_ts: str = "",
                      validate: str = "") -> tuple[dict, list[dict]]:
    """
    若已知 baseInfoId 可直接传入，跳过搜索步骤（无需解验证码）。
    否则尝试通过 POST 搜索获取（可能因验证码失败）。
    """
    import requests
    from bs4 import BeautifulSoup
    from urllib.parse import urljoin

    session = requests.Session()
    session.headers.update(HEADERS)

    def _get(url, **kw):
        for i in range(1, 4):
            try:
                r = session.get(url, timeout=30, **kw)
                r.raise_for_status()
                r.encoding = r.apparent_encoding or "utf-8"
                time.sleep(REQUEST_DELAY)
                return r
            except requests.RequestException as e:
                log.warning("GET %s 第%d次失败: %s", url, i, e)
                time.sleep(2 ** i)
        raise RuntimeError(f"请求失败: {url}")

    target: dict = {"orgName": TARGET_ORG, "certNo": "", "baseInfoId": base_info_id}

    # 若未提供 baseInfoId，尝试搜索
    if not base_info_id:
        log.info("尝试搜索机构（注意：可能因验证码失败）...")
        for qurl in [SEARCH_URL, SEARCH_URL_IB]:
            resp = _get(qurl)
            soup = BeautifulSoup(resp.text, "lxml")
            form = soup.find("form")
            action = urljoin(qurl, form["action"]) if (form and form.get("action")) else qurl
            data = {inp.get("name"): inp.get("value", "")
                    for inp in soup.find_all("input") if inp.get("name")}
            for field in ("orgName", "org_name", "searchName", "name"):
                data[field] = SEARCH_KEYWORD

            try:
                r = session.post(action, data=data, timeout=30)
                r.encoding = r.apparent_encoding or "utf-8"
                time.sleep(REQUEST_DELAY)
                orgs = _parse_org_list(r.text)
                if orgs:
                    target = next(
                        (o for o in orgs if TARGET_ORG in o["orgName"]), orgs[0]
                    )
                    base_info_id = target.get("baseInfoId", "")
                    log.info("找到: %s  %s  baseInfoId: %s",
                             target["orgName"], target["certNo"], base_info_id)
                    break
            except Exception as e:
                log.warning("POST 搜索失败: %s", e)

    if not base_info_id:
        raise RuntimeError(
            "无法自动获取 baseInfoId。\n"
            "请用浏览器手动访问 CNAS 查询页面，在机构详情页 URL 或页面源码中\n"
            "找到 baseInfoId（32位十六进制字符串），然后以命令行参数传入：\n"
            "  python cnas_export.py <baseInfoId> [certUpdateTs] [validate]\n"
            "例：python cnas_export.py 2794652fb3de4fc3a4f11a6f3a22d0a0 2024-01-01 2027-01-01"
        )

    ability_url = _build_ability_url(base_info_id, cert_update_ts, validate)
    log.info("访问检测能力列表: %s", ability_url)
    resp = _get(ability_url)
    items = _parse_items_table(resp.text)
    total = _total_pages(resp.text)

    for pg in range(2, total + 1):
        log.info("  第 %d/%d 页...", pg, total)
        sep = "&" if "?" in ability_url else "?"
        resp = _get(f"{ability_url}{sep}currentPage={pg}")
        items.extend(_parse_items_table(resp.text))

    return target, items


# ══════════════════════════════════════════════════════════════════════════════
# 导出 Excel
# ══════════════════════════════════════════════════════════════════════════════
def export_excel(items: list[dict], org_name: str, cert_no: str,
                 out_file: str = OUTPUT_FILE) -> None:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "CNAS授权项目"

    H_FONT  = Font(name="微软雅黑", bold=True, size=11, color="FFFFFF")
    H_FILL  = PatternFill("solid", fgColor="1F6BB7")
    I_FONT  = Font(name="微软雅黑", bold=True, size=11)
    D_FONT  = Font(name="微软雅黑", size=10)
    STRIPE  = PatternFill("solid", fgColor="EEF4FC")
    CENTER  = Alignment(horizontal="center", vertical="center", wrap_text=True)
    LEFT    = Alignment(horizontal="left",   vertical="center", wrap_text=True)
    THIN    = Side(style="thin", color="BBBBBB")
    BORDER  = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

    # 机构信息行
    ws.merge_cells("A1:D1")
    ws["A1"] = f"机构名称：{org_name}    认可证书号：{cert_no}"
    ws["A1"].font = I_FONT
    ws["A1"].alignment = LEFT
    ws.row_dimensions[1].height = 24

    # 表头
    col_titles = ["序号", "检测对象", "项目-参数", "检测标准（方法）"]
    col_widths  = [7, 28, 42, 55]
    for ci, (title, width) in enumerate(zip(col_titles, col_widths), start=1):
        c = ws.cell(row=2, column=ci, value=title)
        c.font = H_FONT; c.fill = H_FILL
        c.alignment = CENTER; c.border = BORDER
        ws.column_dimensions[c.column_letter].width = width
    ws.row_dimensions[2].height = 22

    # 数据行
    for ri, item in enumerate(items, start=1):
        row = ri + 2
        for ci, val in enumerate(
            [ri, item.get("testObject", ""),
             item.get("parameter", ""), item.get("standard", "")],
            start=1,
        ):
            c = ws.cell(row=row, column=ci, value=val)
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
    # 支持命令行直接传入已知的 baseInfoId，跳过搜索/验证码步骤
    # 用法: python cnas_export.py <baseInfoId> [certUpdateTs] [validate]
    cli_args = sys.argv[1:]
    base_info_id  = cli_args[0] if len(cli_args) > 0 else ""
    cert_update_ts = cli_args[1] if len(cli_args) > 1 else ""
    validate       = cli_args[2] if len(cli_args) > 2 else ""

    log.info("目标机构: %s", TARGET_ORG)

    target: dict
    items: list[dict]

    if base_info_id:
        # 已知 baseInfoId：直接用 requests 抓取，无需处理验证码
        log.info("使用 requests 模式（baseInfoId: %s）", base_info_id)
        target, items = run_with_requests(base_info_id, cert_update_ts, validate)
    else:
        # 未知 baseInfoId：优先用 Playwright（可处理验证码交互）
        try:
            import playwright  # noqa: F401
            log.info("使用 Playwright 模式（浏览器将弹出，请根据提示操作验证码）")
            target, items = run_with_playwright()
        except ImportError:
            log.info("Playwright 未安装，回退到 requests 模式")
            target, items = run_with_requests()

    if not items:
        log.error("未解析到任何授权项目。")
        log.error("可能原因：1) 验证码未通过  2) 页面结构变化  3) 网络问题")
        sys.exit(1)

    export_excel(items, target["orgName"], target["certNo"])
    print(f"\n✓ 完成！共 {len(items)} 条授权项目，已保存到：{Path(OUTPUT_FILE).resolve()}")


if __name__ == "__main__":
    main()
