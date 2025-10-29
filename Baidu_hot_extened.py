# 百度热搜加强版.py
import time
import json
import csv
from datetime import datetime, timezone, timedelta
from typing import List, Dict, Any

import requests
from bs4 import BeautifulSoup
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry
import openpyxl
from openpyxl.styles import Font, Alignment
from openpyxl.utils import get_column_letter


BAIDU_TOP_URL = "https://top.baidu.com/board?tab=realtime"
HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8",
    "Cache-Control": "no-cache",
}

OUTPUT_EXCEL = "百度热搜加强版.xlsx"
OUTPUT_CSV = "百度热搜加强版.csv"
OUTPUT_JSON = "百度热搜加强版.json"


def build_session() -> requests.Session:
    s = requests.Session()
    retries = Retry(
        total=5,
        backoff_factor=0.6,
        status_forcelist=[429, 500, 502, 503, 504],
        allowed_methods=["GET", "HEAD", "OPTIONS"],
        raise_on_status=False,
    )
    s.mount("https://", HTTPAdapter(max_retries=retries))
    s.headers.update(HEADERS)
    return s


def parse_items(soup: BeautifulSoup) -> List[Dict[str, Any]]:
    """
    解析百度热榜DOM结构。页面结构可能会变动，以下代码使用了多选择器作为兜底方案。
    """
    records: List[Dict[str, Any]] = []
    # 常见的卡片容器选择器
    cards = soup.select("div.category-wrap_iQLoo") or soup.select("div.category-wrap") or []
    now = datetime.now(timezone(timedelta(hours=8)))  # 北京时间
    ts = now.strftime("%Y-%m-%d %H:%M:%S %z")

    for idx, card in enumerate(cards, start=1):
        # 标题提取
        title_el = card.select_one(".c-single-text-ellipsis")
        title = title_el.get_text(strip=True) if title_el else ""

        # 排名：页面有时会显示明确的排名，有时不显示，默认使用循环序号
        rank_el = card.select_one(".index_1Ew5p, .index")
        try:
            rank = int(rank_el.get_text(strip=True)) if rank_el else idx
        except Exception:
            rank = idx

        # 热度值提取
        heat_el = card.select_one(".hot-index_1Bl1a, .hot-index")
        heat = heat_el.get_text(strip=True) if heat_el else ""

        # 简介提取
        desc_el = card.select_one(".hot-desc_1m_jR, .hot-desc")
        brief = desc_el.get_text(" ", strip=True) if desc_el else ""

        # 分类/标签提取
        tag_el = card.select_one(".tag_1z8Gk, .tag")
        tag = tag_el.get_text(strip=True) if tag_el else ""

        # 链接提取（优先获取卡片主链接）
        link_el = card.select_one("a[href]")
        link = link_el["href"].strip() if link_el and link_el.has_attr("href") else ""

        # 趋势（上升/下降/持平）
        trend_el = card.select_one(".trend-icon, .trend")
        trend = trend_el.get_text(strip=True) if trend_el else ""

        records.append(
            {
                "rank": rank,
                "title": title,
                "heat": heat,
                "tag": tag,
                "brief": brief,
                "link": link,
                "trend": trend,
                "fetched_at": ts,
                "source": BAIDU_TOP_URL,
            }
        )
    return records


def save_excel(rows: List[Dict[str, Any]], path: str) -> None:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "百度热榜实时数据"  # 工作表标题汉化

    headers = ["rank", "title", "heat", "tag", "brief", "link", "trend", "fetched_at", "source"]
    ws.append(headers)

    for r in rows:
        ws.append([r.get(h, "") for h in headers])

    # 样式设置
    header_font = Font(bold=True)
    center_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left_wrap_align = Alignment(horizontal="left", vertical="top", wrap_text=True)

    # 头部样式设置
    for col_idx in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.font = header_font
        cell.alignment = center_align

    # 列宽优化
    col_widths = {
        "A": 6,   # 排名
        "B": 42,  # 标题
        "C": 10,  # 热度
        "D": 12,  # 标签
        "E": 66,  # 简介
        "F": 50,  # 链接
        "G": 8,   # 趋势
        "H": 20,  # 获取时间
        "I": 30,  # 来源
    }
    for col_letter, width in col_widths.items():
        ws.column_dimensions[col_letter].width = width

    # 正文对齐方式
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=len(headers)):
        for i, cell in enumerate(row, start=1):
            if i in (1, 3, 7):  # 排名、热度、趋势列居中
                cell.alignment = center_align
            else:
                cell.alignment = left_wrap_align

    wb.save(path)


def save_csv(rows: List[Dict[str, Any]], path: str) -> None:
    if not rows:
        return
    headers = list(rows[0].keys())
    with open(path, "w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=headers)
        writer.writeheader()
        writer.writerows(rows)


def save_json(rows: List[Dict[str, Any]], path: str) -> None:
    with open(path, "w", encoding="utf-8") as f:
        json.dump(rows, f, ensure_ascii=False, indent=2)


def main():
    session = build_session()
    resp = session.get(BAIDU_TOP_URL, timeout=10)
    resp.raise_for_status()
    html = resp.text

    soup = BeautifulSoup(html, "html.parser")
    data = parse_items(soup)

    # 过滤空标题项，避免写入无效数据
    data = [d for d in data if d.get("title")]

    save_excel(data, OUTPUT_EXCEL)
    save_csv(data, OUTPUT_CSV)
    save_json(data, OUTPUT_JSON)

    print(f"已保存 {len(data)} 条记录至：")
    print(f" - Excel：{OUTPUT_EXCEL}")
    print(f" - CSV：{OUTPUT_CSV}")
    print(f" - JSON：{OUTPUT_JSON}")


if __name__ == "__main__":
    main()