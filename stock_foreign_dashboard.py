#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
股票跌幅 vs 外資買賣超 比對 Dashboard
=====================================
來源1: 富邦 e-Broker 上市股價跌幅排行 (Big5 HTML)
來源2: TWSE 證交所 外資買賣超彙總表 (JSON API)
輸出:  一頁式 HTML Dashboard

用法: python stock_foreign_dashboard.py
"""

import requests
import urllib3
import ssl
import re
import os
import webbrowser
import time
from datetime import datetime, timedelta
from bs4 import BeautifulSoup
from requests.adapters import HTTPAdapter

# ============================================================
# SSL 修復: 富邦網站憑證缺少 Subject Key Identifier，
# Python 3.14 預設會拒絕。以下建立自訂 SSL adapter 來處理。
# 注意: 僅針對富邦網站使用，TWSE 仍使用預設安全驗證。
# ============================================================
class FubonSSLAdapter(HTTPAdapter):
    """自訂 SSL Adapter，放寬對富邦網站的憑證驗證"""
    def init_poolmanager(self, *args, **kwargs):
        ctx = ssl.create_default_context()
        ctx.check_hostname = False
        ctx.verify_mode = ssl.CERT_NONE
        kwargs["ssl_context"] = ctx
        return super().init_poolmanager(*args, **kwargs)

# 關閉 InsecureRequestWarning (僅針對富邦)
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# ============================================================
# 設定區
# ============================================================
# 富邦跌幅排行頁面 (上市 5日跌幅)
FUBON_URL_5D = "https://fubon-ebrokerdj.fbs.com.tw/z/zg/zg_AA_0_5.djhtm"
# 富邦跌幅排行頁面 (上市 10日跌幅)
FUBON_URL_10D = "https://fubon-ebrokerdj.fbs.com.tw/z/zg/zg_AA_0_10.djhtm"

# TWSE 三大法人買賣超日報 JSON API (T86 = 個股明細)
TWSE_FOREIGN_URL = "https://www.twse.com.tw/rwd/zh/fund/T86"

# 輸出 HTML 檔案名稱
OUTPUT_HTML = "stock_foreign_dashboard.html"

# requests 共用 headers
HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/120.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "zh-TW,zh;q=0.9,en-US;q=0.8,en;q=0.7",
}

# 請求間隔 (秒), 避免被封鎖
REQUEST_DELAY = 3

# ============================================================
# 共用 Session (Python 3.14 SSL 嚴格模式修正)
# 富邦 & TWSE 的憑證皆缺少 Subject Key Identifier，
# Python 3.14 預設會拒絕，僅針對這兩個 domain 放寬 SSL。
# ============================================================
SESSION = requests.Session()
SESSION.mount("https://fubon-ebrokerdj.fbs.com.tw", FubonSSLAdapter())
SESSION.mount("https://www.twse.com.tw", FubonSSLAdapter())
SESSION.headers.update(HEADERS)


# ============================================================
# 第一步: 抓取富邦跌幅排行
# ============================================================
def fetch_fubon_ranking(url, label=""):
    """
    抓取富邦 e-Broker 上市股價跌幅排行 (通用)

    注意: 富邦頁面的 rank 1-2 跟表頭被塞在同一個 <tr> 裡，
    且 rank 1-2 的漲跌幅欄有額外空白 cell (9 cells vs 正常 8 cells)。
    因此不能用 <tr> 為邊界解析，改為收集所有 <td> 後逐一掃描。
    """
    print(f"   正在抓取 {label}...")

    resp = SESSION.get(url, timeout=30, verify=False)
    resp.encoding = "big5"

    soup = BeautifulSoup(resp.text, "html.parser")

    # 擷取頁面日期 (格式: "日期：02/05" 或 "日期:02/05")
    page_date = ""
    page_text = soup.get_text()
    date_match = re.search(r"日期[：:]\s*(\d{1,2}/\d{1,2})", page_text)
    if date_match:
        page_date = date_match.group(1)  # e.g. "02/05"
        print(f"   → 頁面資料日期: {page_date}")

    def clean_num(text):
        """清理數字字串，去除逗號和空白"""
        text = text.strip().replace(",", "").replace(" ", "")
        if not text or text == "-":
            return 0.0
        text = text.replace("+", "")
        try:
            return float(text)
        except ValueError:
            return 0.0

    # 收集所有 <td> (依 DOM 順序)
    all_tds = soup.find_all("td")

    stocks = []
    i = 0
    while i < len(all_tds):
        cell_text = all_tds[i].get_text(strip=True)

        # 尋找「名次」: 純數字 1~999
        if cell_text.isdigit() and 1 <= int(cell_text) <= 999:
            rank = int(cell_text)

            # 下一個 cell 應該是「股票名稱」(含連結)
            if i + 1 >= len(all_tds):
                break
            name_td = all_tds[i + 1]
            stock_name_raw = name_td.get_text(strip=True)

            # 從連結中擷取股票代號
            link = name_td.find("a")
            stock_code = ""
            if link and "href" in link.attrs:
                href = link["href"]
                match = re.search(r"Link2Stk\('([^']+)'\)", href)
                if match:
                    stock_code = match.group(1)

            if not stock_code:
                match = re.match(r"(\d{4,6}[A-Z]?)", stock_name_raw)
                if match:
                    stock_code = match.group(1)

            # 沒有股票代號就跳過
            if not stock_code:
                i += 1
                continue

            # 擷取股票名稱 (去除代號)
            stock_name = re.sub(r"^\d{4,6}[A-Z]?\s*", "", stock_name_raw).strip()

            # 接下來的 cells: 收盤價, 漲跌, [可能的空白cell], 漲跌幅, 成交量, N日漲跌, N日跌幅
            # rank 1-2 有額外空白 cell，所以需要動態判斷
            # 策略: 從 i+2 開始，收集接下來的 cells 直到找到 6 個有效數值欄位
            remaining = []
            j = i + 2
            while j < len(all_tds) and len(remaining) < 8:
                val = all_tds[j].get_text(strip=True)
                # 遇到下一個 rank 數字就停
                if val.isdigit() and 1 <= int(val) <= 999 and len(remaining) >= 6:
                    break
                # 跳過空白 cell
                if val == "":
                    j += 1
                    continue
                remaining.append(val)
                j += 1

            if len(remaining) >= 6:
                close_price = clean_num(remaining[0])
                change = clean_num(remaining[1])
                change_pct = clean_num(remaining[2].replace("%", ""))
                volume = clean_num(remaining[3])
                nd_change = clean_num(remaining[4])
                nd_pct = clean_num(remaining[5].replace("%", ""))

                stocks.append({
                    "rank": rank,
                    "code": stock_code,
                    "name": stock_name,
                    "close": close_price,
                    "change": change,
                    "change_pct": change_pct,
                    "volume": volume,
                    "five_day_change": nd_change,
                    "five_day_pct": nd_pct,
                })

                i = j  # 跳到已消耗的位置
                continue

        i += 1

    print(f"   → 成功取得 {len(stocks)} 檔股票跌幅資料")
    return stocks, page_date


# ============================================================
# 第二步: 抓取 TWSE 外資買賣超資料
# ============================================================
def fetch_twse_foreign_data(target_date=None):
    """
    抓取 TWSE 外資買賣超彙總表
    target_date: YYYYMMDD 格式，預設為今天
    """
    print("[2/3] 正在抓取 TWSE 外資買賣超資料...")

    if target_date is None:
        target_date = datetime.now().strftime("%Y%m%d")

    params = {
        "date": target_date,
        "selectType": "ALL",
        "response": "json",
    }

    resp = SESSION.get(
        TWSE_FOREIGN_URL, params=params, timeout=30, verify=False
    )
    data = resp.json()

    foreign_map = {}  # {股票代號: {買張, 賣張, 淨買賣超}}
    actual_date = ""  # 實際取得資料的日期

    if data.get("stat") == "OK" and data.get("data"):
        for row in data["data"]:
            # T86 欄位: [證券代號, 證券名稱, 外陸資買進股數(不含外資自營商),
            #           外陸資賣出股數(不含外資自營商), 外陸資買賣超股數(不含外資自營商),
            #           外資自營商買進股數, 外資自營商賣出股數, 外資自營商買賣超股數,
            #           投信買進股數, 投信賣出股數, 投信買賣超股數,
            #           自營商買賣超股數, ...]
            code = str(row[0]).strip()
            name = str(row[1]).strip()

            def parse_shares(val):
                """解析股數 (可能有逗號)"""
                val = str(val).strip().replace(",", "")
                try:
                    return int(val)
                except ValueError:
                    return 0

            buy_shares = parse_shares(row[2])
            sell_shares = parse_shares(row[3])
            net_shares = parse_shares(row[4])

            foreign_map[code] = {
                "name": name,
                "buy": buy_shares // 1000,     # 轉為張
                "sell": sell_shares // 1000,    # 轉為張
                "net": net_shares // 1000,      # 轉為張
                "buy_shares": buy_shares,
                "sell_shares": sell_shares,
                "net_shares": net_shares,
            }

        actual_date = f"{target_date[4:6]}/{target_date[6:8]}"
        print(f"   → 成功取得 {len(foreign_map)} 檔外資買賣超資料 (日期: {target_date})")
    else:
        # 如果今天沒資料，往前找最近的交易日
        stat_msg = data.get("stat", "未知")
        print(f"   → 日期 {target_date} 無資料 (stat={stat_msg})，嘗試前一交易日...")

        dt = datetime.strptime(target_date, "%Y%m%d")
        for i in range(1, 8):  # 最多往前找 7 天
            prev_dt = dt - timedelta(days=i)
            prev_date = prev_dt.strftime("%Y%m%d")
            time.sleep(REQUEST_DELAY)

            params["date"] = prev_date
            resp2 = SESSION.get(
                TWSE_FOREIGN_URL, params=params, timeout=30, verify=False
            )
            data2 = resp2.json()

            if data2.get("stat") == "OK" and data2.get("data"):
                for row in data2["data"]:
                    code = str(row[0]).strip()
                    name = str(row[1]).strip()

                    def parse_shares2(val):
                        val = str(val).strip().replace(",", "")
                        try:
                            return int(val)
                        except ValueError:
                            return 0

                    buy_shares = parse_shares2(row[2])
                    sell_shares = parse_shares2(row[3])
                    net_shares = parse_shares2(row[4])

                    foreign_map[code] = {
                        "name": name,
                        "buy": buy_shares // 1000,
                        "sell": sell_shares // 1000,
                        "net": net_shares // 1000,
                        "buy_shares": buy_shares,
                        "sell_shares": sell_shares,
                        "net_shares": net_shares,
                    }

                actual_date = f"{prev_date[4:6]}/{prev_date[6:8]}"
                print(f"   → 成功取得 {len(foreign_map)} 檔外資資料 (日期: {prev_date})")
                break
        else:
            print("   ⚠ 最近 7 天都無外資資料，請確認是否為休市期間")

    return foreign_map, actual_date


# ============================================================
# 第三步: 比對 + 產生 HTML Dashboard
# ============================================================
def merge_and_classify(stocks, foreign_map):
    """將跌幅排行與外資買賣超比對合併，分為逢低布局/持續看空"""
    buying_list = []   # 外資逢低買入
    selling_list = []  # 外資持續賣出
    no_data_list = []  # 無外資資料

    for s in stocks:
        code = s["code"]
        fdata = foreign_map.get(code)

        if fdata:
            merged = {**s, **fdata}
            if fdata["net"] > 0:
                buying_list.append(merged)
            else:
                selling_list.append(merged)
        else:
            merged = {
                **s,
                "buy": None, "sell": None, "net": None,
                "buy_shares": None, "sell_shares": None, "net_shares": None,
            }
            no_data_list.append(merged)

    # 逢低布局: 按外資淨買張數 由大到小排列
    buying_list.sort(key=lambda x: x["net"], reverse=True)
    # 持續看空: 按外資淨賣張數 由小到大排列 (賣最多在最前)
    selling_list.sort(key=lambda x: x["net"])

    return buying_list, selling_list, no_data_list


def generate_html(buying_list, selling_list, no_data_list,
                  date_5d="", date_10d="", date_foreign=""):
    """產生 HTML Dashboard"""
    print("[3/3] 正在產生 HTML Dashboard...")

    now_str = datetime.now().strftime("%Y/%m/%d %H:%M")
    total = len(buying_list) + len(selling_list) + len(no_data_list)

    def fmt_num(val, is_pct=False):
        """格式化數字"""
        if val is None:
            return '<span class="na">N/A</span>'
        if is_pct:
            cls = "pos" if val > 0 else "neg" if val < 0 else ""
            sign = "+" if val > 0 else ""
            return f'<span class="{cls}">{sign}{val:.2f}%</span>'
        else:
            cls = "pos" if val > 0 else "neg" if val < 0 else ""
            sign = "+" if val > 0 else ""
            if isinstance(val, float):
                return f'<span class="{cls}">{sign}{val:,.2f}</span>'
            else:
                return f'<span class="{cls}">{sign}{val:,}</span>'

    def make_table_rows(items, group_type):
        """產生表格行"""
        rows_html = ""
        for i, item in enumerate(items, 1):
            net_val = item["net"]
            if net_val is not None:
                net_cls = "buy-highlight" if net_val > 0 else "sell-highlight"
                net_display = fmt_num(net_val)
            else:
                net_cls = ""
                net_display = '<span class="na">N/A</span>'

            buy_display = fmt_num(item["buy"]) if item["buy"] is not None else '<span class="na">N/A</span>'
            sell_display = fmt_num(item["sell"]) if item["sell"] is not None else '<span class="na">N/A</span>'

            rows_html += f"""
            <tr>
                <td class="rank-cell">{item['rank']}</td>
                <td class="code-cell">{item['code']}</td>
                <td class="name-cell">{item['name']}</td>
                <td class="num-cell">{item['close']:,.2f}</td>
                <td class="num-cell">{fmt_num(item['five_day_change'])}</td>
                <td class="num-cell">{fmt_num(item['five_day_pct'], is_pct=True)}</td>
                <td class="num-cell">{fmt_num(item.get('ten_day_change'))}</td>
                <td class="num-cell">{fmt_num(item.get('ten_day_pct'), is_pct=True)}</td>
                <td class="num-cell">{fmt_num(item['volume'])}</td>
                <td class="num-cell">{buy_display}</td>
                <td class="num-cell">{sell_display}</td>
                <td class="num-cell {net_cls}">{net_display}</td>
            </tr>"""
        return rows_html

    buying_rows = make_table_rows(buying_list, "buy")
    selling_rows = make_table_rows(selling_list, "sell")
    nodata_rows = make_table_rows(no_data_list, "nodata")

    html = f"""<!DOCTYPE html>
<html lang="zh-TW">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>跌幅 vs 外資買賣超 Dashboard</title>
<style>
  @import url('https://fonts.googleapis.com/css2?family=Noto+Sans+TC:wght@300;400;500;700;900&family=JetBrains+Mono:wght@400;600&display=swap');

  :root {{
    --bg-primary: #0a0e17;
    --bg-card: #111827;
    --bg-card-alt: #1a2332;
    --border: #1e2d3d;
    --text-primary: #e2e8f0;
    --text-secondary: #8899aa;
    --text-muted: #4a5568;
    --accent-green: #10b981;
    --accent-green-bg: rgba(16, 185, 129, 0.08);
    --accent-red: #ef4444;
    --accent-red-bg: rgba(239, 68, 68, 0.08);
    --accent-amber: #f59e0b;
    --accent-blue: #3b82f6;
    --header-bg: #0d1320;
  }}

  * {{ margin: 0; padding: 0; box-sizing: border-box; }}

  body {{
    font-family: 'Noto Sans TC', sans-serif;
    background: var(--bg-primary);
    color: var(--text-primary);
    min-height: 100vh;
    line-height: 1.6;
  }}

  .top-bar {{
    background: linear-gradient(135deg, #0d1117 0%, #161b22 100%);
    border-bottom: 1px solid var(--border);
    padding: 20px 40px;
    display: flex;
    justify-content: space-between;
    align-items: center;
  }}

  .top-bar h1 {{
    font-size: 22px;
    font-weight: 700;
    letter-spacing: 1px;
    background: linear-gradient(135deg, #e2e8f0, #94a3b8);
    -webkit-background-clip: text;
    -webkit-text-fill-color: transparent;
  }}

  .top-bar .meta {{
    font-family: 'JetBrains Mono', monospace;
    font-size: 13px;
    color: var(--text-secondary);
  }}

  .stats-bar {{
    display: flex;
    gap: 24px;
    padding: 16px 40px;
    background: var(--bg-card);
    border-bottom: 1px solid var(--border);
  }}

  .stat-item {{
    display: flex;
    align-items: center;
    gap: 10px;
  }}

  .stat-dot {{
    width: 10px;
    height: 10px;
    border-radius: 50%;
  }}

  .stat-dot.green {{ background: var(--accent-green); box-shadow: 0 0 8px var(--accent-green); }}
  .stat-dot.red {{ background: var(--accent-red); box-shadow: 0 0 8px var(--accent-red); }}
  .stat-dot.gray {{ background: var(--text-muted); }}

  .stat-label {{
    font-size: 13px;
    color: var(--text-secondary);
  }}

  .stat-value {{
    font-family: 'JetBrains Mono', monospace;
    font-size: 18px;
    font-weight: 700;
  }}

  .stat-value.green {{ color: var(--accent-green); }}
  .stat-value.red {{ color: var(--accent-red); }}
  .stat-value.gray {{ color: var(--text-secondary); }}

  .container {{
    max-width: 1400px;
    margin: 0 auto;
    padding: 24px 24px 60px;
  }}

  .section {{
    margin-bottom: 32px;
  }}

  .section-header {{
    display: flex;
    align-items: center;
    gap: 12px;
    margin-bottom: 14px;
    padding: 0 4px;
  }}

  .section-icon {{
    font-size: 20px;
  }}

  .section-title {{
    font-size: 17px;
    font-weight: 700;
    letter-spacing: 0.5px;
  }}

  .section-count {{
    font-family: 'JetBrains Mono', monospace;
    font-size: 12px;
    padding: 3px 10px;
    border-radius: 12px;
    font-weight: 600;
  }}

  .section.buying .section-title {{ color: var(--accent-green); }}
  .section.buying .section-count {{
    background: var(--accent-green-bg);
    color: var(--accent-green);
    border: 1px solid rgba(16, 185, 129, 0.2);
  }}

  .section.selling .section-title {{ color: var(--accent-red); }}
  .section.selling .section-count {{
    background: var(--accent-red-bg);
    color: var(--accent-red);
    border: 1px solid rgba(239, 68, 68, 0.2);
  }}

  .section.nodata .section-title {{ color: var(--text-muted); }}
  .section.nodata .section-count {{
    background: rgba(74, 85, 104, 0.15);
    color: var(--text-muted);
    border: 1px solid rgba(74, 85, 104, 0.2);
  }}

  .table-wrap {{
    background: var(--bg-card);
    border: 1px solid var(--border);
    border-radius: 10px;
    overflow: hidden;
  }}

  table {{
    width: 100%;
    border-collapse: collapse;
    font-size: 13.5px;
  }}

  thead th {{
    background: var(--header-bg);
    color: var(--text-secondary);
    font-size: 12px;
    font-weight: 600;
    text-transform: uppercase;
    letter-spacing: 0.8px;
    padding: 12px 14px;
    text-align: right;
    border-bottom: 1px solid var(--border);
    white-space: nowrap;
    position: sticky;
    top: 0;
    z-index: 2;
  }}

  thead th:nth-child(1),
  thead th:nth-child(2),
  thead th:nth-child(3) {{
    text-align: left;
  }}

  tbody tr {{
    border-bottom: 1px solid rgba(30, 45, 61, 0.4);
    transition: background 0.15s;
  }}

  tbody tr:hover {{
    background: rgba(59, 130, 246, 0.04);
  }}

  td {{
    padding: 10px 14px;
    font-family: 'JetBrains Mono', monospace;
    font-size: 13px;
  }}

  .rank-cell {{
    text-align: center;
    color: var(--text-muted);
    font-weight: 600;
    width: 44px;
  }}

  .code-cell {{
    text-align: left;
    color: var(--accent-blue);
    font-weight: 600;
  }}

  .name-cell {{
    text-align: left;
    font-family: 'Noto Sans TC', sans-serif;
    font-weight: 500;
    color: var(--text-primary);
    min-width: 100px;
  }}

  .num-cell {{
    text-align: right;
    white-space: nowrap;
  }}

  .pos {{ color: var(--accent-red); }}
  .neg {{ color: var(--accent-green); }}
  .na {{ color: var(--text-muted); font-style: italic; }}

  .buy-highlight {{
    background: var(--accent-green-bg);
  }}
  .buy-highlight span {{
    color: var(--accent-green) !important;
    font-weight: 700;
  }}

  .sell-highlight {{
    background: var(--accent-red-bg);
  }}
  .sell-highlight span {{
    color: var(--accent-red) !important;
    font-weight: 700;
  }}

  .footer {{
    text-align: center;
    padding: 24px;
    color: var(--text-muted);
    font-size: 12px;
    border-top: 1px solid var(--border);
    margin-top: 40px;
  }}

  .empty-msg {{
    text-align: center;
    padding: 40px;
    color: var(--text-muted);
    font-size: 14px;
  }}

  /* 台股漲跌顏色: 漲=紅, 跌=綠 (符合台灣習慣) */
  /* 注意: 這裡的 pos/neg class 已對應台灣慣例 */
  /* pos (>0) = 紅色 (漲), neg (<0) = 綠色 (跌) */

  .data-source-bar {{
    background: var(--bg-card);
    border-bottom: 1px solid var(--border);
    padding: 12px 40px;
    display: flex;
    align-items: center;
    gap: 32px;
    flex-wrap: wrap;
  }}

  .source-item {{
    display: flex;
    align-items: center;
    gap: 8px;
  }}

  .source-label {{
    font-size: 12.5px;
    color: var(--text-secondary);
    font-weight: 500;
  }}

  .source-date {{
    font-family: 'JetBrains Mono', monospace;
    font-size: 13px;
    font-weight: 700;
    color: var(--accent-blue);
    background: rgba(59, 130, 246, 0.1);
    padding: 2px 10px;
    border-radius: 4px;
  }}

  .date-warning {{
    color: var(--accent-amber);
    font-size: 12.5px;
    font-weight: 600;
    background: rgba(245, 158, 11, 0.1);
    padding: 4px 14px;
    border-radius: 6px;
    border: 1px solid rgba(245, 158, 11, 0.25);
    margin-left: auto;
  }}
</style>
</head>
<body>

<div class="top-bar">
  <h1>📊 跌幅排行 vs 外資買賣超 Dashboard</h1>
  <div class="meta">
    更新時間: {now_str} ｜ 上市 5日+10日 跌幅合併 共 {total} 檔
  </div>
</div>

<div class="data-source-bar">
  <div class="source-item">
    <span class="source-label">📈 5日跌幅</span>
    <span class="source-date">{date_5d if date_5d else "N/A"}</span>
  </div>
  <div class="source-item">
    <span class="source-label">📉 10日跌幅</span>
    <span class="source-date">{date_10d if date_10d else "N/A"}</span>
  </div>
  <div class="source-item">
    <span class="source-label">🏦 外資買賣超</span>
    <span class="source-date">{date_foreign if date_foreign else "N/A"}</span>
  </div>
  {"" if (date_5d == date_foreign and date_5d) or (not date_5d and not date_foreign) else '<div class="date-warning">⚠ 注意：跌幅資料與外資資料日期不同步，比對結果可能有誤差！</div>'}
</div>

<div class="stats-bar">
  <div class="stat-item">
    <div class="stat-dot green"></div>
    <div>
      <div class="stat-label">外資逢低布局</div>
      <div class="stat-value green">{len(buying_list)}</div>
    </div>
  </div>
  <div class="stat-item">
    <div class="stat-dot red"></div>
    <div>
      <div class="stat-label">外資持續看空</div>
      <div class="stat-value red">{len(selling_list)}</div>
    </div>
  </div>
  <div class="stat-item">
    <div class="stat-dot gray"></div>
    <div>
      <div class="stat-label">無外資資料</div>
      <div class="stat-value gray">{len(no_data_list)}</div>
    </div>
  </div>
</div>

<div class="container">

  <!-- 🟢 外資逢低布局 -->
  <div class="section buying">
    <div class="section-header">
      <span class="section-icon">🟢</span>
      <span class="section-title">外資逢低布局</span>
      <span class="section-count">{len(buying_list)} 檔</span>
    </div>
    <div class="table-wrap">
      {"<table><thead><tr><th>原排名</th><th>代號</th><th>名稱</th><th>收盤價</th><th>5日漲跌</th><th>5日跌幅</th><th>10日漲跌</th><th>10日跌幅</th><th>成交量</th><th>外資買(張)</th><th>外資賣(張)</th><th>外資淨買賣</th></tr></thead><tbody>" + buying_rows + "</tbody></table>" if buying_list else '<div class="empty-msg">目前無跌幅股票被外資逢低買入</div>'}
    </div>
  </div>

  <!-- 🔴 外資持續看空 -->
  <div class="section selling">
    <div class="section-header">
      <span class="section-icon">🔴</span>
      <span class="section-title">外資持續看空</span>
      <span class="section-count">{len(selling_list)} 檔</span>
    </div>
    <div class="table-wrap">
      {"<table><thead><tr><th>原排名</th><th>代號</th><th>名稱</th><th>收盤價</th><th>5日漲跌</th><th>5日跌幅</th><th>10日漲跌</th><th>10日跌幅</th><th>成交量</th><th>外資買(張)</th><th>外資賣(張)</th><th>外資淨買賣</th></tr></thead><tbody>" + selling_rows + "</tbody></table>" if selling_list else '<div class="empty-msg">目前無跌幅股票被外資持續賣出</div>'}
    </div>
  </div>

  <!-- ⚪ 無外資資料 -->
  {"" if not no_data_list else '''
  <div class="section nodata">
    <div class="section-header">
      <span class="section-icon">⚪</span>
      <span class="section-title">無外資資料</span>
      <span class="section-count">''' + str(len(no_data_list)) + ''' 檔</span>
    </div>
    <div class="table-wrap">
      <table><thead><tr><th>原排名</th><th>代號</th><th>名稱</th><th>收盤價</th><th>5日漲跌</th><th>5日跌幅</th><th>10日漲跌</th><th>10日跌幅</th><th>成交量</th><th>外資買(張)</th><th>外資賣(張)</th><th>外資淨買賣</th></tr></thead><tbody>''' + nodata_rows + '''</tbody></table>
    </div>
  </div>
  '''}

</div>

<div class="footer">
  資料來源: 富邦 e-Broker (跌幅排行) ｜ TWSE 臺灣證券交易所 (外資買賣超)<br>
  ⚠ 本工具僅供參考，不構成任何投資建議。投資有風險，請自行判斷。
</div>

</body>
</html>"""

    return html


# ============================================================
# 主程式
# ============================================================
def main():
    print("=" * 60)
    print("  股票跌幅 vs 外資買賣超 比對 Dashboard")
    print("=" * 60)
    print()

    # Step 1: 抓富邦跌幅排行 (5日 + 10日)
    print("[1/3] 正在抓取富邦 e-Broker 跌幅排行...")
    stocks_5d, date_5d = fetch_fubon_ranking(FUBON_URL_5D, "5日跌幅排行")
    if not stocks_5d:
        print("❌ 無法取得5日跌幅排行資料，請檢查網路連線或網址是否有效")
        return

    time.sleep(REQUEST_DELAY)

    stocks_10d, date_10d = fetch_fubon_ranking(FUBON_URL_10D, "10日跌幅排行")

    # ---- 合併邏輯: 以5日為主，補入10日資料；只在10日的也加入 ----
    # 建立 5日 map
    five_day_map = {}
    for s in stocks_5d:
        five_day_map[s["code"]] = s

    # 建立 10日 map
    ten_day_map = {}
    for s in stocks_10d:
        ten_day_map[s["code"]] = {
            "ten_day_change": s["five_day_change"],  # 10日頁面欄位結構同5日
            "ten_day_pct": s["five_day_pct"],
            "ten_day_rank": s["rank"],
            # 保留完整資料，給「只在10日」的股票用
            "full": s,
        }

    # (A) 5日清單: 補入 10日欄位
    stocks = []
    for s in stocks_5d:
        td = ten_day_map.get(s["code"], {})
        s["ten_day_change"] = td.get("ten_day_change")
        s["ten_day_pct"] = td.get("ten_day_pct")
        stocks.append(s)

    # (B) 只在10日、不在5日的股票: 補入（5日欄位填 None）
    only_10d_count = 0
    for code, td in ten_day_map.items():
        if code not in five_day_map:
            orig = td["full"]
            stocks.append({
                "rank": orig["rank"],
                "code": orig["code"],
                "name": orig["name"],
                "close": orig["close"],
                "change": orig["change"],
                "change_pct": orig["change_pct"],
                "volume": orig["volume"],
                "five_day_change": None,
                "five_day_pct": None,
                "ten_day_change": td["ten_day_change"],
                "ten_day_pct": td["ten_day_pct"],
            })
            only_10d_count += 1

    print(f"   → 合併後共 {len(stocks)} 檔 (5日:{len(stocks_5d)}, 僅10日:{only_10d_count})")

    time.sleep(REQUEST_DELAY)

    # Step 2: 抓 TWSE 外資買賣超
    foreign_map, date_foreign = fetch_twse_foreign_data()

    # Step 3: 比對 + 產生 HTML
    buying, selling, nodata = merge_and_classify(stocks, foreign_map)

    html_content = generate_html(buying, selling, nodata,
                                 date_5d=date_5d, date_10d=date_10d,
                                 date_foreign=date_foreign)

    # 寫入檔案
    output_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), OUTPUT_HTML)
    with open(output_path, "w", encoding="utf-8") as f:
        f.write(html_content)

    print()
    print(f"✅ Dashboard 已產生: {output_path}")
    print()

    # 自動開啟瀏覽器
    try:
        webbrowser.open(f"file://{output_path}")
        print("🌐 已自動開啟瀏覽器")
    except Exception:
        print(f"📂 請手動開啟: {output_path}")

    print()
    print("=" * 60)
    print(f"  🟢 外資逢低布局: {len(buying)} 檔")
    print(f"  🔴 外資持續看空: {len(selling)} 檔")
    if nodata:
        print(f"  ⚪ 無外資資料:   {len(nodata)} 檔")
    print("=" * 60)


if __name__ == "__main__":
    main()
