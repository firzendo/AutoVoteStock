#!/usr/bin/env python
# coding: utf-8
import os
import time
import datetime
import logging
import unicodedata
from selenium.webdriver.common.by import By

logger = logging.getLogger(__name__)

# 欄位固定寬度（以半形字元為單位，中文佔 2）
_COL_WIDTHS = [6, 12, 12, 22, 10, 6, 12, 18, 12]


def _display_width(text: str) -> int:
    """計算字串的顯示寬度（中文/全形 = 2，其餘 = 1）"""
    w = 0
    for ch in text:
        ea = unicodedata.east_asian_width(ch)
        w += 2 if ea in ('W', 'F') else 1
    return w


def _pad(text: str, width: int) -> str:
    """依顯示寬度右補空白至指定寬度"""
    return text + " " * max(0, width - _display_width(text))


def _fmt_row(cells: list) -> str:
    parts = [_pad(str(c), _COL_WIDTHS[i]) for i, c in enumerate(cells)]
    return "  ".join(parts).rstrip()


class ReportGenerator:
    def __init__(self, driver, page_navigator, output_dir: str = "output"):
        self.driver = driver
        self.page_navigator = page_navigator
        self.output_dir = output_dir

    def _scan_all_companies_from_page(self, log_msg_func) -> list:
        """逐頁掃描列表，回傳所有公司的完整欄位清單
        
        頁面欄位結構：
          col[0] = 證券代號 + 公司簡稱  (e.g. "3686 達能")
          col[1] = 會議日期 + 投票起訖日 (e.g. "115/05/26 115/04/25~115/05/23")
          col[2] = 投票狀況             (e.g. "已投票")
          col[3] = 作業項目             (修改/查詢/撤銷...)
          col[4] = 符合eGift資格 + 開始領取日 (e.g. "Y\n115/06/05" 或空)
        """
        all_companies = []
        self.page_navigator.go_to_first_page()

        while True:
            rows = self.driver.find_elements(By.TAG_NAME, "tr")
            for row in rows[1:]:  # 跳過表頭
                try:
                    cols = row.find_elements(By.TAG_NAME, "td")
                    if len(cols) < 3:
                        continue

                    # col[0]: 證券代號 + 公司簡稱
                    first_col = cols[0].text.strip()
                    parts = first_col.split()
                    if not parts or not parts[0].isdigit():
                        continue  # 不是公司列
                    code = parts[0]
                    name = " ".join(parts[1:]) if len(parts) > 1 else "-"

                    # col[1]: 會議日期 + 投票起訖日
                    date_col = cols[1].text.strip() if len(cols) > 1 else ""
                    date_parts = date_col.split()
                    meeting_date = date_parts[0] if date_parts else "-"
                    vote_period  = date_parts[1] if len(date_parts) > 1 else "-"

                    # col[2]: 投票狀況
                    vote_status_text = cols[2].text.strip() if len(cols) > 2 else ""
                    if "已投票" in vote_status_text:
                        vote_status = "已投票"
                    elif "未投票" in vote_status_text:
                        vote_status = "未投票"
                    else:
                        vote_status = vote_status_text or "-"

                    # col[4]: 符合eGift資格 + 開始領取日
                    egift_qualify = "-"
                    receipt_date  = "-"
                    if len(cols) > 4:
                        egift_col = cols[4].text.strip()
                        egift_lines = [l.strip() for l in egift_col.splitlines() if l.strip()]
                        if egift_lines:
                            egift_qualify = egift_lines[0]           # "Y" 或其他
                            receipt_date  = egift_lines[1] if len(egift_lines) > 1 else "-"

                    all_companies.append({
                        'code':         code,
                        'name':         name,
                        'meeting_date': meeting_date,
                        'vote_period':  vote_period,
                        'vote_status':  vote_status,
                        'egift_qualify': egift_qualify,
                        'receipt_date': receipt_date,
                    })
                except Exception:
                    continue

            if self.page_navigator.go_to_next_page():
                time.sleep(1)
                continue
            else:
                break

        log_msg_func(f"ℹ️  名單掃描完畢，共 {len(all_companies)} 家公司")
        return all_companies

    def generate_voting_report(self, companies_info, screenshotted_companies, log_msg_func,
                               egift_skipped_companies=None, manual_skipped_companies=None):
        """生成投票結果報告：掃描完整名單，標記投票狀況與截圖狀況

        Args:
            companies_info: 本次投票的公司列表（補充來源）
                            [{'code': '2102', 'name': '泰豐', 'status': '已投票'}, ...]
            screenshotted_companies: 已截圖的公司代碼集合 {'2102', '2103', ...}
            log_msg_func: 日誌函數
            egift_skipped_companies: 因符合eGift資格而略過截圖的代碼集合（可選）
            manual_skipped_companies: .env SCREENSHOT_SKIP_LIST 手動跳過的代碼集合（可選）
        """
        egift_skipped = egift_skipped_companies or set()
        manual_skipped = manual_skipped_companies or set()
        if not os.path.exists(self.output_dir):
            os.makedirs(self.output_dir)

        year = datetime.datetime.now().strftime("%Y")
        report_file = os.path.join(self.output_dir, f"{year}_執行結果.txt")

        headers = [
            "證券代號",
            "公司簡稱",
            "會議日期",
            "投票起訖日",
            "投票狀況",
            "已截圖",
            "截圖日期",
            "符合eGift發放資格",
            "開始領取日"
        ]

        try:
            # 優先使用本次投票記錄（已在投票時記錄完整資訊，避免重新掃頁面）
            if companies_info:
                log_msg_func("ℹ️  使用本次投票紀錄生成報告")
                all_companies = companies_info
            else:
                # 降級方案：若無投票記錄則嘗試從頁面掃描
                log_msg_func("⚠️  無本次投票紀錄，嘗試掃描頁面...")
                all_companies = self._scan_all_companies_from_page(log_msg_func)
                if not all_companies:
                    log_msg_func("ℹ️  頁面掃描結果為空，報告無內容")
                    return

            # 建立本次新資料字典 {code: formatted_line}
            new_rows = {}
            for company in all_companies:
                code = company.get('code', '-')
                name = company.get('name', '-')

                # 投票狀況：優先用 status 欄位（已在投票時記錄），若無則用 vote_status（頁面掃描結果）
                status = company.get('status') or company.get('vote_status', '-')
                if status == '已投票':
                    vote_status = "✓ 已投票"
                elif status == '投票失敗':
                    vote_status = "✗ 失敗"
                elif status == '未投票':
                    vote_status = "✗ 未投票"
                else:
                    vote_status = "✓ 已投票" if status and '投票' not in str(status) else status

                if code in egift_skipped:
                    is_screenshotted = "eGift"
                elif code in manual_skipped:
                    is_screenshotted = "跳過"
                elif code in screenshotted_companies:
                    is_screenshotted = "✓"
                else:
                    is_screenshotted = "-"

                # 截圖日期：eGift/手動跳過無實際截圖，顯示 '-'
                if code in egift_skipped or code in manual_skipped:
                    screenshot_date = '-'
                elif isinstance(screenshotted_companies, dict):
                    screenshot_date = screenshotted_companies.get(code, '-')
                else:
                    screenshot_date = '-'

                row = [
                    code,
                    name,
                    company.get('meeting_date', '-'),
                    company.get('vote_period', '-'),
                    vote_status,
                    is_screenshotted,
                    screenshot_date,
                    company.get('egift_qualify', '-'),
                    company.get('receipt_date', '-'),
                ]
                new_rows[code] = _fmt_row(row)

            # 若檔案已存在，讀取現有內容並合併
            if os.path.exists(report_file):
                with open(report_file, 'r', encoding='utf-8') as f:
                    existing_lines = f.readlines()

                # 解析現有資料行（跳過表頭與分隔線）
                existing_data = {}   # code -> formatted_line
                existing_order = []  # 保持原始順序
                for line in existing_lines[2:]:
                    stripped = line.rstrip('\n')
                    if not stripped:
                        continue
                    parts = stripped.split()
                    if parts and parts[0].isdigit():
                        code = parts[0]
                        existing_data[code] = stripped
                        if code not in existing_order:
                            existing_order.append(code)

                # 合併：相同代號覆蓋，新代號附加在後
                for code, row_line in new_rows.items():
                    existing_data[code] = row_line
                    if code not in existing_order:
                        existing_order.append(code)

                with open(report_file, 'w', encoding='utf-8') as f:
                    f.write(_fmt_row(headers) + "\n")
                    f.write("-" * 100 + "\n")
                    for code in existing_order:
                        f.write(existing_data[code] + "\n")
            else:
                # 新建檔案
                with open(report_file, 'w', encoding='utf-8') as f:
                    f.write(_fmt_row(headers) + "\n")
                    f.write("-" * 100 + "\n")
                    for row_line in new_rows.values():
                        f.write(row_line + "\n")

            log_msg_func(f"✓ 報告已生成: {report_file}")
            logger.info("報告文件: %s", report_file)

        except Exception as e:
            log_msg_func(f"⚠️  報告生成失敗: {str(e)}")
            logger.error("報告生成失敗: %s", e)
