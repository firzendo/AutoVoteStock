#!/usr/bin/env python
# coding: utf-8
import os
import datetime
import logging
import unicodedata
from selenium.webdriver.common.by import By

logger = logging.getLogger(__name__)

# 欄位固定寬度（以半形字元為單位，中文佔 2）
_COL_WIDTHS = [6, 12, 12, 22, 10, 6, 18, 12]


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
                import time
                time.sleep(1)
                continue
            else:
                break

        log_msg_func(f"ℹ️  名單掃描完畢，共 {len(all_companies)} 家公司")
        return all_companies

    def generate_voting_report(self, companies_info, screenshotted_companies, log_msg_func):
        """生成投票結果報告：掃描完整名單，標記投票狀況與截圖狀況

        Args:
            companies_info: 本次投票的公司列表（補充來源）
                            [{'code': '2102', 'name': '泰豐', 'status': '已投票'}, ...]
            screenshotted_companies: 已截圖的公司代碼集合 {'2102', '2103', ...}
            log_msg_func: 日誌函數
        """
        if not os.path.exists(self.output_dir):
            os.makedirs(self.output_dir)

        timestamp = datetime.datetime.now().strftime("%Y%m%d")
        report_file = os.path.join(self.output_dir, f"{timestamp}_執行結果.txt")

        try:
            # 以本次投票記錄建立快速查詢表（code → status）
            voted_this_run = {
                c['code']: c['status']
                for c in companies_info
                if c.get('code') and c['code'] != '未知'
            }

            # 從頁面掃描完整名單
            all_companies = self._scan_all_companies_from_page(log_msg_func)

            # 若頁面掃描失敗（空），退而使用 companies_info
            if not all_companies:
                log_msg_func("⚠️  頁面掃描結果為空，改用本次投票記錄")
                all_companies = [
                    {
                        'code': c.get('code', '-'),
                        'name': c.get('name', '-'),
                        'vote_status': c.get('status', '-'),
                    }
                    for c in companies_info
                ]

            with open(report_file, 'w', encoding='utf-8') as f:
                headers = [
                    "證券代號",
                    "公司簡稱",
                    "會議日期",
                    "投票起訖日",
                    "投票狀況",
                    "已截圖",
                    "符合eGift發放資格",
                    "開始領取日"
                ]
                f.write(_fmt_row(headers) + "\n")
                f.write("-" * 100 + "\n")

                for company in all_companies:
                    code = company.get('code', '-')
                    name = company.get('name', '-')

                    # 投票狀況：優先用頁面顯示，若為未知則補充本次記錄
                    page_status = company.get('vote_status', '-')
                    run_status  = voted_this_run.get(code, '')
                    if page_status == '已投票':
                        vote_status = "✓ 已投票"
                    elif run_status == '已投票':
                        vote_status = "✓ 已投票"
                    elif run_status == '投票失敗':
                        vote_status = "✗ 失敗"
                    elif page_status == '未投票':
                        vote_status = "✗ 未投票"
                    else:
                        vote_status = page_status if page_status != '-' else "-"

                    is_screenshotted = "✓" if code in screenshotted_companies else "-"

                    row = [
                        code,
                        name,
                        company.get('meeting_date', '-'),
                        company.get('vote_period', '-'),
                        vote_status,
                        is_screenshotted,
                        company.get('egift_qualify', '-'),
                        company.get('receipt_date', '-'),
                    ]
                    f.write(_fmt_row(row) + "\n")

            log_msg_func(f"✓ 報告已生成: {report_file}")
            logger.info("報告文件: %s", report_file)

        except Exception as e:
            log_msg_func(f"⚠️  報告生成失敗: {str(e)}")
            logger.error("報告生成失敗: %s", e)

            log_msg_func(f"✓ 報告已生成: {report_file}")
            logger.info("報告文件: %s", report_file)
            
        except Exception as e:
            log_msg_func(f"⚠️  報告生成失敗: {str(e)}")
            logger.error("報告生成失敗: %s", e)
