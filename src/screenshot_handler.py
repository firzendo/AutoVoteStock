#!/usr/bin/env python
# coding: utf-8
import os
import time
import datetime
import logging
from selenium import webdriver
from selenium.webdriver.common.by import By
from page_navigator import PageNavigator


class ScreenshotHandler:
    def __init__(self, driver: webdriver.Chrome, page_navigator: PageNavigator = None, screenshot_dir: str = "screenshots"):
        self.driver = driver
        self.screenshot_dir = screenshot_dir
        self.screenshotted_companies = self._load_screenshotted_from_disk()
        # 使用傳入的 page_navigator 或創建新的
        self.page_navigator = page_navigator
    
    def _load_screenshotted_from_disk(self) -> set:
        codes = set()
        if not os.path.exists(self.screenshot_dir):
            return codes
        for fname in os.listdir(self.screenshot_dir):
            if not fname.endswith('.png'):
                continue
            # 去掉 .png 後，從右邊找最後一個 '_' 後面的內容為代碼
            fname_no_ext = fname[:-4]
            last_idx = fname_no_ext.rfind('_')
            if last_idx != -1:
                code = fname_no_ext[last_idx + 1:]
                codes.add(code)
        return codes

    def save_error_screenshot(self, error_id: str = "") -> str:
        """發生錯誤時截圖並寫入 log。回傳截圖路徑（失敗時回傳空字串）"""
        try:
            if not os.path.exists(self.screenshot_dir):
                os.makedirs(self.screenshot_dir)
            ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            safe_id = error_id.replace(" ", "_")[:40] if error_id else "error"
            filename = f"ERROR_{ts}_{safe_id}.png"
            filepath = os.path.join(self.screenshot_dir, filename)
            self.driver.save_screenshot(filepath)
            logging.error(f"[ERROR截圖] id={safe_id} → {filepath}")
            print(f"   📸 錯誤截圖已儲存: {filename}")
            return filepath
        except Exception as exc:
            print(f"   ⚠️  錯誤截圖失敗: {exc!s:.60}")
            return ""

    def screenshot_all_companies_results(self, log_msg_func, screenshot_func):
        log_msg_func("\n【步驟3】逐一查詢各公司投票結果並截圖...")
        time.sleep(2)
        
        # 第一步：回到第一頁
        log_msg_func("\n📄 【分頁處理】返回到第一頁...")
        self.page_navigator.go_to_first_page()
        time.sleep(2)
        
        # 第二步：逐頁捲動並截圖
        while True:
            log_msg_func(f"\n📄 檢查當前頁的已投票公司...")
            index_take = 1  # 跳過表頭列（與 template 相同）
            had_screenshot_on_this_page = False
            
            rows = self.driver.find_elements(By.TAG_NAME, "tr")
            
            for index, row in enumerate(rows[index_take:]):
                try:
                    if "已投票" not in row.text:
                        continue
                    
                    # 解析代碼和名稱（第1欄）
                    try:
                        cols = row.find_elements(By.TAG_NAME, "td")
                        first_col_text = cols[0].text.strip()
                        parts = first_col_text.split()
                        company_code = parts[0] if parts else ""
                        company_name = " ".join(parts[1:]) if len(parts) > 1 else "未知"
                    except Exception:
                        company_code = "未知"
                        company_name = "未知"
                    
                    # 已截圖過則跳過
                    if company_code in self.screenshotted_companies:
                        log_msg_func(f"   ℹ️  {company_name} ({company_code}) 已截圖，跳過")
                        continue
                    
                    log_msg_func(f"\n   {company_name} ({company_code}) 查詢中...")
                    
                    # 參考 template：操作欄是 td[3]（根據 HTML 結構），找含「查詢」的 <a>
                    query_clicked = False
                    try:
                        op_cells = row.find_elements(By.TAG_NAME, "td")
                        if len(op_cells) > 3:  # 確保有操作欄
                            for link in op_cells[3].find_elements(By.TAG_NAME, "a"):
                                link_text = link.text.strip()
                                if "查詢" in link_text:
                                    link.click()
                                    query_clicked = True
                                    log_msg_func(f"   ✓ 已點擊查詢")
                                    break
                    except Exception as e:
                        log_msg_func(f"   ⚠️  查詢按鈕定位失敗: {str(e)[:60]}")
                    
                    if not query_clicked:
                        log_msg_func(f"   ⚠️  找不到查詢連結，跳過此公司")
                        continue
                    
                    time.sleep(3)  # 等待查詢結果頁面加載
                    
                    # 截圖
                    try:
                        screenshot_func(company_code, company_name)
                        self.screenshotted_companies.add(company_code)
                        log_msg_func(f"   ✓ 截圖完成")
                        had_screenshot_on_this_page = True
                    except Exception as e:
                        log_msg_func(f"   ⚠️  截圖失敗: {str(e)[:50]}")
                    
                    # 返回列表頁
                    time.sleep(1)
                    log_msg_func(f"   返回投票列表...")
                    
                    back_success = False
                    try:
                        # 方法1: 用 CSS 選擇器找返回按鈕
                        self.driver.execute_script(
                            "arguments[0].click();",
                            self.driver.find_element(By.CSS_SELECTOR, 'button[onclick="back(); return false;"]')
                        )
                        back_success = True
                    except Exception:
                        # 方法2: 找含「返回」文字的按鈕
                        try:
                            for by, val in [
                                (By.XPATH, '//a[contains(text(), "返回")]'),
                                (By.XPATH, '//button[contains(text(), "返回")]'),
                            ]:
                                btns = self.driver.find_elements(by, val)
                                for btn in btns:
                                    if btn.is_displayed():
                                        self.driver.execute_script("arguments[0].click();", btn)
                                        back_success = True
                                        break
                                if back_success:
                                    break
                        except Exception:
                            pass
                    
                    if not back_success:
                        # 方法3: 用瀏覽器返回
                        log_msg_func(f"   ⚠️  未找到返回按鈕，使用瀏覽器返回")
                        self.driver.back()
                    
                    time.sleep(2)  # 等待返回到列表頁
                
                except Exception as row_error:
                    # 此行處理失敗，繼續下一行
                    log_msg_func(f"   ⚠️  行處理失敗: {str(row_error)[:50]}")
                    continue
            
            # 檢查是否還有下一頁
            if not had_screenshot_on_this_page:
                log_msg_func(f"\n   ℹ️  當前頁無新的已投票公司需要截圖")
            
            # 嘗試翻到下一頁
            log_msg_func(f"\n📄 檢查是否有下一頁...")
            if self.page_navigator.go_to_next_page():
                log_msg_func(f"   ✓ 已翻到下一頁，繼續處理...")
                time.sleep(2)
                continue
            else:
                log_msg_func(f"   ℹ️  已是最後一頁，截圖完成")
                break
