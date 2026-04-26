#!/usr/bin/env python
# coding: utf-8
import logging
import time
from selenium import webdriver
from selenium.webdriver.common.by import By


class PageNavigator:
    _logger = logging.getLogger(__name__)

    def __init__(self, driver: webdriver.Chrome):
        self.driver = driver
    
    def go_to_first_page(self) -> bool:
        self._logger.info("📄 嘗試返回到第一頁...")
        time.sleep(1)
        
        # 尋找第一頁按鈕的多種可能選擇器
        first_page_selectors = [
            # 方法1: 數字"1"（當前頁通常是 <strong>1</strong>，非當前的是 <a>）
            (By.XPATH, '//span[@class="pagelinks"]//a[text()="1"]'),
            
            # 方法2: 「第一頁」按鈕
            (By.XPATH, '//img[@alt="第一頁"]'),
            (By.XPATH, '//a[contains(text(), "第一頁") or contains(text(), "第一页")]'),
            
            # 方法3: "首頁" 按鈕
            (By.XPATH, '//img[@alt="首頁"] | //img[@alt="首页"]'),
            
            # 方法4: 連結 stockInfo=1
            (By.XPATH, '//a[contains(@href, "stockInfo=1")]'),
        ]
        
        for by, selector in first_page_selectors:
            try:
                elements = self.driver.find_elements(by, selector)
                
                for idx, elem in enumerate(elements):
                    try:
                        # 檢查元素是否可見
                        if not elem.is_displayed():
                            continue
                        
                        # 檢查是否被禁用
                        if elem.get_attribute('disabled') or elem.get_attribute('aria-disabled') == 'true':
                            continue
                        
                        # 如果是 <strong> 標籤表示當前已在第一頁
                        if elem.tag_name == 'strong':
                            self._logger.info("ℹ️  已在第一頁")
                            return True
                        
                        elem_text = elem.text.strip() if elem.text else ""
                        elem_alt = elem.get_attribute('alt') or ""
                        
                        elem_info = elem_text or elem_alt or "按鈕"
                        
                        self._logger.info("✓ 找到第一頁按鈕: %s", elem_info)
                        
                        # 點擊
                        try:
                            elem.click()
                            time.sleep(3)
                            self._logger.info("✓ 已返回第一頁")
                            return True
                        except Exception as click_error:
                            # 嘗試 JavaScript 點擊
                            try:
                                self.driver.execute_script("arguments[0].click();", elem)
                                time.sleep(3)
                                self._logger.info("✓ 已使用 JavaScript 返回第一頁")
                                return True
                            except:
                                continue
                    
                    except Exception:
                        continue
            
            except Exception:
                continue
        
        self._logger.warning("⚠️  找不到第一頁按鈕，可能已在第一頁")
        return False
    
    def go_to_next_page(self) -> bool:
        self._logger.info("📄 檢查是否有下一頁...")
        time.sleep(1)
        
        # 尋找下一頁按鈕的多種可能選擇器
        next_page_selectors = [
            # 方法1: 圖片按鈕 - alt="下一頁"
            (By.XPATH, '//img[@alt="下一頁"]'),
            
            # 方法2: 帶 stockInfo 參數的連結（找最小的頁碼或包含增量的）
            (By.XPATH, '//a[contains(@href, "stockInfo=")]'),
            
            # 方法3: 純文本頁碼連結（2, 3 等）
            (By.XPATH, '//span[@class="pagelinks"]//a[not(@title="Go to page 1")]'),
            
            # 方法4: 按鈕元素
            (By.XPATH, '//button[contains(text(), "下一") or contains(text(), "下一頁")]'),
            
            # 方法5: 通用分頁框架
            (By.XPATH, '//a[@class="next-page"], //li[@class="next"]/a'),
        ]
        
        for by, selector in next_page_selectors:
            try:
                elements = self.driver.find_elements(by, selector)
                
                for idx, elem in enumerate(elements):
                    try:
                        # 檢查元素是否可見
                        if not elem.is_displayed():
                            self._logger.warning("⚠️  元素 [%s] 不可見，跳過", idx)
                            continue
                        
                        # 檢查是否被禁用
                        if elem.get_attribute('disabled') or elem.get_attribute('aria-disabled') == 'true':
                            self._logger.warning("⚠️  元素 [%s] 已禁用，跳過", idx)
                            continue
                        
                        # 獲取元素信息用於日誌
                        elem_text = elem.text.strip() if elem.text else ""
                        elem_href = elem.get_attribute('href') or ""
                        elem_alt = elem.get_attribute('alt') or ""
                        elem_title = elem.get_attribute('title') or ""
                        
                        elem_info = elem_text or elem_alt or elem_title or elem_href[:30]
                        
                        self._logger.info("✓ 找到下一頁按鈕 [%s]: %s", idx, elem_info)
                        
                        # 檢查是否是指向當前頁的連結（跳過）
                        if "stockInfo=1" in elem_href or elem_text == "1":
                            self._logger.warning("⚠️  這是當前頁連結，跳過")
                            continue
                        
                        # 嘗試點擊
                        self._logger.info("正在點擊下一頁按鈕...")
                        try:
                            elem.click()
                            time.sleep(3)  # 等待新頁面加載
                            self._logger.info("✓ 已翻到下一頁")
                            return True
                        except Exception as click_error:
                            error_type = str(type(click_error).__name__)
                            
                            # 如果是被遮擋，嘗試 JavaScript 點擊
                            if "not clickable" in str(click_error).lower() or "ElementClickInterceptedException" in error_type:
                                self._logger.warning("⚠️  元素被遮擋，嘗試 JavaScript 點擊...")
                                try:
                                    self.driver.execute_script("arguments[0].scrollIntoView(true);", elem)
                                    time.sleep(0.3)
                                    self.driver.execute_script("arguments[0].click();", elem)
                                    time.sleep(3)
                                    self._logger.info("✓ 已使用 JavaScript 翻到下一頁")
                                    return True
                                except Exception as js_error:
                                    self._logger.warning("⚠️  JavaScript 點擊失敗，嘗試下一個按鈕...")
                                    continue
                            elif "stale" in error_type.lower():
                                self._logger.warning("⚠️  元素已失效，嘗試下一個按鈕...")
                                continue
                            else:
                                self._logger.error("⚠️  點擊失敗: %s，嘗試下一個按鈕...", error_type)
                                continue
                    
                    except Exception as elem_error:
                        error_type = str(type(elem_error).__name__)
                        if "stale" in error_type.lower():
                            continue
                        else:
                            continue
            
            except Exception as selector_error:
                self._logger.warning("⚠️  選擇器錯誤，嘗試下一個...")
                continue
        
        self._logger.info("ℹ️  找不到下一頁按鈕（已在最後一頁或分頁已禁用）")
        return False
    
    def has_next_page(self) -> bool:
        self._logger.info("📄 檢測是否有下一頁按鈕...")
        
        # 尋找下一頁按鈕的選擇器
        next_page_selectors = [
            (By.XPATH, '//img[@alt="下一頁"][@style!="display: none"]'),
            (By.XPATH, '//a[contains(@href, "stockInfo=") and @style!="display: none"]'),
            (By.XPATH, '//span[@class="pagelinks"]//a[not(@title="Go to page 1")][@style!="display: none"]'),
        ]
        
        for by, selector in next_page_selectors:
            try:
                elements = self.driver.find_elements(by, selector)
                for elem in elements:
                    try:
                        if elem.is_displayed():
                            # 檢查是否被禁用
                            if not (elem.get_attribute('disabled') or elem.get_attribute('aria-disabled') == 'true'):
                                elem_text = elem.text.strip() if elem.text else ""
                                elem_href = elem.get_attribute('href') or ""
                                
                                # 跳過當前頁連結
                                if "stockInfo=1" in elem_href or elem_text == "1":
                                    continue
                                
                                self._logger.info("✓ 發現下一頁按鈕")
                                return True
                    except Exception:
                        continue
            except Exception:
                continue
        
        self._logger.info("✗ 未發現下一頁按鈕")
        return False
    
    def is_on_first_page(self) -> bool:
        try:
            # 檢查是否有 <strong>1</strong> 表示當前頁
            current_page_elem = self.driver.find_element(By.XPATH, '//span[@class="pagelinks"]//strong[text()="1"]')
            if current_page_elem.is_displayed():
                return True
        except Exception:
            pass
        
        return False
    
    def scroll_to_top(self) -> None:
        self.driver.execute_script("window.scrollTo(0, 0);")
        time.sleep(0.5)
    
    def get_device_pixel_ratio(self) -> float:
        try:
            ratio = self.driver.execute_script("return window.devicePixelRatio;")
            return ratio if ratio else 1.0
        except Exception:
            return 1.0
    
    def get_header_bottom_y(self) -> int:
        header_selectors = [
            'div[class="c-header_pageInfo"]',
            'header',
            '.c-header',
        ]
        
        for selector in header_selectors:
            try:
                el = self.driver.find_element(By.CSS_SELECTOR, selector)
                bottom = el.location["y"] + el.size["height"]
                if bottom > 0:
                    self._logger.info("✓ 找到 header，底部 Y=%s", bottom)
                    return bottom
            except Exception:
                continue
        
        self._logger.info("ℹ️  未找到 header，預設 Y=0")
        return 0
    
    def get_report_item_top_y(self) -> int:
        try:
            # 找含「報告事項」文字的元素
            els = self.driver.find_elements(By.XPATH, '//*[contains(text(), "報告事項")]')
            for el in els:
                if el.is_displayed():
                    top_y = el.location["y"]
                    self._logger.info("✓ 找到「報告事項」，頂部 Y=%s", top_y)
                    return top_y
        except Exception:
            pass
        
        self._logger.info("ℹ️  未找到「報告事項」")
        return None
    
    def find_all_unvoted_companies(self):
        self._logger.info("🔍 掃描未投票的公司...")
        time.sleep(1)  # 等待頁面加載
        
        # 尋試多個 XPATH 選擇器（避免使用無效的 CSS 選擇器）
        selectors = [
            (By.XPATH, '//tr[contains(., "未投票")]'),
            (By.XPATH, '//table//tr[contains(., "未投票")]'),
            (By.XPATH, '//tbody//tr[contains(., "未投票")]'),
            (By.XPATH, '//*[contains(text(), "未投票")]//ancestor::tr'),
        ]
        
        for selector_by, selector_value in selectors:
            try:
                rows = self.driver.find_elements(selector_by, selector_value)
                if rows:
                    self._logger.info("✓ 找到 %s 個未投票的公司", len(rows))
                    return rows
            except Exception as e:
                self._logger.warning("⚠️  選擇器失敗: %s", str(e)[:50])
                continue
        
        # 調試：顯示頁面內容
        self._logger.info("📋 頁面內容分析：")
        try:
            page_text = self.driver.find_element(By.TAG_NAME, 'body').text
            if "未投票" in page_text:
                self._logger.info("✓ 頁面包含「未投票」文字，但選擇器無法定位")
                self._logger.warning("⚠️  可能所有公司都已投票 (已投票|投票結果)")
            else:
                self._logger.info("✗ 頁面不包含「未投票」文字")
        except Exception:
            pass
        
        # 返回空列表而不是 raise exception，讓上層處理
        return []
    
    def find_unvoted_items(self):
        self._logger.info("🔍 掃描未投票的議案...")
        
        try:
            # 嘗試多種方式尋找未投票的議案
            selectors = [
                (By.XPATH, '//*[contains(text(), "未投票")]//ancestor::tr'),
                (By.XPATH, '//tr[contains(., "未投票")]'),
            ]
            
            for by, value in selectors:
                items = self.driver.find_elements(by, value)
                if items:
                    self._logger.info("✓ 找到 %s 個未投票的議案", len(items))
                    return items
            
            self._logger.info("ℹ️  嘗試備用方案...")
            return self.find_unvoted_items_fallback()
        
        except Exception as e:
            self._logger.error("❌ 查找未投票議案失敗: %s", str(e))
            return []
    
    def find_unvoted_items_fallback(self):
        try:
            items = []
            rows = self.driver.find_elements(By.XPATH, '//table//tr')
            
            for row in rows:
                vote_options = row.find_elements(By.XPATH, './/input[@type="radio"]')
                if vote_options:
                    checked = row.find_elements(By.XPATH, './/input[@type="radio"][@checked]')
                    if not checked:
                        items.append(row)
            
            self._logger.info("✓ 備用方案找到 %s 個未投票的議案", len(items))
            return items
        
        except Exception as e:
            self._logger.error("❌ 備用方案失敗: %s", str(e))
            return []
    
    def safe_click(self, element, element_name="按鈕", remove_disabled=False):
        try:
            # 嘗試1: 直接點擊
            element.click()
            time.sleep(1)
            return True
        except Exception as click_error:
            error_type = str(type(click_error).__name__)
            
            # 如果是被遮擋，嘗試 JavaScript 點擊
            if "not clickable" in str(click_error).lower() or "ElementClickInterceptedException" in error_type:
                try:
                    self._logger.warning("⚠️  %s被遮擋，嘗試 JavaScript 點擊...", element_name)
                    self.driver.execute_script("arguments[0].scrollIntoView(true);", element)
                    time.sleep(0.5)
                    
                    if remove_disabled:
                        self.driver.execute_script("arguments[0].removeAttribute('disabled');", element)
                        time.sleep(0.3)
                    
                    self.driver.execute_script("arguments[0].click();", element)
                    time.sleep(1)
                    return True
                except Exception as js_error:
                    self._logger.error("⚠️  JavaScript 點擊失敗: %s", str(type(js_error).__name__))
                    return False
            else:
                # 如果是陳舊元素或其他異常
                self._logger.error("⚠️  點擊失敗 (%s): %s", error_type, str(click_error)[:50])
                return False

    def click_vote_button(self):
        self._logger.info("=" * 60)
        self._logger.info("【投票流程】點擊投票按鈕")
        
        time.sleep(1)
        
        # 尋找投票按鈕的選擇器
        vote_button_selectors = [
            (By.XPATH, '//a[contains(text(), "投票")] | //button[contains(text(), "投票")]'),
            (By.XPATH, '//tr[contains(., "未投票")]//a[contains(text(), "投票")]'),
            (By.XPATH, '//tr[contains(., "未投票")]//button[contains(text(), "投票")]'),
            (By.CSS_SELECTOR, 'a.vote-link, button.vote-button'),
        ]
        
        for by, value in vote_button_selectors:
            buttons = self.driver.find_elements(by, value)
            self._logger.info("找到 %s 個投票按鈕", len(buttons))
            
            for button in buttons:
                if button.is_displayed():
                    self._logger.info("正在點擊投票按鈕...")
                    if self.safe_click(button, "投票按鈕"):
                        self._logger.info("✓ 已經點擊投票按鈕")
                        return (0, "已點擊投票按鈕")
        
        raise Exception("無法找到投票按鈕")

    def click_all_agree(self):
        self._logger.info("=" * 60)
        self._logger.info("【投票流程】點擊全部贊成")
        
        time.sleep(1)
        
        # 尋找「全部贊成」按鈕的選擇器（參考 template.py）
        agree_button_selectors = [
            # template.py 中的 JavaScript onclick 選擇器
            (By.CSS_SELECTOR, 'a[onclick="optionAll(0); return false;"]'),
            (By.CSS_SELECTOR, 'a[onclick="optionAll(0);return false;"]'),
            
            # 備選方案：文本匹配
            (By.XPATH, '//button[contains(text(), "全部贊成") or contains(text(), "全部承認")]'),
            (By.XPATH, '//a[contains(text(), "全部贊成") or contains(text(), "全部承認")]'),
            (By.XPATH, '//*[contains(text(), "全部贊成") or contains(text(), "全部承認")]'),
            
            # CSS 類或自訂屬性
            (By.CSS_SELECTOR, 'button.all-agree, button.agree-all, button.all-approve'),
        ]
        
        for by, value in agree_button_selectors:
            try:
                buttons = self.driver.find_elements(by, value)
                self._logger.info("找到 %s 個贊成按鈕", len(buttons))
                
                for button in buttons:
                    if button.is_displayed():
                        self._logger.info("正在點擊全部贊成按鈕...")
                        if self.safe_click(button, "全部贊成按鈕"):
                            self._logger.info("✓ 已點擊全部贊成按鈕")
                            return (0, "已點擊全部贊成")
            except Exception as e:
                self._logger.warning("⚠️  尋找全部贊成按鈕時出錯: %s", str(e)[:50])
                # 選擇器可能無效，繼續嘗試下一個
                continue
        
        raise Exception("無法找到全部贊成按鈕")

    def click_next_step(self):
        self._logger.info("【投票流程】點擊下一步...")
        time.sleep(1)
        
        # 尋找「下一步」按鈕的選擇器（參考 template.py）
        next_button_selectors = [
            # template.py 中的 JavaScript onclick 選擇器
            (By.CSS_SELECTOR, 'button[onclick="voteObj.ignoreVote();voteObj.goNext(); return false;"]'),
            (By.CSS_SELECTOR, 'button[onclick="voteObj.ignoreVote();voteObj.goNext();return false;"]'),
            
            # 備選方案：文本匹配
            (By.XPATH, '//button[contains(text(), "下一步")]'),
            (By.XPATH, '//a[contains(text(), "下一步")]'),
            (By.XPATH, '//button[contains(text(), "Next") or contains(text(), "next")]'),
            (By.CSS_SELECTOR, 'button.next, a.next, button[type="submit"]'),
            (By.XPATH, '//button[contains(text(), "確認") or contains(text(), "提交")]'),
        ]
        
        for by, value in next_button_selectors:
            buttons = self.driver.find_elements(by, value)
            for button in buttons:
                if button.is_displayed():
                    button_text = button.text.strip() or button.get_attribute('title')
                    self._logger.info("找到按鈕: %s", button_text)
                    if self.safe_click(button, "下一步按鈕"):
                        self._logger.info("✓ 已點擊下一步按鈕")
                        return (0, "已點擊下一步")
        
        raise Exception("無法找到下一步按鈕")

    def find_agree_button_for_director(self):
        try:
            self._logger.info("檢查是否有直接的『全部贊成』按鈕...")
            
            # 尋找「全部贊成」相關的選項（可能是按鈕、連結或其他元素）
            agree_selectors = [
                # 匹配包含「全部贊成」或「全部承認」的任何元素
                (By.XPATH, '//*[contains(text(), "全部贊成") or contains(text(), "全部承認")]'),
                # 匹配 a 標籤
                (By.XPATH, '//a[contains(text(), "全部贊成") or contains(text(), "全部承認")]'),
                # 匹配 button
                (By.XPATH, '//button[contains(text(), "全部贊成") or contains(text(), "全部承認")]'),
                # 匹配包含「贊成」的連結（處理不同的措詞）
                (By.XPATH, '//a[contains(., "贊成") and not(contains(., "反對"))]'),
                # 匹配任何包含「贊成」但不是「反對」的可點擊元素
                (By.XPATH, '//*[contains(text(), "贊成") and not(contains(text(), "反對")) and not(contains(text(), "棄權"))]'),
            ]
            
            for by, value in agree_selectors:
                try:
                    elements = self.driver.find_elements(by, value)
                    self._logger.info("[檢查] 尋找『全部贊成』... 找到 %s 個", len(elements))
                    
                    for elem in elements:
                        try:
                            if elem.is_displayed():
                                elem_text = elem.text.strip()
                                if not elem_text:
                                    elem_text = elem.get_attribute('title') or "按鈕"
                                
                                # 過濾掉「反對」和「棄權」
                                if "反對" not in elem_text and "棄權" not in elem_text:
                                    self._logger.info("✓ 找到『全部贊成』按鈕: %s", elem_text)
                                    
                                    if self.safe_click(elem, "全部贊成按鈕", remove_disabled=True):
                                        self._logger.info("✓ 已點擊『全部贊成』按鈕")
                                        return elem
                        except Exception as e2:
                            # 元素可能已過期，跳過
                            continue
                except Exception as e:
                    # 選擇器可能無效，跳過
                    continue
            
            self._logger.info("ℹ️  未找到直接的『全部贊成』按鈕")
            return None
        
        except Exception as e:
            self._logger.warning("⚠️  檢查同意按鈕時出錯: %s", str(e))
            return None

    def check_all_directors(self):
        try:
            self._logger.info("正在全部勾選...")
            
            # 尋找全部勾選按鈕
            check_all_selectors = [
                (By.XPATH, '//button[contains(text(), "全選")]'),
                (By.XPATH, '//button[contains(text(), "全部")]'),
                (By.XPATH, '//a[contains(text(), "全選")]'),
                (By.XPATH, '//*[contains(text(), "全部勾選")]'),
            ]
            
            for by, value in check_all_selectors:
                elements = self.driver.find_elements(by, value)
                self._logger.info("[選擇器] %s... 找到 %s 個", value[:50], len(elements))
                
                for elem in elements:
                    if elem.is_displayed():
                        elem_text = elem.text.strip() or elem.get_attribute('title') or "按鈕"
                        self._logger.info("✓ 找到「全選」按鈕: %s", elem_text)
                        
                        elem.click()
                        time.sleep(2)
                        self._logger.info("✓ 已點擊「全選」按鈕")
                        
                        all_checkboxes = self.driver.find_elements(By.XPATH, '//input[@type="checkbox"]')
                        checked_count = sum(1 for cb in all_checkboxes if cb.is_selected())
                        self._logger.info("✓ 驗證勾選: 已勾選 %s 個", checked_count)
                        return True
            
            # 備用方案：手動勾選所有 checkbox
            self._logger.info("未找到「全選」按鈕，嘗試手動勾選...")
            
            all_checkboxes = self.driver.find_elements(By.XPATH, '//table//input[@type="checkbox"]')
            if len(all_checkboxes) == 0:
                all_checkboxes = self.driver.find_elements(By.XPATH, '//input[@type="checkbox"]')
            
            self._logger.info("找到 %s 個複選框", len(all_checkboxes))
            
            checked_count = 0
            unchecked_count = 0
            for idx, checkbox in enumerate(all_checkboxes):
                if checkbox.is_displayed() and not checkbox.is_selected():
                    checkbox.click()
                    unchecked_count += 1
                    time.sleep(0.2)
                else:
                    if checkbox.is_selected():
                        checked_count += 1
            
            total = checked_count + unchecked_count
            self._logger.info("✓ 已勾選 %s 個複選框", total)
            time.sleep(3)
            
            return total > 0
        
        except Exception as e:
            self._logger.error("❌ 勾選複選框時出錯 (代碼:101): %s", str(e))
            return False

    def click_average_distribution(self):
        try:
            self._logger.info("正在平均分配...")
            
            all_checkboxes = self.driver.find_elements(By.XPATH, '//input[@type="checkbox"]')
            checked_count = sum(1 for cb in all_checkboxes if cb.is_selected())
            self._logger.info("當前已勾選的複選框: %s 個 (共 %s 個)", checked_count, len(all_checkboxes))
            
            # 尋找平均分配按鈕
            distribution_selectors = [
                (By.XPATH, '//button[contains(text(), "平均分配")]'),
                (By.XPATH, '//a[contains(text(), "平均分配")]'),
                (By.XPATH, '//*[contains(text(), "平均分配")]'),
                (By.XPATH, '//img[@title="平均分配"]'),
            ]
            
            for by, value in distribution_selectors:
                buttons = self.driver.find_elements(by, value)
                self._logger.info("[選擇器] %s... 找到 %s 個", str(by)[:30], len(buttons))
                
                for button in buttons:
                    if button.is_displayed():
                        button_text = button.text.strip() or button.get_attribute('value') or button.get_attribute('title') or "按鈕"
                        self._logger.info("找到: %s", button_text[:50])
                        
                        is_disabled = button.get_attribute('disabled')
                        if is_disabled:
                            self._logger.warning("⚠️  按鈕已禁用，跳過")
                            continue
                        
                        button.click()
                        time.sleep(2)
                        self._logger.info("✓ 已點擊平均分配")
                        return True
            
            self._logger.warning("⚠️  未找到平均分配按鈕")
            return False
        
        except Exception as e:
            self._logger.error("❌ 點擊平均分配時出錯 (代碼:102): %s", str(e))
            return False

    def find_agree_option(self, item):
        try:
            # 尋找 radio button
            agree_options = item.find_elements(By.XPATH, './/input[@type="radio"]')
            if agree_options:
                return agree_options[0]
            
            # 尋找按鈕
            buttons = item.find_elements(By.XPATH, './/button')
            for btn in buttons:
                if "同意" in btn.text or "Agree" in btn.text:
                    return btn
            
            if buttons:
                return buttons[0]
            
            return None
        
        except Exception as e:
            self._logger.warning("⚠️  尋找同意選項失敗: %s", str(e))
            return None

    def submit_vote(self):
        self._logger.info("【步驟】提交投票...")
        time.sleep(2)  # 等待頁面穩定
        
        # 尋找提交按鈕（參考 template.py）
        submit_buttons = [
            # template.py 中的 JavaScript onclick 選擇器
            (By.CSS_SELECTOR, 'button[onclick="doProcess();"]'),
            (By.CSS_SELECTOR, 'button[onclick="voteObj.checkMeetingPartner(); return false;"]'),
            (By.CSS_SELECTOR, 'button[onclick="voteObj.checkVote(); return false;"]'),
            
            # 備選方案：文本匹配
            (By.XPATH, '//button[contains(text(), "提交")]'),
            (By.XPATH, '//button[contains(text(), "確認")]'),
            (By.CSS_SELECTOR, 'button[type="submit"]'),
            (By.XPATH, '//input[@type="submit"]'),
        ]
        
        for by, value in submit_buttons:
            buttons = self.driver.find_elements(by, value)
            self._logger.info("→ 嘗試選擇器: %s..., 找到 %s 個", str(by)[:20], len(buttons))
            
            for idx, button in enumerate(buttons):
                try:
                    # 檢查是否可見
                    if not button.is_displayed():
                        self._logger.info("[按鈕 %s] 不可見，跳過", idx)
                        continue
                    
                    # 嘗試1: 直接點擊
                    self._logger.info("[按鈕 %s] 嘗試直接點擊...", idx)
                    if self.safe_click(button, f"提交按鈕[{idx}]", remove_disabled=True):
                        self._logger.info("✓ 已提交投票")
                        return (0, "已提交投票")
                
                except Exception as e:
                    continue
        
        self._logger.warning("⚠️  未找到可點擊的提交按鈕")
        raise Exception("無法找到或點擊提交按鈕")

    def click_query_button(self):
        self._logger.info("【步驟】點擊查詢按鈕...")
        time.sleep(2)  # 等待頁面完全穩定
        
        # 尋找查詢按鈕或投票確認按鈕
        query_button_selectors = [
            # 投票結果確認頁面的按鈕
            (By.XPATH, '//button[contains(text(), "確認投票結果")]'),
            (By.XPATH, '//a[contains(text(), "確認投票結果")]'),
            (By.CSS_SELECTOR, 'button[onclick="voteObj.checkVote(); return false;"]'),
            
            # 其他可能的確認按鈕
            (By.XPATH, '//button[contains(text(), "確認")]'),
            (By.CSS_SELECTOR, 'button[type="submit"]'),
        ]
        
        for by, value in query_button_selectors:
            try:
                # 每次新的選擇器都重新獲取元素，避免跨越多個查詢使用陳舊元素
                buttons = self.driver.find_elements(by, value)
                if not buttons:
                    continue
                
                for idx, button in enumerate(buttons):
                    try:
                        # 檢查是否可見（同時檢查元素仍然有效）
                        if not button.is_displayed():
                            continue
                        
                        button_text = button.text.strip() or button.get_attribute('title') or "按鈕"
                        self._logger.info("找到按鈕 [%s]: %s，正在點擊...", idx, button_text)
                        
                        if self.safe_click(button, button_text):
                            self._logger.info("✓ 已點擊按鈕「%s」", button_text)
                            return (0, f"已點擊按鈕: {button_text}")
                    
                    except Exception as elem_error:
                        # 檢查元素屬性時出錯（可能是 is_displayed 檢查）
                        error_type = str(type(elem_error).__name__)
                        if "stale" in error_type.lower() or "StaleElementReferenceException" in error_type:
                            self._logger.warning("⚠️  元素已失效，使用新的選擇器重試...")
                            break  # 嘗試下一個選擇器
                        else:
                            continue
            
            except Exception as selector_error:
                # 選擇器本身可能有問題，繼續嘗試下一個
                continue
        
        self._logger.warning("⚠️  所有選擇器均失敗，無法找到查詢或確認按鈕")
        raise Exception("無法找到查詢或確認按鈕")

    def click_query_button_in_table(self, company_code):
        self._logger.info("【步驟】點擊表格中 %s 的查詢按鈕...", company_code)
        
        try:
            tbody = self.driver.find_element(By.TAG_NAME, "tbody")
            rows = tbody.find_elements(By.TAG_NAME, "tr")
            
            for row in rows:
                try:
                    # 檢查第一列是否是目標公司代碼
                    first_col = row.find_element(By.TAG_NAME, "td")
                    if company_code in first_col.text:
                        cols = row.find_elements(By.TAG_NAME, "td")
                        
                        # 操作項目列是第5個 (索引4)
                        # 結構: 修改 查詢 撤銷
                        if len(cols) >= 5:
                            operations_col = cols[4]  # 操作項目
                            links = operations_col.find_elements(By.TAG_NAME, "a")
                            
                            # 查詢按鈕是第二個链接 (索引1)
                            # [0]=修改, [1]=查詢, [2]=撤銷
                            query_link = None
                            
                            if len(links) >= 2:
                                # 優先用索引方式
                                query_link = links[1]
                            else:
                                # 備選：找包含"查詢"文字的link
                                for link in links:
                                    if "查詢" in link.text:
                                        query_link = link
                                        break
                            
                            if query_link:
                                # 滾動到按鈕位置
                                self.driver.execute_script("arguments[0].scrollIntoView(true);", query_link)
                                time.sleep(0.5)
                                
                                if self.safe_click(query_link, f"表格查詢按鈕({company_code})"):
                                    self._logger.info("✓ 已點擊表格中 %s 的查詢按鈕", company_code)
                                    return (0, f"已點擊表格查詢: {company_code}")
                except Exception as e:
                    continue
            
            self._logger.error("❌ 無法找到公司 %s 的查詢按鈕", company_code)
            return (-1, f"無法找到公司 {company_code} 的查詢按鈕")
            
        except Exception as e:
            self._logger.error("❌ 表格查詢失敗: %s", str(e))
            return (-1, f"表格查詢失敗: {str(e)}")

    def go_back_to_list(self):
        self._logger.info("【步驟】返回投票列表...")
        
        # 方式1: 點擊返回/返回投票列表按鈕
        back_buttons = [
            (By.XPATH, '//a[contains(text(), "返回投票列表")]'),
            (By.XPATH, '//button[contains(text(), "返回投票列表")]'),
            (By.XPATH, '//a[contains(text(), "返回")]'),
            (By.XPATH, '//button[contains(text(), "返回")]'),
            (By.CSS_SELECTOR, 'a[href*="back"], a[href*="list"], button[onclick*="back"]'),
        ]
        
        for by, value in back_buttons:
            try:
                buttons = self.driver.find_elements(by, value)
                for button in buttons:
                    if button.is_displayed():
                        button_text = button.text.strip() or button.get_attribute('title') or "返回按鈕"
                        self._logger.info("找到按鈕: %s", button_text)
                        
                        if self.safe_click(button, button_text):
                            self._logger.info("✓ 已返回列表")
                            return (True, "已返回列表")
            except Exception as e:
                # 選擇器可能無效，繼續嘗試
                continue
        
        # 方式2: 使用瀏覽器返回
        self._logger.info("📍 未發現返回按鈕，使用瀏覽器返回功能...")
        try:
            self.driver.back()
            time.sleep(2)
            self._logger.info("✓ 已使用瀏覽器返回")
            return (True, "已使用瀏覽器返回")
        except Exception as e:
            self._logger.error("❌ 瀏覽器返回失敗: %s", str(e))
            return (False, f"瀏覽器返回失敗: {str(e)}")
