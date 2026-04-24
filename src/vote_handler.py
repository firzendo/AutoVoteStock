#!/usr/bin/env python
# coding: utf-8
import os
import time
from selenium import webdriver
from selenium.webdriver.common.by import By


class VoteHandler:
    """投票處理器"""
    
    def __init__(self, driver: webdriver.Chrome, screenshot_dir: str = "screenshots"):
        """
        初始化投票處理器
        
        Args:
            driver: Selenium WebDriver 實例
            screenshot_dir: 截圖目錄，用於偵測已截圖的公司
        """
        self.driver = driver
        self.vote_timeout = 30
        self.voted_count = 0
        self.total_count = 0
        self.screenshot_dir = screenshot_dir
        self.screenshotted_companies = self._load_screenshotted_from_disk()
    
    def _load_screenshotted_from_disk(self) -> set:
        """
        掃描截圖目錄，回傳已截圖過的公司代碼集合
        檔名格式: {date}_{company_name}_{company_code}.png
        """
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

    def _click_vote_button(self):
        """
        點擊投票按鈕進入投票頁面
        
        Returns:
            tuple: (code, msg) - 0成功，<0為錯誤代碼
        """
        print("\n" + "=" * 60)
        print("【投票流程】點擊投票按鈕")
        print("=" * 60)
        
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
            print(f"   找到 {len(buttons)} 個投票按鈕")
            
            for button in buttons:
                if button.is_displayed():
                    print(f"   正在點擊投票按鈕...")
                    try:
                        # 嘗試1: 直接點擊
                        button.click()
                        time.sleep(2)
                        print(f"   ✓ 已經點擊投票按鈕")
                        return (0, "已點擊投票按鈕")
                    except Exception as e:
                        # 嘗試2: 滾動到元素，再用 JavaScript 點擊
                        if "ElementClickInterceptedException" in str(type(e).__name__) or "not clickable" in str(e):
                            print(f"   ⚠️  元素被遮擋，嘗試 JavaScript 點擊...")
                            self.driver.execute_script("arguments[0].scrollIntoView(true);", button)
                            time.sleep(1)
                            self.driver.execute_script("arguments[0].click();", button)
                            time.sleep(2)
                            print(f"   ✓ 已使用 JavaScript 點擊投票按鈕")
                            return (0, "已使用 JavaScript 點擊投票按鈕")
                        else:
                            raise
        
        raise Exception("無法找到投票按鈕")  # 異常由上層捕獲
    
    def _print_vote_button_debug_info(self):
        """打印投票按鈕的調試信息"""
        try:
            print("\n📋 ===== 投票按鈕調試信息 =====")
            
            # 列出所有 link
            links = self.driver.find_elements(By.TAG_NAME, 'a')
            print(f"\n🔗 發現 {len(links)} 個連結 (前10個):")
            vote_links = []
            for idx, link in enumerate(links[:20]):
                text = link.text.strip()
                href = link.get_attribute('href') or ''
                if '投票' in text or 'vote' in href.lower():
                    print(f"   [{idx}] ✓ {text} -> {href[:50]}")
                    vote_links.append(link)
                elif text:
                    print(f"   [{idx}] {text}")
            
            if vote_links:
                print(f"\n💡 找到 {len(vote_links)} 個可能的投票連結，建議點擊第一個")
                print("   或在開發者工具中查看具體的 HTML 結構")
            
            print("\n💡 建議:")
            print("   1. 打開開發者工具 (F12)")
            print("   2. 查看投票按鈕的 HTML 代碼")
            print("   3. 查看是 <a> 還是 <button>")
            print("   4. 或手動點擊投票按鈕後按 Enter 繼續\n")
            print("=" * 40)
        
        except Exception as e:
            print(f"調試信息讀取失敗: {str(e)}")
    
    def _click_all_agree(self):
        """
        點擊「全部贊成(承認)」按鈕
        
        Returns:
            tuple: (code, msg) - 0成功，<0為錯誤代碼
        """
        print("\n" + "=" * 60)
        print("【投票流程】點擊全部贊成")
        print("=" * 60)
        
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
                print(f"   找到 {len(buttons)} 個贊成按鈕")
                
                for button in buttons:
                    if button.is_displayed():
                        print(f"   正在點擊全部贊成按鈕...")
                        try:
                            button.click()
                            time.sleep(1)
                            print(f"   ✓ 已點擊全部贊成按鈕")
                            return (0, "已點擊全部贊成")
                        except Exception as e:
                            if "not clickable" in str(e) or "ElementClickInterceptedException" in str(type(e).__name__):
                                print(f"   ⚠️  元素被遮擋，嘗試 JavaScript 點擊...")
                                self.driver.execute_script("arguments[0].scrollIntoView(true);", button)
                                time.sleep(0.5)
                                self.driver.execute_script("arguments[0].click();", button)
                                time.sleep(1)
                                print(f"   ✓ 已使用 JavaScript 點擊")
                                return (0, "已使用 JavaScript 點擊全部贊成")
                            else:
                                raise
            except Exception as e:
                # 選擇器可能無效，繼續嘗試下一個
                continue
        
        raise Exception("無法找到全部贊成按鈕")
    
    def _click_next_step(self):
        """
        點擊「下一步」按鈕
        
        Returns:
            tuple: (code, msg) - 0成功，<0為錯誤代碼
        """
        print("\n【投票流程】點擊下一步...")
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
                    print(f"   找到按鈕: {button_text}")
                    try:
                        button.click()
                        time.sleep(2)
                        print(f"   ✓ 已點擊下一步按鈕")
                        return (0, "已點擊下一步")
                    except Exception as e:
                        if "not clickable" in str(e):
                            print(f"   ⚠️  元素被遮擋，嘗試 JavaScript 點擊...")
                            self.driver.execute_script("arguments[0].scrollIntoView(true);", button)
                            time.sleep(0.5)
                            self.driver.execute_script("arguments[0].click();", button)
                            time.sleep(2)
                            print(f"   ✓ 已使用 JavaScript 點擊")
                            return (0, "已使用 JavaScript 點擊下一步")
                        else:
                            raise
        
        raise Exception("無法找到下一步按鈕")
    
    def _handle_director_election(self):
        """
        處理董事選舉 - 全部勾選 + 平均分配 + 下一步
        
        Returns:
            bool: 成功返回True
        """
        try:
            print("\n" + "=" * 60)
            print("【投票流程】董事選舉 - 全部勾選 + 平均分配")
            print("=" * 60)
            
            time.sleep(1)
            
            # 檢查是否是董事選舉頁面
            page_html = self.driver.page_source
            if "董事" not in page_html:
                print("   ⚠️  未檢測到董事選舉頁面")
                return False
            
            print("   ✓ 檢測到董事選舉頁面")
            
            # 新規則：檢查是否有「全部贊成」等按鈕（不需要勾選複選框）
            agree_button = self._find_agree_button_for_director()
            if agree_button:
                print("   ✓ 已點擊『全部贊成』按鈕，準備進入下一步...")
                # 重要：點擊「全部贊成」後也要執行「下一步」來推進流程
                time.sleep(2)
                try:
                    code, msg = self._click_next_step()
                    if code == 0:
                        print("   ✓ 已點擊下一步，董事選舉完成，進入下一個議案")
                        time.sleep(2)
                        return True  # 成功完成董事選舉並進入下一個議案
                    else:
                        print(f"   ⚠️  點擊下一步失敗: {msg}")
                        return True  # 仍然返回True，讓主流程繼續
                except Exception as e:
                    print(f"   ⚠️  無法點擊下一步: {str(e)[:50]}")
                    return True  # 即使下一步失敗也返回True，避免重複
            
            # 步驟1: 全部勾選複選框
            if self._check_all_directors():
                print("   ✓ 已全部勾選")
            else:
                print("   ⚠️  未能成功全部勾選")
                print("   ⚠️  等待用戶手動勾選...")
                input("\n⏳ 請手動勾選所有董事，然後按 Enter 繼續...\n")
            
            # 確保勾選已完成，等待足夠長時間
            print("   ⏳ 等待勾選狀態確認（3秒）...")
            time.sleep(3)
            
            # 步驟2: 平均分配
            print("\n   準備點擊平均分配按鈕...")
            if self._click_average_distribution():
                print("   ✓ 已點擊平均分配")
            else:
                print("   ⚠️  未能點擊平均分配")
                print("   ⚠️  等待用戶手動操作...")
                input("\n⏳ 請手動點擊平均分配，然後按 Enter 繼續...\n")
            
            time.sleep(2)
            
            # 步驟3: 點擊下一步
            try:
                code, msg = self._click_next_step()
                if code == 0:
                    print("   ✓ 已點擊下一步")
                    time.sleep(2)
                    
                    # 步驟4: 檢查下一個頁面是否還有其他投票項目
                    print("\n   ⏳ 等待新頁面加載...")
                    time.sleep(2)
                    
                    page_html = self.driver.page_source
                    if "全部贊成" in page_html or "全部反對" in page_html or "全部棄權" in page_html:
                        print("   ✓ 檢測到更多投票項目在此頁面")
                        print("   ℹ️  需要在此頁面投票其他議案")
                        return True  # 返回 True 表示已處理董事選舉，主流程需繼續處理其他議案
                    
                    return True
                else:
                    print("   ⚠️  無法點擊下一步")
                    print("   ⚠️  等待用戶手動操作...")
                    input("\n⏳ 請手動點擊下一步，然後按 Enter 繼續...\n")
                    return True  # 即使自動失敗，也返回True（用戶已手動操作）
            except Exception as e:
                print(f"   ❌ 點擊下一步失敗: {str(e)[:50]}")
                return True  # 返回True避免無限循環
        
        except Exception as e:
            print(f"   ❌ 處理董事選舉時出錯: {str(e)}")
            return False
    
    def _find_agree_button_for_director(self):
        """
        檢查董事選舉頁面是否有「全部贊成」等直接按鈕（不需要勾選複選框）
        
        Returns:
            WebElement or None: 找到按鈕返回元素，否則返回None
        """
        try:
            print("\n   檢查是否有直接的『全部贊成』按鈕...")
            
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
                    print(f"   [檢查] 尋找『全部贊成』... 找到 {len(elements)} 個")
                    
                    for elem in elements:
                        try:
                            if elem.is_displayed():
                                elem_text = elem.text.strip()
                                if not elem_text:
                                    elem_text = elem.get_attribute('title') or "按鈕"
                                
                                # 過濾掉「反對」和「棄權」
                                if "反對" not in elem_text and "棄權" not in elem_text:
                                    print(f"   ✓ 找到『全部贊成』按鈕: {elem_text}")
                                    
                                    try:
                                        elem.click()
                                        time.sleep(2)
                                        print(f"   ✓ 已點擊『全部贊成』按鈕")
                                        return elem
                                    except Exception as e:
                                        print(f"   ⚠️  直接點擊失敗，嘗試 JavaScript 點擊...")
                                        self.driver.execute_script("arguments[0].scrollIntoView(true);", elem)
                                        time.sleep(0.5)
                                        self.driver.execute_script("arguments[0].removeAttribute('disabled');", elem)
                                        time.sleep(0.3)
                                        self.driver.execute_script("arguments[0].click();", elem)
                                        time.sleep(2)
                                        print(f"   ✓ 已使用 JavaScript 點擊『全部贊成』")
                                        return elem
                        except Exception as e2:
                            # 元素可能已過期，跳過
                            continue
                except Exception as e:
                    # 選擇器可能無效，跳過
                    continue
            
            print("   ℹ️  未找到直接的『全部贊成』按鈕")
            return None
        
        except Exception as e:
            print(f"   ⚠️  檢查同意按鈕時出錯: {str(e)}")
            return None
    
    def _check_all_directors(self):
        """
        全部勾選董事複選框
        
        Returns:
            bool: 成功返回True
        """
        try:
            print("\n   正在全部勾選...")
            
            # 尋找全部勾選按鈕
            check_all_selectors = [
                (By.XPATH, '//button[contains(text(), "全選")]'),
                (By.XPATH, '//button[contains(text(), "全部")]'),
                (By.XPATH, '//a[contains(text(), "全選")]'),
                (By.XPATH, '//*[contains(text(), "全部勾選")]'),
            ]
            
            for by, value in check_all_selectors:
                elements = self.driver.find_elements(by, value)
                print(f"   [選擇器] {value[:50]}... 找到 {len(elements)} 個")
                
                for elem in elements:
                    if elem.is_displayed():
                        elem_text = elem.text.strip() or elem.get_attribute('title') or "按鈕"
                        print(f"   ✓ 找到「全選」按鈕: {elem_text}")
                        
                        elem.click()
                        time.sleep(2)
                        print(f"   ✓ 已點擊「全選」按鈕")
                        
                        all_checkboxes = self.driver.find_elements(By.XPATH, '//input[@type="checkbox"]')
                        checked_count = sum(1 for cb in all_checkboxes if cb.is_selected())
                        print(f"   ✓ 驗證勾選: 已勾選 {checked_count} 個")
                        return True
            
            # 備用方案：手動勾選所有 checkbox
            print("   未找到「全選」按鈕，嘗試手動勾選...")
            
            all_checkboxes = self.driver.find_elements(By.XPATH, '//table//input[@type="checkbox"]')
            if len(all_checkboxes) == 0:
                all_checkboxes = self.driver.find_elements(By.XPATH, '//input[@type="checkbox"]')
            
            print(f"   找到 {len(all_checkboxes)} 個複選框")
            
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
            print(f"   ✓ 已勾選 {total} 個複選框")
            time.sleep(3)
            
            return total > 0
        
        except Exception as e:
            print(f"   ❌ 勾選複選框時出錯 (代碼:101): {str(e)}")
            return False
    
    def _click_average_distribution(self):
        """
        點擊平均分配按鈕
        
        Returns:
            bool: 成功返回True
        """
        try:
            print("   正在平均分配...")
            
            all_checkboxes = self.driver.find_elements(By.XPATH, '//input[@type="checkbox"]')
            checked_count = sum(1 for cb in all_checkboxes if cb.is_selected())
            print(f"   當前已勾選的複選框: {checked_count} 個 (共 {len(all_checkboxes)} 個)")
            
            # 尋找平均分配按鈕
            distribution_selectors = [
                (By.XPATH, '//button[contains(text(), "平均分配")]'),
                (By.XPATH, '//a[contains(text(), "平均分配")]'),
                (By.XPATH, '//*[contains(text(), "平均分配")]'),
                (By.XPATH, '//img[@title="平均分配"]'),
            ]
            
            for by, value in distribution_selectors:
                buttons = self.driver.find_elements(by, value)
                print(f"   [選擇器] {str(by)[:30]}... 找到 {len(buttons)} 個")
                
                for button in buttons:
                    if button.is_displayed():
                        button_text = button.text.strip() or button.get_attribute('value') or button.get_attribute('title') or "按鈕"
                        print(f"   找到: {button_text[:50]}")
                        
                        is_disabled = button.get_attribute('disabled')
                        if is_disabled:
                            print(f"   ⚠️  按鈕已禁用，跳過")
                            continue
                        
                        button.click()
                        time.sleep(2)
                        print(f"   ✓ 已點擊平均分配")
                        return True
            
            print("   ⚠️  未找到平均分配按鈕")
            return False
        
        except Exception as e:
            print(f"   ❌ 點擊平均分配時出錯 (代碼:102): {str(e)}")
            return False
    
    def _vote_with_agree_button(self):
        """
        使用全部贊成按鈕的投票流程（自動偵測董事選舉和多個議案）
        
        Returns:
            dict: 投票統計
        """
        try:
            print("\n" + "=" * 60)
            print("【投票流程】快速投票")
            print("=" * 60)
            
            # 增加等待時間讓投票頁面完全加載
            time.sleep(3)
            
            page_source = self.driver.page_source
            page_text = page_source.lower()
            
            # 診斷：如果還是列表頁面，進行額外等待
            if "未投票" in page_text and "投票狀況" in page_text and "全部贊成" not in page_text:
                print("   ⚠️  似乎還在列表頁面，進行額外等待...")
                time.sleep(3)
                page_source = self.driver.page_source
                page_text = page_source.lower()
            
            iteration = 0
            max_iterations = 10  # 防止無限迴圈
            last_page_text = ""  # 用於檢測重複頁面
            repeat_count = 0  # 重複計數
            
            while iteration < max_iterations:
                iteration += 1
                print(f"\n【投票流程 - 循環 {iteration}】")
                time.sleep(1)  # 每個循環增加等待時間
                
                # 檢查頁面內容
                page_html = self.driver.page_source
                page_text = self.driver.find_element(By.TAG_NAME, 'body').text
                
                # 調試信息
                print(f"   □ 頁面文字長度: {len(page_text)} 字符")
                if "董事" in page_html:
                    print(f"   ✓ 頁面包含『董事』")
                if "全部贊成" in page_html:
                    print(f"   ✓ 頁面包含『全部贊成』")
                if "下一步" in page_html:
                    print(f"   ✓ 頁面包含『下一步』")
                
                # 檢測重複頁面（避免無限循環）
                if last_page_text and len(page_text) == len(last_page_text) and page_text[:100] == last_page_text[:100]:
                    repeat_count += 1
                    print(f"   ⚠️  偵測到重複頁面 (第 {repeat_count} 次)")
                    if repeat_count >= 3:
                        print("   ⚠️  已檢測到重複頁面多次，跳出循環避免無限迴圈")
                        return {'total': 1, 'voted': 1, 'failed': 0}
                else:
                    repeat_count = 0
                
                last_page_text = page_text
                
                # **第一優先：檢測是否是投票結果確認頁面** - 此時投票流程已完成
                if "投票內容確認" in page_html or "投票結果" in page_html:
                    print("   ✓ 檢測到『投票內容確認』或『投票結果』頁面")
                    print("   ✓ 投票流程已完成，準備提交...")
                    return {'total': 1, 'voted': 1, 'failed': 0}
                
                # 檢測是否是董事選舉頁面
                if "董事" in page_html:
                    # 再次確認：不是結果確認頁面的董事選舉
                    if "投票內容確認" not in page_html and "投票結果" not in page_html:
                        print("   ✓ 檢測到董事選舉頁面（投票操作）")
                        if self._handle_director_election():
                            print("   ✓ 董事選舉已完成，準備進入下一個議案...")
                            time.sleep(2)
                            continue  # 繼續檢查下一個頁面/議案
                        else:
                            print("   ⚠️  董事選舉處理失敗")
                            return {'total': 1, 'voted': 1, 'failed': 0}
                    else:
                        print("   ✓ 檢測到結果確認頁面中的董事資訊，投票已完成")
                        return {'total': 1, 'voted': 1, 'failed': 0}
                
                # 檢測是否有標準投票選項（全部贊成、全部反對、全部棄權）
                has_vote_options = "全部贊成" in page_html or "全部反對" in page_html or "全部棄權" in page_html
                
                if has_vote_options:
                    print("   ✓ 檢測到投票選項")
                    
                    # 點擊全部贊成
                    try:
                        code, msg = self._click_all_agree()
                        if code == 0:
                            print(f"   ✓ {msg}")
                    except Exception as e:
                        print(f"   ⚠️  無法自動點擊全部贊成: {str(e)[:50]}")
                        print("   請手動操作後按 Enter 繼續")
                        input("\n按 Enter 鍵繼續...\n")
                    
                    time.sleep(1)
                    
                    # 嘗試點擊下一步
                    try:
                        code, msg = self._click_next_step()
                        if code == 0:
                            print(f"   ✓ 已點擊下一步")
                            time.sleep(2)
                            
                            # 檢查是否還有更多投票項目
                            next_page_html = self.driver.page_source
                            if "董事" in next_page_html or "全部贊成" in next_page_html:
                                print("   ℹ️  還有更多投票項目，繼續投票...")
                                continue  # 繼續下一次迴圈
                            else:
                                print("   ✓ 所有投票項目已完成")
                                return {'total': 1, 'voted': 1, 'failed': 0}
                    except Exception as e:
                        print(f"   ⚠️  無法自動點擊下一步: {str(e)[:50]}")
                        # 檢查是否已到達最終頁面
                        if "完成" in page_text or "提交成功" in page_text:
                            print("   ✓ 投票已完成")
                            return {'total': 1, 'voted': 1, 'failed': 0}
                        else:
                            print("   等待用戶手動操作...")
                            input("\n按 Enter 鍵繼續...\n")
                            continue
                else:
                    # 沒有檢測到投票選項，可能已完成或進入其他頁面
                    print("   ℹ️  未檢測到投票選項")
                    
                    # 檢查是否有提交按鈕
                    submit_buttons = self.driver.find_elements(By.XPATH, '//button[contains(text(), "提交")] | //button[contains(text(), "確認")]')
                    if submit_buttons and any(b.is_displayed() for b in submit_buttons):
                        print("   ✓ 發現『提交』或『確認』按鈕，投票流程應該已完成")
                        print("   ℹ️  返回投票完成狀態，主程式會提交投票")
                        return {'total': 1, 'voted': 1, 'failed': 0}
                    
                    # 檢查是否有「下一步」按鈕
                    next_buttons = self.driver.find_elements(By.XPATH, '//button[contains(text(), "下一步")] | //a[contains(text(), "下一步")]')
                    if next_buttons and any(b.is_displayed() for b in next_buttons):
                        print("   ℹ️  還有『下一步』按鈕")
                        try:
                            code, msg = self._click_next_step()
                            if code == 0:
                                time.sleep(2)
                                continue
                        except Exception as e:
                            print(f"   ⚠️  點擊下一步失敗: {str(e)[:50]}")
                    
                    # 調試：顯示頁面內容片段
                    print(f"   📄 頁面內容片段: {page_text[:200]}")
                    
                    # 沒有更多操作了
                    print("   ✓ 投票流程完成")
                    return {'total': 1, 'voted': 1, 'failed': 0}
            
            print("⚠️  超過最大循環次數，投票流程終止")
            return {'total': 1, 'voted': 1, 'failed': 0}
        
        except Exception as e:
            print(f"❌ 投票流程錯誤: {str(e)}")
            import traceback
            print(traceback.format_exc())
            return {'total': 0, 'voted': 0, 'failed': 0}
    
    def _find_and_vote_all(self):
        """
        尋找所有未投票的議案並進行投票
        
        Returns:
            dict: 投票統計 {'total': 總數, 'voted': 已投票, 'failed': 失敗}
        """
        try:
            print("\n" + "=" * 60)
            print("【投票流程】搜尋未投票的議案")
            print("=" * 60)
            
            time.sleep(2)
            
            # 尋找未投票的議案行
            unvoted_items = self._find_unvoted_items()
            
            if not unvoted_items:
                print("❌ 未找到任何未投票的議案")
                return {'total': 0, 'voted': 0, 'failed': 0}
            
            self.total_count = len(unvoted_items)
            print(f"\n🔍 找到 {self.total_count} 個未投票的議案")
            
            # 逐個投票
            for idx, item in enumerate(unvoted_items, 1):
                print(f"\n【議案 {idx}/{self.total_count}】")
                if self._vote_item(item, idx):
                    self.voted_count += 1
                time.sleep(0.5)
            
            # 統計結果
            result = {
                'total': self.total_count,
                'voted': self.voted_count,
                'failed': self.total_count - self.voted_count
            }
            
            print("\n" + "=" * 60)
            print("【投票完成】統計結果")
            print(f"  總計: {result['total']} 個議案")
            print(f"  ✓ 已投票: {result['voted']} 個")
            print(f"  ✗ 失敗: {result['failed']} 個")
            print("=" * 60 + "\n")
            
            return result
        
        except Exception as e:
            print(f"\n❌ 投票流程錯誤: {str(e)}")
            import traceback
            print(traceback.format_exc())
            return {'total': 0, 'voted': 0, 'failed': 0}
    
    def _find_unvoted_items(self):
        """
        尋找未投票的議案
        
        Returns:
            list: 未投票項目的 WebElement 列表
        """
        try:
            # 嘗試多種方式尋找未投票的議案
            selectors = [
                (By.XPATH, '//*[contains(text(), "未投票")]//ancestor::tr'),
                (By.XPATH, '//tr[contains(., "未投票")]'),
            ]
            
            for by, value in selectors:
                items = self.driver.find_elements(by, value)
                if items:
                    print(f"   找到 {len(items)} 個未投票的議案")
                    return items
            
            print("   嘗試備用方案...")
            return self._find_unvoted_items_fallback()
        
        except Exception as e:
            print(f"❌ 查找未投票議案失敗 (代碼:103): {str(e)}")
            return []
    
    def _find_unvoted_items_fallback(self):
        """
        備用方案：尋找未投票的議案
        """
        try:
            items = []
            rows = self.driver.find_elements(By.XPATH, '//table//tr')
            
            for row in rows:
                vote_options = row.find_elements(By.XPATH, './/input[@type="radio"]')
                if vote_options:
                    checked = row.find_elements(By.XPATH, './/input[@type="radio"][@checked]')
                    if not checked:
                        items.append(row)
            
            print(f"   備用方案找到 {len(items)} 個未投票的議案")
            return items
        
        except Exception as e:
            print(f"❌ 備用方案失敗 (代碼:104): {str(e)}")
            return []
    
    def _vote_item(self, item, item_index):
        """
        對單個議案進行投票（選擇同意）
        
        Args:
            item: 議案元素
            item_index: 議案索引
        
        Returns:
            bool: 投票成功返回True
        """
        try:
            agree_button = self._find_agree_option(item)
            
            if agree_button:
                if agree_button.get_attribute('type') == 'radio':
                    if not agree_button.is_selected():
                        self.driver.execute_script("arguments[0].click();", agree_button)
                        time.sleep(0.5)
                        if agree_button.is_selected():
                            print(f"   ✓ 已選擇「同意」")
                            return True
                else:
                    self.driver.execute_script("arguments[0].click();", agree_button)
                    time.sleep(0.5)
                    print(f"   ✓ 已點擊「同意」按鈕")
                    return True
            
            if self._vote_item_fallback(item):
                print(f"   ✓ 已選擇「同意」")
                return True
            
            print(f"   ❌ 無法投票")
            return False
        
        except Exception as e:
            print(f"❌ 投票失敗 (代碼:105): {str(e)}")
            return False
    
    def _find_agree_option(self, item):
        """
        在議案行中尋找同意選項
        
        Args:
            item: 議案元素
        
        Returns:
            WebElement or None: 同意選項的元素
        """
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
            print(f"⚠️  尋找同意選項失敗: {str(e)}")
            return None
    
    def _vote_item_fallback(self, item):
        """
        備用方案：對議案進行投票
        
        Args:
            item: 議案元素
        
        Returns:
            bool: 成功返回True
        """
        try:
            # 嘗試點擊第一個 radio button
            radios = item.find_elements(By.XPATH, './/input[@type="radio"]')
            if radios:
                self.driver.execute_script("arguments[0].click();", radios[0])
                time.sleep(0.5)
                return True
            
            # 嘗試點擊第一個按鈕
            buttons = item.find_elements(By.XPATH, './/button')
            if buttons:
                self.driver.execute_script("arguments[0].click();", buttons[0])
                time.sleep(0.5)
                return True
            
            return False
        
        except Exception as e:
            print(f"⚠️  備用投票失敗: {str(e)}")
            return False
    
    def _submit_vote(self):
        """
        提交投票
        
        Returns:
            tuple: (code, msg) - 0成功，<0為錯誤代碼
        """
        print("\n【步驟】提交投票...")
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
            print(f"   → 嘗試選擇器: {str(by)[:20]}..., 找到 {len(buttons)} 個")
            
            for idx, button in enumerate(buttons):
                try:
                    # 檢查是否可見
                    if not button.is_displayed():
                        print(f"     [按鈕 {idx}] 不可見，跳過")
                        continue
                    
                    # 嘗試1: 直接點擊
                    print(f"     [按鈕 {idx}] 嘗試直接點擊...")
                    button.click()
                    time.sleep(2)
                    print(f"   ✓ 已提交投票")
                    return (0, "已提交投票")
                    
                except Exception as e:
                    error_name = str(type(e).__name__)
                    print(f"     [按鈕 {idx}] 直接點擊失敗: {error_name}")
                    
                    # 嘗試2: 使用 JavaScript 點擊
                    try:
                        print(f"     [按鈕 {idx}] 嘗試 JavaScript 點擊...")
                        self.driver.execute_script("arguments[0].scrollIntoView(true);", button)
                        time.sleep(0.3)
                        self.driver.execute_script("arguments[0].removeAttribute('disabled');", button)
                        time.sleep(0.3)
                        self.driver.execute_script("arguments[0].click();", button)
                        time.sleep(2)
                        print(f"   ✓ 已提交投票")
                        return (0, "已使用 JavaScript 提交")
                    except Exception as e2:
                        print(f"     [按鈕 {idx}] JavaScript 也失敗: {str(type(e2).__name__)}")
                        continue
        
        print(f"   ⚠️  未找到可點擊的提交按鈕")
        raise Exception("無法找到或點擊提交按鈕")
    
    def _click_query_button(self):
        """
        點擊查詢按鈕獲取投票確認
        
        Returns:
            tuple: (code, msg) - 0成功，<0為錯誤代碼
        """
        print("\n【步驟】點擊查詢按鈕...")
        time.sleep(1)
        
        # 尋找查詢按鈕或投票確認按鈕
        query_button_selectors = [
            # 投票結果確認頁面的按鈕
            (By.XPATH, '//button[contains(text(), "確認投票結果")]'),
            (By.XPATH, '//a[contains(text(), "確認投票結果")]'),
            (By.CSS_SELECTOR, 'button[onclick="voteObj.checkVote(); return false;"]'),
            
            # # 備選：查詢按鈕
            # (By.XPATH, '//a[contains(text(), "查詢")]'),
            # (By.XPATH, '//button[contains(text(), "查詢")]'),
            
            # 其他可能的確認按鈕
            (By.XPATH, '//button[contains(text(), "確認")]'),
            (By.CSS_SELECTOR, 'button[type="submit"]'),
        ]
        
        for by, value in query_button_selectors:
            try:
                buttons = self.driver.find_elements(by, value)
                for button in buttons:
                    if button.is_displayed():
                        button_text = button.text.strip() or button.get_attribute('title') or "按鈕"
                        print(f"   找到按鈕: {button_text}，正在點擊...")
                        try:
                            button.click()
                            time.sleep(3)
                            print(f"   ✓ 已點擊按鈕")
                            return (0, f"已點擊按鈕: {button_text}")
                        except Exception as e:
                            if "not clickable" in str(e) or "ElementClickInterceptedException" in str(type(e).__name__):
                                print(f"   ⚠️  元素被遮擋，嘗試 JavaScript 點擊...")
                                self.driver.execute_script("arguments[0].scrollIntoView(true);", button)
                                time.sleep(0.5)
                                self.driver.execute_script("arguments[0].click();", button)
                                time.sleep(3)
                                print(f"   ✓ 已點擊按鈕")
                                return (0, f"已點擊按鈕: {button_text}")
                            else:
                                raise
            except Exception as e:
                # 選擇器可能無效，繼續嘗試
                continue
        
        raise Exception("無法找到查詢或確認按鈕")
    
    def _click_query_button_in_table(self, company_code):
        """
        點擊表格中指定公司行的查詢按鈕
        參考 template.py 的表格操作方式
        
        Args:
            company_code: 公司代碼 (例: "1528")
        
        Returns:
            tuple: (code, msg) - 0成功，<0為錯誤代碼
        """
        print(f"\n【步驟】點擊表格中 {company_code} 的查詢按鈕...")
        
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
                                
                                try:
                                    query_link.click()
                                    time.sleep(3)
                                    print(f"   ✓ 已點擊表格中 {company_code} 的查詢按鈕")
                                    return (0, f"已點擊表格查詢: {company_code}")
                                except Exception as e:
                                    if "not clickable" in str(e) or "ElementClickInterceptedException" in str(type(e).__name__):
                                        print(f"   ⚠️  元素被遮擋，嘗試 JavaScript 點擊...")
                                        self.driver.execute_script("arguments[0].click();", query_link)
                                        time.sleep(3)
                                        print(f"   ✓ 已使用 JavaScript 點擊表格查詢")
                                        return (0, f"已使用 JavaScript 點擊表格查詢: {company_code}")
                                    else:
                                        raise
                except Exception as e:
                    continue
            
            print(f"   ❌ 無法找到公司 {company_code} 的查詢按鈕")
            return (-1, f"無法找到公司 {company_code} 的查詢按鈕")
            
        except Exception as e:
            print(f"   ❌ 表格查詢失敗: {str(e)}")
            return (-1, f"表格查詢失敗: {str(e)}")
    
    def _go_back_to_list(self):
        """
        返回投票列表
        
        Returns:
            tuple: (code, msg) - 0成功，<0為錯誤代碼
        """
        print("\n【步驟】返回投票列表...")
        
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
                        print(f"   找到按鈕: {button_text}")
                        try:
                            button.click()
                            time.sleep(2)
                            print(f"   ✓ 已返回列表")
                            return (0, "已返回列表")
                        except Exception as e:
                            if "not clickable" in str(e) or "ElementClickInterceptedException" in str(type(e).__name__):
                                print(f"   ⚠️  元素被遮擋，嘗試 JavaScript 點擊...")
                                self.driver.execute_script("arguments[0].scrollIntoView(true);", button)
                                time.sleep(0.5)
                                self.driver.execute_script("arguments[0].click();", button)
                                time.sleep(2)
                                print(f"   ✓ 已使用 JavaScript 返回列表")
                                return (0, "已使用 JavaScript 返回列表")
                            else:
                                raise
            except Exception as e:
                # 選擇器可能無效，繼續嘗試
                continue
        
        # 方式2: 使用瀏覽器返回
        print("   📍 未發現返回按鈕，使用瀏覽器返回功能...")
        self.driver.back()
        time.sleep(2)
        print(f"   ✓ 已使用瀏覽器返回")
        return (0, "已使用瀏覽器返回")

    def screenshot_all_companies_results(self, log_msg_func, screenshot_func):
        """
        對列表中所有已投票公司依序點選查詢、截圖投票結果、點返回
        參考 template.py 的 auto_screenshot 實作方式

        Args:
            log_msg_func: log輸出函數
            screenshot_func: 截圖回調函數，簽名: screenshot_func(company_code, company_name)
        """
        log_msg_func("\n【步驟3】逐一查詢各公司投票結果並截圖...")
        time.sleep(2)

        index_take = 1  # 跳過表頭列（與 template 相同）
        all_done = False

        while not all_done:
            all_done = True
            had_refresh = False

            rows = self.driver.find_elements(By.TAG_NAME, "tr")
            # log_msg_func(f"   共找到 {len(rows) - index_take} 個公司行")

            for index, row in enumerate(rows[index_take:]):
                if had_refresh:
                    break

                if "已投票" not in row.text:
                    continue

                all_done = False

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
                    log_msg_func(f"{company_name} ({company_code}) 已截圖，跳過")
                    index_take = index_take + 1
                    had_refresh = True
                    break

                # log_msg_func(f"\n   {company_name} ({company_code}) 查詢中...")

                # 參考 template：操作欄是 td[3]，找含「查詢」的 <a>
                query_clicked = False
                try:
                    for link in row.find_elements(By.TAG_NAME, "td")[3].find_elements(By.TAG_NAME, "a"):
                        if "查詢" in link.text:
                            link.click()
                            query_clicked = True
                            break
                except Exception as e:
                    log_msg_func(f"   ⚠️  點擊查詢失敗: {str(e)[:60]}")

                if not query_clicked:
                    log_msg_func(f"   ⚠️  找不到查詢連結，跳過")
                    continue

                time.sleep(3)

                # 截圖
                try:
                    screenshot_func(company_code, company_name)
                    self.screenshotted_companies.add(company_code)
                except Exception as e:
                    log_msg_func(f"   ⚠️  截圖失敗: {str(e)[:50]}")

                # 捲到底部
                self.driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
                time.sleep(1)

                # 參考 template 的返回方式
                try:
                    self.driver.execute_script(
                        "arguments[0].click();",
                        self.driver.find_element(By.CSS_SELECTOR, 'button[onclick="back(); return false;"]')
                    )
                except Exception:
                    # 備選：找含「返回」文字的按鈕
                    back_clicked = False
                    for by, val in [
                        (By.XPATH, '//a[contains(text(), "返回")]'),
                        (By.XPATH, '//button[contains(text(), "返回")]'),
                    ]:
                        btns = self.driver.find_elements(by, val)
                        for btn in btns:
                            if btn.is_displayed():
                                self.driver.execute_script("arguments[0].click();", btn)
                                back_clicked = True
                                break
                        if back_clicked:
                            break
                    if not back_clicked:
                        self.driver.back()
                        log_msg_func(f"   ⚠️  未找到返回按鈕，使用瀏覽器返回")

                time.sleep(2)
                log_msg_func(f"   ✓ 已完成 {company_name} ({company_code})")

                had_refresh = True
                index_take = index_take + 1

    def _find_all_unvoted_companies(self):
        """
        尋找所有未投票的公司
        
        Returns:
            list: 未投票的公司行元素列表，若無則返回空列表
        """
        print("\n🔍 掃描未投票的公司...")
        time.sleep(3)  # 等待頁面加載
        
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
                    print(f"   ✓ 找到 {len(rows)} 個未投票的公司")
                    return rows
            except Exception as e:
                print(f"   ⚠️  選擇器失敗: {str(e)[:50]}")
                continue
        
        # 調試：顯示頁面內容
        print("\n   📋 頁面內容分析：")
        page_text = self.driver.find_element(By.TAG_NAME, 'body').text
        if "未投票" in page_text:
            print(f"   ✓ 頁面包含「未投票」文字，但選擇器無法定位")
            print(f"   ⚠️  可能所有公司都已投票 (已投票|投票結果)")
        else:
            print(f"   ✗ 頁面不包含「未投票」文字")
            print(f"   頁面前 500 字符: {page_text[:500]}")
        
        # 返回空列表而不是 raise exception，讓上層處理
        return []
    
    def execute_voting_loop(self, log_msg_func, screenshot_func=None):
        """
        執行投票循環並返回統計結果
        
        Args:
            log_msg_func: 用於輸出log的函數
            screenshot_func: 可選的截圖回調函數，簽名: screenshot_func(company_code)
        
        Returns:
            dict: {
                'total': 總數, 
                'voted': 已投票數, 
                'failed': 失敗數,
                'has_unvoted': 是否找到未投票公司
            }
        """
        # 掃描未投票的公司
        unvoted_companies = self._find_all_unvoted_companies()
        
        # 如果沒找到未投票公司，返回狀態
        if not unvoted_companies:
            log_msg_func("❌ 未找到未投票的公司")
            return {
                'total': 0,
                'voted': 0,
                'failed': 0,
                'has_unvoted': False
            }
        
        log_msg_func(f"\n📋 找到 {len(unvoted_companies)} 個未投票的公司\n")
        
        total_voted = 0
        total_failed = 0
        
        for idx, company_row in enumerate(unvoted_companies, 1):
            # 獲取公司代碼和名稱
            code_elem = company_row.find_element(By.XPATH, './/td[1]')
            first_col_text = code_elem.text.strip()
            
            # 解析代碼和名稱 (格式: "2102 泰豐" 或類似)
            parts = first_col_text.split()
            company_code = parts[0] if parts else ""
            company_name = " ".join(parts[1:]) if len(parts) > 1 else "未知"
            
            log_msg_func(f"\n═══ 公司 [{idx}] {company_name} ({company_code}) ═══")
            
            # 在當前公司行中查找和點擊投票按鈕
            try:
                vote_button = company_row.find_element(By.XPATH, './/a[contains(text(), "投票")] | .//button[contains(text(), "投票")]')
                log_msg_func(f"【投票流程】點擊投票按鈕 (代碼: {company_code})")
                vote_button.click()
                time.sleep(3)  # 等待投票頁面加載
                log_msg_func("✓ 已點擊投票按鈕")
            except Exception as e:
                log_msg_func(f"⚠️  無法找到或點擊投票按鈕: {str(e)[:50]}")
                log_msg_func("嘗試使用備用方法...")
                code, msg = self._click_vote_button()
            
            vote_result = self._vote_with_agree_button()
            
            if vote_result.get('failed', 0) > 0:
                log_msg_func("⚠️  快速投票失敗，進行逐個投票...")
                vote_result = self._find_and_vote_all()
            
            if vote_result.get('total', 0) > 0:
                log_msg_func("✓ 提交投票...")
                code, msg = self._submit_vote()
                time.sleep(1)
                
                # 點擊確認按鈕就會自動返回投票列表
                code, msg = self._click_query_button()
                time.sleep(2)
                
                # 確認後進行截圖（如果提供了截圖函數）
                if screenshot_func:
                    try:
                        log_msg_func(f"\n【截圖】記錄公司 {company_name} ({company_code}) 的投票結果...")
                        screenshot_func(company_code, company_name)
                        self.screenshotted_companies.add(company_code)
                    except Exception as e:
                        log_msg_func(f"⚠️  截圖失敗: {str(e)[:50]}")
                
                total_voted += 1
            else:
                total_failed += 1
            
            # 準備處理下一個公司
            time.sleep(1)
        
        return {
            'total': len(unvoted_companies),
            'voted': total_voted,
            'failed': total_failed,
            'has_unvoted': True
        }

