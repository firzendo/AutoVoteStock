#!/usr/bin/env python
# coding: utf-8
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from screenshot_handler import ScreenshotHandler


class VoteHandler:    
    def __init__(self, driver: webdriver.Chrome, page_navigator, screenshot_handler: ScreenshotHandler = None, screenshot_dir: str = "screenshots"):
        self.driver = driver
        self.vote_timeout = 30
        self.voted_count = 0
        self.total_count = 0
        self.page_navigator = page_navigator
        # 使用傳入的 screenshot_handler 或創建新的
        self.screenshot_handler = screenshot_handler if screenshot_handler else ScreenshotHandler(driver, page_navigator, screenshot_dir)
    
    def _handle_director_election(self):
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
            agree_button = self.page_navigator.find_agree_button_for_director()
            if agree_button:
                print("   ✓ 已點擊『全部贊成』按鈕，準備進入下一步...")
                # 重要：點擊「全部贊成」後也要執行「下一步」來推進流程
                time.sleep(2)
                try:
                    code, msg = self.page_navigator.click_next_step()
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
            if self.page_navigator.check_all_directors():
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
            if self.page_navigator.click_average_distribution():
                print("   ✓ 已點擊平均分配")
            else:
                print("   ⚠️  未能點擊平均分配")
                print("   ⚠️  等待用戶手動操作...")
                input("\n⏳ 請手動點擊平均分配，然後按 Enter 繼續...\n")
            
            time.sleep(2)
            
            # 步驟3: 點擊下一步
            try:
                code, msg = self.page_navigator.click_next_step()
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

    def _vote_with_agree_button(self):
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
                        code, msg = self.page_navigator.click_all_agree()
                        if code == 0:
                            print(f"   ✓ {msg}")
                    except Exception as e:
                        print(f"   ⚠️  無法自動點擊全部贊成: {str(e)[:50]}")
                        print("   請手動操作後按 Enter 繼續")
                        input("\n按 Enter 鍵繼續...\n")
                    
                    time.sleep(1)
                    
                    # 嘗試點擊下一步
                    try:
                        code, msg = self.page_navigator.click_next_step()
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
        try:
            print("\n" + "=" * 60)
            print("【投票流程】搜尋未投票的議案")
            print("=" * 60)
            
            time.sleep(2)
            
            # 尋找未投票的議案行
            unvoted_items = self.page_navigator.find_unvoted_items()
            
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
    
    def _vote_item(self, item, item_index):
        try:
            agree_button = self.page_navigator.find_agree_option(item)
            
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
    
    def _go_back_to_list(self):
        success, msg = self.page_navigator.go_back_to_list()
        return (0 if success else -1, msg)


    def execute_voting_loop(self, log_msg_func):
        total_voted = 0
        total_failed = 0
        total_scanned = 0
        
        # 循環投票：每次投票完成後重新掃描，而非預先保存所有元素
        # 原因：每次投票返回列表時頁面會刷新，預先保存的元素引用會變成陳舊(stale)
        while True:
            # 重新掃描未投票的公司（使用 PageNavigator）
            unvoted_companies = self.page_navigator.find_all_unvoted_companies()
            
            if not unvoted_companies:
                # 當前頁沒有未投票公司，嘗試翻到下一頁
                log_msg_func("📄 當前頁無未投票公司，嘗試翻頁...")
                
                if self.page_navigator.go_to_next_page():
                    # 成功翻頁，繼續掃描
                    time.sleep(2)
                    continue
                else:
                    # 無下一頁或翻頁失敗，投票循環完成
                    log_msg_func("✓ 已完成所有頁面投票，投票循環結束")
                    break
            
            total_scanned += 1
            
            # 獲取第一個未投票公司
            company_row = unvoted_companies[0]
            
            try:
                # 重新獲取公司代碼和名稱（避免使用陳舊的元素引用）
                code_elem = company_row.find_element(By.XPATH, './/td[1]')
                first_col_text = code_elem.text.strip()
                
                # 解析代碼和名稱 (格式: "2102 泰豐" 或類似)
                parts = first_col_text.split()
                company_code = parts[0] if parts else ""
                company_name = " ".join(parts[1:]) if len(parts) > 1 else "未知"
                
                log_msg_func(f"\n═══ 公司 [{total_voted + total_failed + 1}] {company_name} ({company_code}) ═══")
            
            except Exception as e:
                log_msg_func(f"⚠️  無法取得公司信息: {str(e)[:50]}")
                total_failed += 1
                continue
            
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
                try:
                    code, msg = self.page_navigator.click_vote_button()
                except Exception as e2:
                    log_msg_func(f"❌ 備用方法也失敗: {str(e2)[:50]}")
                    total_failed += 1
                    continue
            
            # 進行投票
            vote_result = self._vote_with_agree_button()
            
            if vote_result.get('failed', 0) > 0:
                log_msg_func("⚠️  快速投票失敗，進行逐個投票...")
                vote_result = self._find_and_vote_all()
            
            if vote_result.get('total', 0) > 0:
                log_msg_func("✓ 提交投票...")
                try:
                    code, msg = self.page_navigator.submit_vote()
                    time.sleep(1)
                    
                    # 點擊確認按鈕就會自動返回投票列表
                    code, msg = self.page_navigator.click_query_button()
                    time.sleep(3)  # 等待頁面返回列表並刷新
                    
                    log_msg_func("✓ 投票完成，準備掃描下一個公司...")
                    total_voted += 1
                
                except Exception as e:
                    log_msg_func(f"❌ 提交或確認失敗: {str(e)[:60]}")
                    total_failed += 1
                    # 嘗試返回列表
                    try:
                        self.driver.back()
                        time.sleep(2)
                    except:
                        pass
            else:
                log_msg_func(f"❌ 投票內容為空或無法識別")
                total_failed += 1
            
            # 防止無限迴圈：如果掃描次數太多仍有未投票公司，可能是頁面問題
            if total_scanned > 1000:
                log_msg_func(f"⚠️  掃描次數達到1000次，可能存在頁面問題，停止投票")
                break
        
        return {
            'total': total_voted + total_failed,
            'voted': total_voted,
            'failed': total_failed,
            'has_unvoted': bool(unvoted_companies) if total_scanned < 100 else False
        }

