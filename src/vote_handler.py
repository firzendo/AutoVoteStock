#!/usr/bin/env python
# coding: utf-8
import time
import functools
import logging
from enum import Enum
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, StaleElementReferenceException
from screenshot_handler import ScreenshotHandler


# ── VoteState ────────────────────────────────────────────────────
class VoteState(Enum):
    LIST    = "list"      # 投票列表頁
    VOTING  = "voting"    # 投票操作中
    CONFIRM = "confirm"   # 確認頁面
    DONE    = "done"      # 完成


# ── retry decorator ──────────────────────────────────────────────
def retry(n: int = 3, delay: float = 1.0, exceptions: tuple = (Exception,)):
    """重試 decorator：失敗時最多重試 n 次，每次間隔 delay 秒"""
    def decorator(func):
        @functools.wraps(func)
        def wrapper(*args, **kwargs):
            for attempt in range(1, n + 1):
                try:
                    return func(*args, **kwargs)
                except exceptions as exc:
                    if attempt == n:
                        raise
                    print(f"   ⚠️  [{func.__name__}] 第 {attempt} 次失敗: {exc!s:.80} — 重試中...")
                    time.sleep(delay)
        return wrapper
    return decorator


class VoteHandler:    
    def __init__(self, driver: webdriver.Chrome, page_navigator, screenshot_handler: ScreenshotHandler = None, screenshot_dir: str = "screenshots"):
        self.driver = driver
        self.vote_timeout = 30
        self.voted_count = 0
        self.total_count = 0
        self.page_navigator = page_navigator
        # 使用傳入的 screenshot_handler 或創建新的
        self.screenshot_handler = screenshot_handler if screenshot_handler else ScreenshotHandler(driver, page_navigator, screenshot_dir)

    # ── 共用工具 ─────────────────────────────────────────────────
    def _has_dom_element(self, xpath: str) -> bool:
        """結構判斷：找到至少一個可見元素則回傳 True（取代 'xxx' in page_html）"""
        try:
            return any(e.is_displayed() for e in self.driver.find_elements(By.XPATH, xpath))
        except Exception:
            return False

    def _wait_clickable(self, xpath: str, timeout: int = 10) -> bool:
        """等待 XPATH 元素可點擊；逾時不拋例外，回傳 False"""
        try:
            WebDriverWait(self.driver, timeout).until(
                EC.element_to_be_clickable((By.XPATH, xpath))
            )
            return True
        except TimeoutException:
            return False

    def _wait_page_ready(self, timeout: int = 10) -> None:
        """等待頁面 body 出現（輕量頁面就緒判斷）"""
        try:
            WebDriverWait(self.driver, timeout).until(
                EC.presence_of_element_located((By.TAG_NAME, "body"))
            )
        except TimeoutException:
            pass
    
    # ── 董事選舉 ─────────────────────────────────────────────────
    @retry(n=3, delay=1.5, exceptions=(StaleElementReferenceException, TimeoutException))
    def _handle_director_election(self):
        try:
            print("\n" + "=" * 60)
            print("【投票流程】董事選舉 - 全部勾選 + 平均分配")
            print("=" * 60)

            # 等待頁面就緒（取代 time.sleep(1)）
            self._wait_page_ready()

            # 結構判斷：確認是董事選舉頁面（取代 "董事" in page_html）
            if not self._has_dom_element("//*[contains(text(),'董事')]"):
                print("   ⚠️  未檢測到董事選舉頁面")
                return False

            print("   ✓ 檢測到董事選舉頁面")

            # 新規則：直接有「全部贊成」按鈕
            agree_button = self.page_navigator.find_agree_button_for_director()
            if agree_button:
                print("   ✓ 已點擊『全部贊成』按鈕，準備進入下一步...")
                # 等待下一步按鈕可點擊（取代 time.sleep(2)）
                self._wait_clickable("//button[contains(text(),'下一步')] | //a[contains(text(),'下一步')]")
                try:
                    code, msg = self.page_navigator.click_next_step()
                    if code == 0:
                        print("   ✓ 已點擊下一步，董事選舉完成，進入下一個議案")
                        self._wait_page_ready()     # 取代 time.sleep(2)
                        return True
                    else:
                        print(f"   ⚠️  點擊下一步失敗: {msg}")
                        return True
                except Exception as e:
                    print(f"   ⚠️  無法點擊下一步: {str(e)[:50]}")
                    return True

            # 步驟1: 全部勾選複選框
            if self.page_navigator.check_all_directors():
                print("   ✓ 已全部勾選")
            else:
                print("   ⚠️  未能成功全部勾選，等待用戶手動勾選...")
                input("\n⏳ 請手動勾選所有董事，然後按 Enter 繼續...\n")

            # 等待平均分配按鈕可點擊（取代 time.sleep(3)）
            print("   ⏳ 等待勾選狀態確認...")
            self._wait_clickable("//*[contains(text(),'平均分配')]")

            # 步驟2: 平均分配
            print("\n   準備點擊平均分配按鈕...")
            if self.page_navigator.click_average_distribution():
                print("   ✓ 已點擊平均分配")
            else:
                print("   ⚠️  未能點擊平均分配，等待用戶手動操作...")
                input("\n⏳ 請手動點擊平均分配，然後按 Enter 繼續...\n")

            # 等待下一步按鈕可點擊（取代 time.sleep(2)）
            self._wait_clickable("//button[contains(text(),'下一步')] | //a[contains(text(),'下一步')]")

            # 步驟3: 點擊下一步
            try:
                code, msg = self.page_navigator.click_next_step()
                if code == 0:
                    print("   ✓ 已點擊下一步")
                    # 等待新頁面就緒（取代 time.sleep(2) + time.sleep(2)）
                    print("\n   ⏳ 等待新頁面加載...")
                    self._wait_page_ready()

                    # 結構判斷：是否還有投票選項（取代字串搜尋）
                    _vote_xpath = (
                        "//button[contains(text(),'全部贊成')] | //a[contains(text(),'全部贊成')]"
                        " | //button[contains(text(),'全部反對')] | //button[contains(text(),'全部棄權')]"
                    )
                    if self._has_dom_element(_vote_xpath):
                        print("   ✓ 檢測到更多投票項目，繼續處理")
                    return True
                else:
                    print("   ⚠️  無法點擊下一步，等待用戶手動操作...")
                    input("\n⏳ 請手動點擊下一步，然後按 Enter 繼續...\n")
                    return True
            except Exception as e:
                print(f"   ❌ 點擊下一步失敗: {str(e)[:50]}")
                return True

        except Exception as e:
            print(f"   ❌ 處理董事選舉時出錯: {str(e)}")
            self.screenshot_handler.save_error_screenshot("director_election_error")
            return False

    # ── 快速投票（主流程）────────────────────────────────────────
    @retry(n=3, delay=1.5, exceptions=(StaleElementReferenceException, TimeoutException))
    def _vote_with_agree_button(self):
        try:
            print("\n" + "=" * 60)
            print("【投票流程】快速投票")
            print("=" * 60)

            state = VoteState.VOTING

            # 等待投票頁面出現任何操作元素（取代 time.sleep(3)）
            _any_vote_elem = (
                "//button[contains(text(),'全部贊成')] | //a[contains(text(),'全部贊成')]"
                " | //button[contains(text(),'下一步')] | //*[contains(text(),'董事')]"
            )
            self._wait_clickable(_any_vote_elem, timeout=10)

            # 診斷：如果還在列表頁面，等待跳轉（取代 time.sleep(3)）
            if (self._has_dom_element("//*[contains(text(),'未投票')]")
                    and self._has_dom_element("//*[contains(text(),'投票狀況')]")
                    and not self._has_dom_element("//button[contains(text(),'全部贊成')] | //a[contains(text(),'全部贊成')]")):
                print("   ⚠️  似乎還在列表頁面，等待跳轉...")
                self._wait_clickable("//button[contains(text(),'全部贊成')] | //a[contains(text(),'全部贊成')]", timeout=10)

            iteration    = 0
            max_iter     = 10
            last_page_text = ""
            repeat_count   = 0

            while iteration < max_iter:
                iteration += 1
                print(f"\n【投票流程 - 循環 {iteration}】(state={state.value})")

                # 輕量等待（取代 time.sleep(1)）
                self._wait_page_ready(timeout=5)

                page_text = self.driver.find_element(By.TAG_NAME, 'body').text
                print(f"   □ 頁面文字長度: {len(page_text)} 字符")

                # ── 結構判斷（取代所有 "xxx" in page_html）──────────
                _vote_opts_xpath = (
                    "//button[contains(text(),'全部贊成')] | //a[contains(text(),'全部贊成')]"
                    " | //button[contains(text(),'全部反對')] | //button[contains(text(),'全部棄權')]"
                )
                has_confirm   = self._has_dom_element(
                    "//*[contains(text(),'投票內容確認')] | //*[contains(text(),'投票結果')]")
                has_director  = self._has_dom_element("//*[contains(text(),'董事')]")
                has_vote_opts = self._has_dom_element(_vote_opts_xpath)
                has_next_step = self._has_dom_element(
                    "//button[contains(text(),'下一步')] | //a[contains(text(),'下一步')]")

                if has_vote_opts:  print("   ✓ 頁面包含投票選項")
                if has_director:   print("   ✓ 頁面包含『董事』")
                if has_next_step:  print("   ✓ 頁面包含『下一步』")

                # ── 重複頁面偵測 ───────────────────────────────────
                if last_page_text and page_text[:100] == last_page_text[:100] and len(page_text) == len(last_page_text):
                    repeat_count += 1
                    print(f"   ⚠️  偵測到重複頁面 (第 {repeat_count} 次)")
                    if repeat_count >= 3:
                        print("   ⚠️  重複頁面過多，跳出避免無限迴圈")
                        state = VoteState.DONE
                        return {'total': 1, 'voted': 1, 'failed': 0}
                else:
                    repeat_count = 0
                last_page_text = page_text

                # ════════════════ 狀態機 ════════════════════════════

                # 優先1: 確認/結果頁面 → DONE
                if has_confirm:
                    print("   ✓ 檢測到確認/結果頁面 → 投票完成")
                    state = VoteState.DONE
                    return {'total': 1, 'voted': 1, 'failed': 0}

                # 優先2: 董事選舉（非確認頁）
                if has_director and not has_confirm:
                    print("   ✓ 檢測到董事選舉頁面（投票操作）")
                    state = VoteState.VOTING
                    if self._handle_director_election():
                        print("   ✓ 董事選舉已完成，繼續下一議案...")
                        self._wait_page_ready()     # 取代 time.sleep(2)
                        continue
                    else:
                        print("   ⚠️  董事選舉處理失敗")
                        state = VoteState.DONE
                        return {'total': 1, 'voted': 1, 'failed': 0}

                # 優先3: 標準投票選項
                if has_vote_opts:
                    print("   ✓ 檢測到標準投票選項")
                    state = VoteState.VOTING
                    try:
                        code, msg = self.page_navigator.click_all_agree()
                        if code == 0:
                            print(f"   ✓ {msg}")
                    except Exception as e:
                        print(f"   ⚠️  無法自動點擊全部贊成: {str(e)[:50]}")
                        input("\n按 Enter 鍵繼續...\n")

                    # 等待下一步按鈕可點擊（取代 time.sleep(1)）
                    self._wait_clickable(
                        "//button[contains(text(),'下一步')] | //a[contains(text(),'下一步')]",
                        timeout=5)

                    try:
                        code, msg = self.page_navigator.click_next_step()
                        if code == 0:
                            print("   ✓ 已點擊下一步")
                            state = VoteState.CONFIRM
                            self._wait_page_ready()     # 取代 time.sleep(2)
                            # 結構判斷：還有更多投票項目？
                            if self._has_dom_element("//*[contains(text(),'董事')]") or self._has_dom_element(_vote_opts_xpath):
                                print("   ℹ️  還有更多投票項目，繼續...")
                                state = VoteState.VOTING
                                continue
                            else:
                                print("   ✓ 所有投票項目已完成")
                                state = VoteState.DONE
                                return {'total': 1, 'voted': 1, 'failed': 0}
                    except Exception as e:
                        print(f"   ⚠️  無法自動點擊下一步: {str(e)[:50]}")
                        if "完成" in page_text or "提交成功" in page_text:
                            state = VoteState.DONE
                            return {'total': 1, 'voted': 1, 'failed': 0}
                        input("\n按 Enter 鍵繼續...\n")
                        continue

                else:
                    # 沒有投票選項，檢查提交/下一步
                    print("   ℹ️  未檢測到投票選項")
                    submit_btns = self.driver.find_elements(
                        By.XPATH, '//button[contains(text(), "提交")] | //button[contains(text(), "確認")]')
                    if submit_btns and any(b.is_displayed() for b in submit_btns):
                        print("   ✓ 發現提交/確認按鈕，投票完成")
                        state = VoteState.DONE
                        return {'total': 1, 'voted': 1, 'failed': 0}

                    if has_next_step:
                        print("   ℹ️  還有『下一步』按鈕")
                        try:
                            code, msg = self.page_navigator.click_next_step()
                            if code == 0:
                                self._wait_page_ready()
                                continue
                        except Exception as e:
                            print(f"   ⚠️  點擊下一步失敗: {str(e)[:50]}")

                    print(f"   📄 頁面內容片段: {page_text[:200]}")
                    print("   ✓ 投票流程完成")
                    state = VoteState.DONE
                    return {'total': 1, 'voted': 1, 'failed': 0}

            print("⚠️  超過最大循環次數，投票流程終止")
            return {'total': 1, 'voted': 1, 'failed': 0}

        except Exception as e:
            print(f"❌ 投票流程錯誤: {str(e)}")
            import traceback
            print(traceback.format_exc())
            self.screenshot_handler.save_error_screenshot("vote_with_agree_error")
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
                    # 成功翻頁，等待頁面就緒後繼續掃描（取代 time.sleep(2)）
                    self._wait_page_ready(timeout=10)
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
                # 等待投票頁面任何操作元素出現（取代 time.sleep(3)）
                self._wait_clickable(
                    "//button[contains(text(),'全部贊成')] | //a[contains(text(),'全部贊成')]"
                    " | //button[contains(text(),'下一步')] | //*[contains(text(),'董事')]",
                    timeout=10)
                log_msg_func("✓ 已點擊投票按鈕")
            except Exception as e:
                log_msg_func(f"⚠️  無法找到或點擊投票按鈕: {str(e)[:50]}")
                log_msg_func("嘗試使用備用方法...")
                try:
                    code, msg = self.page_navigator.click_vote_button()
                except Exception as e2:
                    log_msg_func(f"❌ 備用方法也失敗: {str(e2)[:50]}")
                    self.screenshot_handler.save_error_screenshot(f"vote_btn_fail_{company_code}")
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
                    # 等待確認按鈕（取代 time.sleep(1)）
                    self._wait_clickable(
                        "//button[contains(text(),'確認')] | //a[contains(text(),'確認')]",
                        timeout=5)
                    
                    # 點擊確認按鈕就會自動返回投票列表
                    code, msg = self.page_navigator.click_query_button()
                    # 等待列表頁面刷新（取代 time.sleep(3)）
                    self._wait_clickable(
                        "//*[contains(text(),'未投票')] | //*[contains(text(),'已投票')]",
                        timeout=10)
                    
                    log_msg_func("✓ 投票完成，準備掃描下一個公司...")
                    total_voted += 1
                
                except Exception as e:
                    log_msg_func(f"❌ 提交或確認失敗: {str(e)[:60]}")
                    self.screenshot_handler.save_error_screenshot(f"submit_fail_{company_code}")
                    total_failed += 1
                    # 嘗試返回列表
                    try:
                        self.driver.back()
                        self._wait_page_ready(timeout=5)
                    except Exception:
                        pass
            else:
                log_msg_func(f"❌ 投票內容為空或無法識別")
                self.screenshot_handler.save_error_screenshot(f"vote_empty_{company_code}")
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

