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

logger = logging.getLogger(__name__)

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
                    logger.warning("[%s] 第 %d 次失敗: %.80s — 重試中...", func.__name__, attempt, exc)
                    time.sleep(delay)
        return wrapper
    return decorator


class VoteHandler:    
    def __init__(self, driver: webdriver.Chrome, page_navigator, screenshot_handler: ScreenshotHandler = None, screenshot_dir: str = "screenshots"):
        self.driver = driver
        self.vote_timeout = 30
        self.voted_count = 0
        self.total_count = 0
        self.is_maintenance = False  # 標記是否檢測到系統維護中
        self.companies_info = []  # 記錄每個公司的投票信息 (code, name, status)
        self.page_navigator = page_navigator
        # 使用傳入的 screenshot_handler 或創建新的
        self.screenshot_handler = screenshot_handler if screenshot_handler else ScreenshotHandler(driver, page_navigator, screenshot_dir)

    # ── 共用工具 ─────────────────────────────────────────────────
    def _has_dom_element(self, xpath: str) -> bool:
        """結構判斷：DOM 中存在至少一個匹配元素則回傳 True。
        不呼叫 is_displayed()，避免在 while loop 裡觸發昂貴的 render check。"""
        try:
            return len(self.driver.find_elements(By.XPATH, xpath)) > 0
        except Exception:
            return False

    def _wait_and_get(self, xpath: str, timeout: int = 10):
        """等待元素可點擊並回傳該 element（同一個 instance，避免 race condition）。
        逾時則回傳 None，不拋例外。"""
        try:
            return WebDriverWait(self.driver, timeout).until(
                EC.element_to_be_clickable((By.XPATH, xpath))
            )
        except TimeoutException:
            return None

    def _wait_clickable(self, xpath: str, timeout: int = 10) -> bool:
        """等待 XPATH 元素可點擊；只需確認可點擊時使用（不需要操作 element）。
        逾時回傳 False，不拋例外。"""
        return self._wait_and_get(xpath, timeout) is not None

    def _retry_click(self, xpath: str, timeout: int = 10, retries: int = 3) -> bool:
        """對單一動作（點擊）做細粒度 retry：
        wait + get + click，遇到 StaleElementReferenceException 時重試。
        只 retry 這一個動作，不影響外層流程狀態。"""
        for attempt in range(1, retries + 1):
            try:
                el = self._wait_and_get(xpath, timeout)
                if el is None:
                    return False
                el.click()
                return True
            except StaleElementReferenceException:
                if attempt == retries:
                    return False
                logger.warning("_retry_click stale (第%d次)，重試...", attempt)
                time.sleep(0.5)
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
    def _handle_director_election(self):
        # ⚠️  不加 @retry：此方法為 workflow，重試可能導致重複點擊已改變的 DOM
        try:
            logger.info("=" * 60)
            logger.info("【投票流程】董事選舉 - 全部勾選 + 平均分配")

            self._wait_page_ready()

            # 結構判斷（. = 包含所有子節點文字，比 text() 更穩）
            if not self._has_dom_element("//*[contains(.,'董事')]"):
                logger.warning("未檢測到董事選舉頁面")
                return False

            logger.info("✓ 檢測到董事選舉頁面")

            # 直接有「全部贊成」按鈕 → 用 _retry_click 點擊（單一動作 retry）
            _next_xpath = "//button[contains(.,'下一步')] | //a[contains(.,'下一步')]"
            agree_button = self.page_navigator.find_agree_button_for_director()
            if agree_button:
                logger.info("✓ 已點擊『全部贊成』按鈕，準備進入下一步...")
                # _retry_click：wait + get + click 原子操作，避免 stale race condition
                if self._retry_click(_next_xpath):
                    logger.info("✓ 已點擊下一步，董事選舉完成，進入下一個議案")
                    self._wait_page_ready()
                    return True
                else:
                    logger.warning("下一步按鈕未出現或點擊失敗")
                    return True

            # 步驟1: 全部勾選複選框
            if self.page_navigator.check_all_directors():
                logger.info("✓ 已全部勾選")
            else:
                logger.warning("未能成功全部勾選，等待用戶手動勾選...")
                input("\n⏳ 請手動勾選所有董事，然後按 Enter 繼續...\n")

            # 等待平均分配按鈕（只做 guard，click 由 page_navigator 負責）
            logger.info("⏳ 等待勾選狀態確認...")
            self._wait_clickable("//*[contains(.,'平均分配')]")

            # 步驟2: 平均分配
            logger.info("準備點擊平均分配按鈕...")
            # self.screenshot_handler.capture("before_director_agree_next")
            if self.page_navigator.click_average_distribution():
                logger.info("✓ 已點擊平均分配")
            else:
                logger.warning("未能點擊平均分配，等待用戶手動操作...")
                input("\n⏳ 請手動點擊平均分配，然後按 Enter 繼續...\n")

            # 步驟3: 用 _retry_click 點擊下一步（單一動作 retry）
            logger.info("⏳ 等待新頁面加載...")
            # self.screenshot_handler.capture("before_director_next_step")
            if self._retry_click(_next_xpath):
                logger.info("✓ 已點擊下一步")
                self._wait_page_ready()

                # 結構判斷：是否還有投票選項
                _vote_xpath = (
                    "//button[contains(.,'全部贊成')] | //a[contains(.,'全部贊成')]"
                    " | //button[contains(.,'全部反對')] | //button[contains(.,'全部棄權')]"
                )
                if self._has_dom_element(_vote_xpath):
                    logger.info("✓ 檢測到更多投票項目，繼續處理")
                return True
            else:
                logger.warning("無法點擊下一步，等待用戶手動操作...")
                input("\n⏳ 請手動點擊下一步，然後按 Enter 繼續...\n")
                return True

        except Exception as e:
            logger.exception("處理董事選舉時出錯")
            self.screenshot_handler.save_error_screenshot("director_election_error")
            return False

    # ── FSM 狀態處理器 ────────────────────────────────────────────
    # XPath 常數（. = 包含所有子節點文字，比 text() 更穩健）
    _XPATH_VOTE_OPTS = (
        "//button[contains(.,'全部贊成')] | //a[contains(.,'全部贊成')]"
        " | //button[contains(.,'全部反對')] | //button[contains(.,'全部棄權')]"
    )
    _XPATH_NEXT      = "//button[contains(.,'下一步')] | //a[contains(.,'下一步')]"
    _XPATH_CONFIRM   = "//*[contains(.,'投票內容確認')] | //*[contains(.,'投票結果')]"
    _XPATH_DIRECTOR  = "//*[contains(.,'董事')]"
    _XPATH_SUBMIT    = "//button[contains(.,'提交')] | //button[contains(.,'確認')]"
    _XPATH_MAINTENANCE = "//*[contains(.,'系統維護中')] | //*[contains(.,'Scheduled system maintenance')]"

    def _detect_maintenance_message(self) -> bool:
        """檢測系統是否顯示維護訊息
        
        返回值：
        - True: 系統在維護中，應直接進入登出流程
        - False: 系統正常運作，可進行投票
        """
        if self._has_dom_element(self._XPATH_MAINTENANCE):
            logger.info("🔧 檢測到系統維護訊息，略過投票流程")
            return True
        return False

    def _detect_state(self) -> VoteState:
        """依 DOM 結構決定目前應處於哪個 VoteState"""
        if self._has_dom_element(self._XPATH_CONFIRM):
            return VoteState.CONFIRM
        if self._has_dom_element(self._XPATH_DIRECTOR):
            return VoteState.VOTING   # 董事選舉也屬 VOTING
        if self._has_dom_element(self._XPATH_VOTE_OPTS):
            return VoteState.VOTING
        if self._has_dom_element(self._XPATH_SUBMIT):
            return VoteState.CONFIRM
        return VoteState.DONE         # 無可操作元素 → 視為已完成

    def _handle_state_confirm(self) -> dict:
        logger.info("✓ 確認/結果頁面 → 投票完成")
        # self.screenshot_handler.capture("state_confirm")
        return {'total': 1, 'voted': 1, 'failed': 0}

    def _handle_state_voting(self, page_text: str) -> dict | None:
        """回傳 dict 表示流程結束；回傳 None 表示繼續迴圈"""
        # 董事選舉子流程
        if self._has_dom_element(self._XPATH_DIRECTOR):
            logger.info("✓ 檢測到董事選舉頁面")
            if self._handle_director_election():
                logger.info("✓ 董事選舉完成，繼續下一議案...")
                self._wait_page_ready()
                return None   # continue loop
            else:
                logger.warning("董事選舉處理失敗")
                return {'total': 1, 'voted': 1, 'failed': 0}

        # 標準投票：_retry_click 做單一動作 retry（wait+get+click 原子）
        if self._has_dom_element(self._XPATH_VOTE_OPTS):
            logger.info("✓ 標準投票選項")
            # self.screenshot_handler.capture("before_agree")
            try:
                code, msg = self.page_navigator.click_all_agree()
                if code == 0:
                    logger.info("✓ %s", msg)
            except Exception as e:
                logger.warning("無法自動點擊全部贊成: %s", str(e)[:50])
                input("\n按 Enter 鍵繼續...\n")

            # 用 _retry_click 點「下一步」（單一動作細粒度 retry，不重跑整個 workflow）
            # self.screenshot_handler.capture("before_next_step")
            clicked = self._retry_click(self._XPATH_NEXT, timeout=5)
            if clicked:
                logger.info("✓ 已點擊下一步")
                self._wait_page_ready()
                return None   # continue loop — 下次迴圈 _detect_state 判斷新狀態
            else:
                logger.warning("無法自動點擊下一步")
                if "完成" in page_text or "提交成功" in page_text:
                    return {'total': 1, 'voted': 1, 'failed': 0}
                input("\n按 Enter 鍵繼續...\n")
                return None

        return None   # 無動作，讓迴圈繼續偵測

    def _handle_state_done(self, page_text: str) -> dict:
        """無可辨識操作元素時的終態處理"""
        submit_btns = self.driver.find_elements(By.XPATH, self._XPATH_SUBMIT)
        if submit_btns and any(b.is_displayed() for b in submit_btns):
            logger.info("✓ 發現提交/確認按鈕，投票完成")
            return {'total': 1, 'voted': 1, 'failed': 0}

        if self._has_dom_element(self._XPATH_NEXT):
            logger.info("ℹ️  還有『下一步』，嘗試點擊...")
            if self._retry_click(self._XPATH_NEXT):
                self._wait_page_ready()
                return None   # type: ignore  # 讓呼叫方 continue loop

        logger.debug("頁面內容片段: %s", page_text[:200])
        logger.info("✓ 投票流程完成")
        return {'total': 1, 'voted': 1, 'failed': 0}

    # ── 快速投票（主流程）────────────────────────────────────────
    # ⚠️  不加 @retry：此方法為 workflow，重試可能對已改變的 DOM 產生重複操作
    def _vote_with_agree_button(self):
        try:
            logger.info("=" * 60)
            logger.info("【投票流程】快速投票")

            # 等待任意操作元素出現後再進入迴圈
            _any_elem = f"{self._XPATH_VOTE_OPTS} | {self._XPATH_NEXT} | {self._XPATH_DIRECTOR}"
            self._wait_clickable(_any_elem, timeout=10)

            # 診斷：還在列表頁，等待跳轉
            if (self._has_dom_element("//*[contains(.,'未投票')]")
                    and self._has_dom_element("//*[contains(.,'投票狀況')]")
                    and not self._has_dom_element(self._XPATH_VOTE_OPTS)):
                logger.warning("似乎還在列表頁面，等待跳轉...")
                self._wait_clickable(self._XPATH_VOTE_OPTS, timeout=10)

            iteration    = 0
            max_iter     = 10
            last_page_text = ""
            repeat_count   = 0

            # ── FSM dispatch table ─────────────────────────────────
            _dispatch = {
                VoteState.CONFIRM: lambda pt: self._handle_state_confirm(),
                VoteState.VOTING:  lambda pt: self._handle_state_voting(pt),
                VoteState.DONE:    lambda pt: self._handle_state_done(pt),
            }

            while iteration < max_iter:
                iteration += 1
                self._wait_page_ready(timeout=5)
                page_text = self.driver.find_element(By.TAG_NAME, 'body').text

                logger.info("【投票流程 - 循環 %d】(state檢測中...)", iteration)
                logger.debug("頁面文字長度: %d 字符", len(page_text))

                # 重複頁面偵測
                if last_page_text and page_text[:100] == last_page_text[:100] and len(page_text) == len(last_page_text):
                    repeat_count += 1
                    logger.warning("偵測到重複頁面 (第 %d 次)", repeat_count)
                    if repeat_count >= 3:
                        logger.warning("重複頁面過多，跳出迴圈")
                        return {'total': 1, 'voted': 1, 'failed': 0}
                else:
                    repeat_count = 0
                last_page_text = page_text

                # 偵測狀態 → dispatch
                state = self._detect_state()
                logger.info("→ state=%s", state.value)
                # self.screenshot_handler.capture(f"state_{state.value}_iter{iteration}")

                handler = _dispatch.get(state)
                if handler is None:
                    logger.info("✓ 投票流程完成（無 handler）")
                    return {'total': 1, 'voted': 1, 'failed': 0}

                result = handler(page_text)
                if result is not None:
                    return result
                # result is None → continue loop

            logger.warning("超過最大循環次數，投票流程終止")
            return {'total': 1, 'voted': 1, 'failed': 0}

        except Exception as e:
            logger.exception("投票流程錯誤")
            self.screenshot_handler.save_error_screenshot("vote_with_agree_error")
            return {'total': 0, 'voted': 0, 'failed': 0}
    
    def _find_and_vote_all(self):
        try:
            logger.info("=" * 60)
            logger.info("【投票流程】搜尋未投票的議案")

            time.sleep(2)

            unvoted_items = self.page_navigator.find_unvoted_items()

            if not unvoted_items:
                logger.error("未找到任何未投票的議案")
                return {'total': 0, 'voted': 0, 'failed': 0}

            self.total_count = len(unvoted_items)
            logger.info("🔍 找到 %d 個未投票的議案", self.total_count)

            for idx, item in enumerate(unvoted_items, 1):
                logger.info("【議案 %d/%d】", idx, self.total_count)
                if self._vote_item(item, idx):
                    self.voted_count += 1
                time.sleep(0.5)

            result = {
                'total': self.total_count,
                'voted': self.voted_count,
                'failed': self.total_count - self.voted_count
            }

            logger.info("=" * 60)
            logger.info("【投票完成】總計: %d 個議案，已投票: %d 個，失敗: %d 個",
                        result['total'], result['voted'], result['failed'])
            return result

        except Exception as e:
            logger.exception("投票流程錯誤")
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
                            logger.info("✓ 已選擇「同意」")
                            return True
                else:
                    self.driver.execute_script("arguments[0].click();", agree_button)
                    time.sleep(0.5)
                    logger.info("✓ 已點擊「同意」按鈕")
                    return True

            if self._vote_item_fallback(item):
                logger.info("✓ 已選擇「同意」")
                return True

            logger.error("無法投票 (item_index=%d)", item_index)
            return False

        except Exception as e:
            logger.error("投票失敗 (代碼:105): %s", str(e))
            return False
    
    def _go_back_to_list(self):
        success, msg = self.page_navigator.go_back_to_list()
        return (0 if success else -1, msg)


    def execute_voting_loop(self, log_msg_func):
        # 檢測系統維護訊息
        if self._detect_maintenance_message():
            log_msg_func("⏸️  系統維護中，直接進入結束流程")
            self.is_maintenance = True
            return {
                'total': 0,
                'voted': 0,
                'failed': 0,
                'has_unvoted': False
            }
        
        self.is_maintenance = False  # 系統正常，非維護狀態
        
        total_voted = 0
        total_failed = 0
        total_scanned = 0
        
        log_msg_func("ℹ️  開始掃描未投票公司...")
        
        # 循環投票：每次投票完成後重新掃描，而非預先保存所有元素
        # 原因：每次投票返回列表時頁面會刷新，預先保存的元素引用會變成陳舊(stale)
        while True:
            # 重新掃描未投票的公司（使用 PageNavigator）
            unvoted_companies = self.page_navigator.find_all_unvoted_companies()
            logger.debug("掃描到 %d 家未投票公司", len(unvoted_companies))
            
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
                self.companies_info.append({
                    'code': '未知',
                    'name': '未知',
                    'status': '投票失敗'
                })
                continue
            
            # 在當前公司行中查找和點擊投票按鈕
            try:
                vote_button = company_row.find_element(By.XPATH, './/a[contains(.,"投票")] | .//button[contains(.,"投票")]')
                log_msg_func(f"【投票流程】點擊投票按鈕 (代碼: {company_code})")
                vote_button.click()
                # 等待投票頁面任何操作元素出現
                self._wait_clickable(
                    f"{self._XPATH_VOTE_OPTS} | {self._XPATH_NEXT} | {self._XPATH_DIRECTOR}",
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
                    self.companies_info.append({
                        'code': company_code,
                        'name': company_name,
                        'status': '投票失敗'
                    })
                    continue
            
            # 進行投票
            vote_result = self._vote_with_agree_button()
            
            if vote_result.get('failed', 0) > 0:
                log_msg_func("⚠️  快速投票失敗，進行逐個投票...")
                vote_result = self._find_and_vote_all()
            
            if vote_result.get('total', 0) > 0:
                log_msg_func("✓ 提交投票...")
                try:
                    # self.screenshot_handler.capture(f"before_submit_{company_code}")
                    code, msg = self.page_navigator.submit_vote()
                    # 等待確認按鈕（取代 time.sleep(1)）
                    self._wait_clickable(
                        "//button[contains(.,'確認')] | //a[contains(.,'確認')]",
                        timeout=5)
                    
                    # 點擊確認按鈕就會自動返回投票列表
                    code, msg = self.page_navigator.click_query_button()
                    # 等待列表頁面刷新
                    self._wait_clickable(
                        "//*[contains(.,'未投票')] | //*[contains(.,'已投票')]",
                        timeout=10)
                    
                    log_msg_func("✓ 投票完成，準備掃描下一個公司...")
                    total_voted += 1
                    self.companies_info.append({
                        'code': company_code,
                        'name': company_name,
                        'status': '已投票'
                    })
                
                except Exception as e:
                    log_msg_func(f"❌ 提交或確認失敗: {str(e)[:60]}")
                    self.screenshot_handler.save_error_screenshot(f"submit_fail_{company_code}")
                    total_failed += 1
                    self.companies_info.append({
                        'code': company_code,
                        'name': company_name,
                        'status': '投票失敗'
                    })
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
                self.companies_info.append({
                    'code': company_code,
                    'name': company_name,
                    'status': '投票失敗'
                })
            
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

