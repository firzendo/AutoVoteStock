#!/usr/bin/env python
# coding: utf-8

"""
登入模組 - 處理股東e票通登入流程
"""

import os
from dotenv import load_dotenv
import logging
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support import expected_conditions as EC
import time

# 載入.env配置
load_dotenv()


logger = logging.getLogger(__name__)

class LoginHandler:
    """登入處理器"""
    
    def __init__(self, driver: webdriver.Chrome):
        """
        初始化登入處理器
        
        Args:
            driver: Selenium WebDriver 實例
        """
        self.driver = driver
        self.id_number = os.getenv('ID_NUMBER', '')
        self.login_timeout = int(os.getenv('LOGIN_TIMEOUT', '30'))
        self.login_url = os.getenv('LOGIN_URL', '')
        self.cert_type = os.getenv('CERT_TYPE', '券商網路')
    
    def _get_id_number(self):
        """從.env取得身份證字號"""
        return self.id_number
    
    def _set_id_number(self, id_number: str):
        """設置身份證字號"""
        if len(id_number) != 10:
            raise ValueError(f"身份證字號長度應為10個字元，收到: {len(id_number)}")
        self.id_number = id_number
    
    def _login(self):
        """
        執行登入流程
        
        Returns:
            tuple: (code, msg) - 0成功，<0為錯誤代碼
        """
        if not self.id_number:
            return (-200, "身份證字號未設置，請先在.env中設置ID_NUMBER")
        
        # 導向至登入頁面
        if self.login_url:
            logger.info("正在導入登入頁面: %s", self.login_url)
            self.driver.get(self.login_url)
            time.sleep(3)
        
        # 尋找身份證字號輸入框
        code, result = self._find_id_input_field()
        if code != 0:
            logger.error("❌ 無法自動找到輸入框")
            logger.info("👉 手動操作模式已啟動")
            logger.info("請在瀏覽器中手動進行以下操作:")
            logger.info("1. 輸入身份證號")
            logger.info("2. 選擇憑證種類: 券商網路")
            logger.info("3. 點擊登入按鈕")
            logger.info("定位元素提示:")
            logger.info("- 打開開發者工具 (F12)")
            logger.info("- 檢查身份證輸入框的屬性 (id, name, placeholder)")
            logger.info("- 或在下方輸入框找到元素說明")
            input("🔔 請完成登入，然後按 Enter 鍵繼續...\n")
            return (0, "手動登入完成")
        
        id_input = result
        id_input.clear()
        id_input.send_keys(self.id_number)
        logger.info("✓ 已輸入身份證字號: %s****", self.id_number[:4])
        
        # 選擇憑證種類
        if self.cert_type:
            code, msg = self._select_cert_type()
            if code == 0:
                logger.info("✓ 已選擇憑證種類: %s", self.cert_type)
            else:
                logger.warning("⚠️  無法自動選擇憑證種類")
        
        # 嘗試找到並點擊登入/確認按鈕
        code, msg = self._click_login_button()
        if code == 0:
            logger.info("✓ 已點擊登入按鈕")
            time.sleep(2)
        
        # 處理登入後可能出現的權限請求對話框
        logger.info("🔍 檢查是否有權限請求對話框...")
        code, msg = self._handle_permission_dialog()
        if code == 0:
            logger.info("✓ %s", msg)
            time.sleep(2)
        
        return (0, "登入流程完成")
    
    def _click_login_button(self):
        """
        尋找並點擊登入/確認按鈕
        
        Returns:
            tuple: (code, msg) - 0成功，-201為錯誤
        """
        button_selectors = [
            (By.CSS_SELECTOR, 'button[type="submit"]'),
            (By.XPATH, '//button[contains(text(), "登入")]'),
            (By.XPATH, '//button[contains(text(), "確認")]'),
            (By.CSS_SELECTOR, 'input[type="submit"]'),
        ]
        
        for by, value in button_selectors:
            buttons = self.driver.find_elements(by, value)
            for button in buttons:
                button.click()
                return (0, "已點擊登入按鈕")
        
        raise Exception("無法找到登入按鈕")
    
    def _handle_permission_dialog(self):
        """
        處理登入後的權限請求對話框
        （例如：存取這部裝置上的其他應用程式和服務）
        
        Returns:
            tuple: (code, msg) - 0成功，-205為沒有對話框（不是錯誤）
        """
        try:
            # 尋找「允許」相關的按鈕
            allow_button_selectors = [
                # 權限對話框的「允許」按鈕
                (By.XPATH, '//button[contains(text(), "允許")]'),
                (By.XPATH, '//a[contains(text(), "允許")]'),
                
                # 通用確認按鈕
                (By.XPATH, '//button[contains(text(), "確認")]'),
                (By.XPATH, '//button[contains(text(), "同意")]'),
                (By.XPATH, '//button[@class="btn-primary" or contains(@class, "allow")]'),
                
                # 對話框確認按鈕
                (By.CSS_SELECTOR, 'button.permission-allow'),
                (By.CSS_SELECTOR, 'button[onclick*="allow"]'),
            ]
            
            for by, value in allow_button_selectors:
                buttons = self.driver.find_elements(by, value)
                for button in buttons:
                    if button.is_displayed():
                        button_text = button.text.strip() or button.get_attribute('title') or "按鈕"
                        logger.info("📍 找到權限按鈕: %s", button_text)
                        try:
                            button.click()
                            logger.info("✓ 已點擊權限允許按鈕")
                            return (0, "已處理權限請求")
                        except Exception as e:
                            # 嘗試 JavaScript 點擊
                            logger.warning("⚠️  直接點擊失敗，嘗試 JavaScript...")
                            self.driver.execute_script("arguments[0].click();", button)
                            logger.info("✓ 已使用 JavaScript 點擊權限按鈕")
                            return (0, "已使用 JavaScript 處理權限請求")
            
            # 如果沒有找到按鈕，不視為錯誤
            logger.info("ℹ️  未發現權限請求對話框（正常）")
            return (0, "無權限對話框")
        
        except Exception as e:
            logger.warning("⚠️  處理權限對話框時出錯: %s", str(e)[:50])
            return (0, "權限對話框處理完畢")
    
    def _select_cert_type(self):
        """
        選擇憑證種類
        
        Returns:
            tuple: (code, msg) - 0成功，-202為錯誤
        """
        logger.info("🔍 正在尋找憑證種類選擇器（目標: %s）...", self.cert_type)
        time.sleep(1)
        
        # 方式1：Radio buttons
        radios = self.driver.find_elements(By.XPATH, '//input[@type="radio"]')
        for radio in radios:
            if not radio.is_selected():
                radio.click()
                time.sleep(0.3)
                if radio.is_selected():
                    logger.info("✓ 已選擇")
                    return (0, "已選擇憑證")
        
        # 方式2：Select dropdown
        selects = self.driver.find_elements(By.TAG_NAME, 'select')
        for dropdown in selects:
            if dropdown.is_displayed():
                select = Select(dropdown)
                for option in select.options:
                    if self.cert_type in option.text or option.text in self.cert_type:
                        option.click()
                        logger.info("✓ 已選擇: %s", option.text)
                        return (0, f"已選擇: {option.text}")
        
        # 方式3：Buttons
        buttons = self.driver.find_elements(By.XPATH, f'//button[contains(text(), "{self.cert_type}")]')
        for button in buttons:
            if button.is_displayed():
                button.click()
                logger.info("✓ 已點擊按鈕")
                return (0, "已點擊憑證按鈕")
        
        raise Exception("無法自動選擇憑證種類")

    def _find_id_input_field(self):
        """
        尋找身份證輸入框
        
        Returns:
            tuple: (code, element or msg) - 0成功返回元素，-203為錯誤
        """
        # 尋找 iframe
        iframes = self.driver.find_elements(By.TAG_NAME, 'iframe')
        logger.info("🔍 發現 %s 個 iframe", len(iframes))
        
        for idx, iframe in enumerate(iframes):
            self.driver.switch_to.frame(iframe)
            code, element = self._search_input_in_current_frame()
            if code == 0:
                logger.info("✓ 在 iframe [%s] 中找到輸入框", idx)
                return (0, element)
            self.driver.switch_to.default_content()
        
        # 在主頁面中尋找
        logger.info("🔍 在主頁面中尋找輸入框...")
        code, element = self._search_input_in_current_frame()
        if code == 0:
            logger.info("✓ 在主頁面中找到輸入框")
            return (0, element)
        
        raise Exception("無法找到身份證輸入框")
    
    def _search_input_in_current_frame(self):
        """
        在當前 frame 中搜尋輸入框
        
        Returns:
            tuple: (code, element or msg) - 0成功返回元素，非0為找不到
        """
        selectors = [
            (By.ID, 'idno'),
            (By.ID, 'idNo'),
            (By.ID, 'userId'),
            (By.CSS_SELECTOR, 'input[type="text"]'),
            (By.XPATH, '//input[@type="text" and @required]'),
        ]
        
        for by, value in selectors:
            elements = self.driver.find_elements(by, value)
            for elem in elements:
                if elem.is_displayed():
                    logger.info("✓ 找到輸入框")
                    return (0, elem)
        
        return (-1, "未找到輸入框")
    
    def _verify_login_success(self):
        """
        驗證登入成功
        
        Returns:
            tuple: (code, msg) - 0成功，-204為錯誤
        """
        time.sleep(2)
        
        current_url = self.driver.current_url
        logger.info("當前頁面URL: %s", current_url)
        
        if "login" not in current_url.lower():
            logger.info("✓ 已離開登入頁面")
            return (0, "登入成功")
        
        error_elements = self.driver.find_elements(By.CSS_SELECTOR, ".alert-danger, .error")
        if error_elements:
            logger.error("❌ 登入失敗: %s", error_elements[0].text[:50])
            raise Exception(f"登入失敗: {error_elements[0].text[:50]}")
        
        logger.warning("⚠️  無法確定登入狀態")
        return (0, "登入狀態未知，繼續進行")
    
    def execute_login_flow(self, log_msg_func):
        """
        執行完整的登入流程（包括讀取 ID、登入、驗證）
        
        Args:
            log_msg_func: 用於輸出log的函數
        
        Returns:
            tuple: (success, msg) - (True/False, 消息文本)
                - 成功: (True, "登入成功")
                - 失敗: (False, 錯誤訊息)
        """
        try:
            # 1. 讀取身份證號
            id_number = self._get_id_number()
            if not id_number:
                log_msg_func("⚠️  .env中未設置ID_NUMBER，請先設置身份證字號")
                return (False, "未設置身份證")
            
            log_msg_func(f"✓ 已讀取身份證號: {id_number[:4]}****")
            
            # 2. 執行登入
            code, msg = self._login()
            if code != 0:
                log_msg_func(f"✗ 登入失敗 (代碼:{code}): {msg}")
                return (False, msg)
            
            log_msg_func("✓ 身份證號輸入成功")
            
            # 3. 驗證登入
            code, msg = self._verify_login_success()
            if code != 0:
                log_msg_func(f"✗ 登入驗證失敗 (代碼:{code}): {msg}")
                return (False, msg)
            
            log_msg_func("✓ 登入成功")
            return (True, "登入成功")
            
        except Exception as e:
            error_msg = f"登入流程異常: {str(e)}"
            log_msg_func(f"❌ {error_msg}")
            return (False, error_msg)

    def logout(self, log_msg_func=None):
        """
        執行登出流程（點擊左上角登出按鈕，並確認對話框）
        
        Returns:
            tuple: (code, msg) - 0成功，<0為錯誤代碼
        """
        def _log(msg):
            if log_msg_func:
                log_msg_func(msg)

        try:
            _log("   尋找登出按鈕...")

            # 嘗試點擊頁面上的登出按鈕（左上角）
            logout_selectors = [
                (By.XPATH, '//*[contains(text(), "登出") and not(contains(text(), "登出後"))]'),
                (By.CSS_SELECTOR, 'a[href*="logout"]'),
                (By.CSS_SELECTOR, 'button[onclick*="logout"]'),
            ]

            clicked = False
            for by, selector in logout_selectors:
                try:
                    elements = self.driver.find_elements(by, selector)
                    for el in elements:
                        if el.is_displayed():
                            el.click()
                            clicked = True
                            _log("   ✓ 已點擊登出按鈕")
                            break
                    if clicked:
                        break
                except Exception:
                    pass

            if not clicked:
                # Fallback：直接導航到登出頁
                _log("   ⚠️ 找不到登出按鈕，使用URL導航登出")
                self.driver.get("https://stockservices.tdcc.com.tw/evote/logout.html")

            time.sleep(1)

            # 處理 JavaScript alert 確認框
            try:
                alert = self.driver.switch_to.alert
                alert.accept()
                _log("   ✓ 已確認登出 alert")
                time.sleep(1)
            except Exception:
                pass

            # 處理頁面內彈出的確認按鈕
            confirm_selectors = [
                (By.ID, 'comfirmDialog_okBtn'),
                (By.XPATH, '//button[contains(text(), "確認") or contains(text(), "確定") or contains(text(), "OK")]'),
                (By.XPATH, '//a[contains(text(), "確認") or contains(text(), "確定")]'),
            ]
            for by, selector in confirm_selectors:
                try:
                    els = self.driver.find_elements(by, selector)
                    for el in els:
                        if el.is_displayed():
                            el.click()
                            _log("   ✓ 已確認登出對話框")
                            time.sleep(1)
                            break
                except Exception:
                    pass

            _log("✓ 登出完成")
            return (0, "登出成功")

        except Exception as e:
            _log(f"⚠️ 登出時發生例外: {str(e)}")
            return (-1, f"登出失敗: {str(e)}")

