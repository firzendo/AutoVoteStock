#!/usr/bin/env python
# coding: utf-8

print("啟動中，請稍待...")
import os
import sys
import time
import datetime
import logging
from dotenv import load_dotenv

# 加入src目錄到Python path
script_dir = os.path.dirname(os.path.abspath(__file__))
src_path = os.path.join(script_dir, 'src')
if src_path not in sys.path:
    sys.path.insert(0, src_path)

from login_handler import LoginHandler
from vote_handler import VoteHandler
from screenshot_handler import ScreenshotHandler
from page_navigator import PageNavigator
from screenshot import execute_final_screenshot
from selenium import webdriver

# 載入.env配置
load_dotenv()

# pyinstaller -D .\股東e票通輔助工具.py 打包成exe
# python -m PyInstaller -D .\股東e票通輔助工具.py  打包成exe


# --- 1. 初始化 Log 系統 ---
LOG_DIR = "log"
if not os.path.exists(LOG_DIR):
    os.makedirs(LOG_DIR)

log_filename = os.path.join(LOG_DIR, datetime.datetime.now().strftime("%Y%m%d_%H%M%S.log"))
file_handler = logging.FileHandler(log_filename, encoding="utf-8")
stream_handler = logging.StreamHandler()

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[file_handler, stream_handler],
)

def log_msg(msg: str):
    """同時輸出到 console + log 檔"""
    logging.info(msg)


# --- 2. 初始化WebDriver ---
def setup_chrome_driver():
    """設置Chrome WebDriver"""
    options = webdriver.ChromeOptions()
    
    # 開啟Chrome無痕模式
    options.add_argument("--incognito")
    
    # 強制解決區域網路彈窗與通知
    prefs = {
        "profile.default_content_setting_values.local_network_access": 1,
        "profile.managed_default_content_settings.local_network_access": 1,
        "profile.default_content_setting_values.notifications": 2
    }
    options.add_experimental_option("prefs", prefs)
    
    # 移除自動化偵測特徵
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_argument("--log-level=3")
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option('useAutomationExtension', False)
    
    driver = webdriver.Chrome(options=options)
    return driver


# --- 3. 主程式 ---
def main():
    """主程式 - 執行登入和後續操作"""
    
    log_msg("=" * 60)
    log_msg("股東e票通輔助工具 - 啟動")
    log_msg("=" * 60)
    
    driver = None
    login_handler = None
    completed_successfully = False

    try:
        # 設置WebDriver
        driver = setup_chrome_driver()
        
        # 初始化登入處理器
        login_handler = LoginHandler(driver)

        # --- 步驟1：登入流程 ---
        log_msg("\n【步驟1】執行登入...")
        login_success, login_msg = login_handler.execute_login_flow(log_msg)
        if not login_success:
            raise Exception(f"登入失敗: {login_msg}")

        # --- 步驟2：循環投票流程 ---
        log_msg("\n【步驟2】進行循環投票...")
        
        # 在此處統一創建所有實體
        # 1. 創建 PageNavigator（頁面導航）
        page_navigator = PageNavigator(driver)
        
        # 2. 創建 ScreenshotHandler（截圖處理）
        screenshot_handler = ScreenshotHandler(driver, page_navigator, screenshot_dir="screenshots")
        
        # 3. 創建 VoteHandler（投票處理）
        vote_handler = VoteHandler(driver, page_navigator, screenshot_handler, screenshot_dir="screenshots")
        
        # ⚠️  注意：execute_voting_loop 中不再進行截圖（避免 stale element reference 錯誤）
        # 截圖將在投票完成後統一進行
        voting_stats = vote_handler.execute_voting_loop(log_msg)
        
        # --- 步驟2.5：投票完成後進行統一截圖 ---
        # 等待所有投票完成，頁面恢復穩定狀態
        log_msg("\n✓ 所有投票已完成，準備進行截圖...")
        time.sleep(2)  # 截圖前讓頁面穩定（此處保留 sleep 因為截圖不需要等待特定元素）
        execute_final_screenshot(driver, voting_stats, log_msg, vote_handler, page_navigator, output_dir="screenshots")

        completed_successfully = True

    except Exception as e:
        log_msg(f"❌ 程式錯誤: {str(e)}")
        import traceback
        log_msg(traceback.format_exc())
        input("\n⛔ 發生錯誤，程式已暫停（網頁維持現狀）\n   確認後按 Enter 鍵關閉瀏覽器...")

    finally:
        # 步驟3：登出（僅在正常完成時執行）
        if completed_successfully and driver and login_handler:
            log_msg("\n【步驟3】執行登出...")
            login_handler.logout(log_msg)

        if driver:
            driver.quit()
            log_msg("✓ 已關閉瀏覽器")

        log_msg("\n【完成】程式執行結束")
        log_msg("=" * 60)


if __name__ == "__main__":
    main()
