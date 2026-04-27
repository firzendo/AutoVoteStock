#!/usr/bin/env python
# coding: utf-8
import os
import io
import time
import datetime

from PIL import Image
from selenium.webdriver.common.by import By
from page_navigator import PageNavigator


def execute_final_screenshot(driver, voting_stats, log_msg_func, vote_handler, page_navigator, output_dir="screenshots"):

    # 投票統計
    log_msg_func("\n" + "=" * 60)
    log_msg_func("【投票完成】統計結果")
    log_msg_func(f"  總計: {voting_stats['total']} 個公司")
    log_msg_func(f"  ✓ 已投票: {voting_stats['voted']} 個")
    log_msg_func(f"  ✗ 失敗: {voting_stats['failed']} 個")

    if not voting_stats.get('has_unvoted', True):
        log_msg_func("  ℹ️  無未投票公司（所有公司已投票）")

    log_msg_func("=" * 60)

    # 建立截圖回調函數（傳入 page_navigator）
    screenshot_func = create_company_screenshot_callback(driver, log_msg_func, page_navigator, output_dir=output_dir)

    # 使用 ScreenshotHandler 進行多頁分頁截圖
    vote_handler.screenshot_handler.screenshot_all_companies_results(log_msg_func, screenshot_func)

    log_msg_func("【截圖完成】所有投票結果已保存到 screenshots/ 目錄")
    # log_msg_func(f"文件名格式: 日期_名稱_代號.png")


def create_company_screenshot_callback(driver, log_msg_func, page_navigator, output_dir="screenshots"):
    """
    創建一個公司投票結果截圖回調函數
    截圖範圍：從「貴股東對」開始，到「最近一次投票時間」結束
    
    Returns:
        function: 接收 company_code 和 company_name 進行截圖的函數
    """
    def screenshot_callback(company_code, company_name):
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)

        # 檢查是否已有該公司代碼的截圖
        for fname in os.listdir(output_dir):
            if fname.endswith('.png') and f"_{company_code}.png" in fname:
                log_msg_func(f"   ✓ 已有截圖: {fname}，跳過")
                return

        timestamp = datetime.datetime.now().strftime("%Y%m%d")
        filename = os.path.join(output_dir, f"{timestamp}_{company_name}_{company_code}.png")

        try:
            # 使用 PageNavigator 滾回頁面頂部
            page_navigator.scroll_to_top()

            # 取得頁面設備像素比（高 DPI）
            device_pixel_ratio = page_navigator.get_device_pixel_ratio()

            # 尋找裁剪範圍
            # crop_top: 從「貴股東對」開始
            crop_top = 0
            try:
                title_elem = driver.find_element(By.XPATH, "//*[contains(text(), '貴股東對')]")
                crop_top = driver.execute_script("return arguments[0].getBoundingClientRect().top + window.scrollY;", title_elem)
            except Exception:
                log_msg_func(f"   ⚠️  無法定位標題，使用預設頂部")
                crop_top = 100

            # crop_bottom: 到「最近一次投票時間」結束
            crop_bottom = None
            try:
                time_elem = driver.find_element(By.XPATH, "//*[contains(text(), '最近一次投票時間')]")
                # 取得該元素的下邊界
                rect = driver.execute_script(
                    "const r = arguments[0].getBoundingClientRect(); return {bottom: r.bottom + window.scrollY};",
                    time_elem
                )
                crop_bottom = rect['bottom'] + 20  # 多加 20px 的空白
            except Exception:
                log_msg_func(f"   ⚠️  無法定位投票時間，使用全頁高度")
                crop_bottom = None

            # 全頁截圖存到記憶體
            png_bytes = driver.get_screenshot_as_png()
            img = Image.open(io.BytesIO(png_bytes))
            img_w, img_h = img.size

            # 換算像素座標（考慮 devicePixelRatio）
            top_px = int(crop_top * device_pixel_ratio)
            bottom_px = int(crop_bottom * device_pixel_ratio) if crop_bottom else img_h

            # 確保範圍合法
            top_px = max(0, min(top_px, img_h - 1))
            bottom_px = max(top_px + 10, min(bottom_px, img_h))

            # 裁剪並儲存
            cropped = img.crop((0, top_px, img_w, bottom_px))
            cropped.save(filename)

            log_msg_func(f"   ✓ 已保存截圖: {filename}")
        except Exception as e:
            log_msg_func(f"   ❌ 截圖失敗: {str(e)}")

    return screenshot_callback
