#!/usr/bin/env python
# coding: utf-8
import os
import io
import time
import datetime

from PIL import Image
from selenium.webdriver.common.by import By


def execute_final_screenshot(driver, voting_stats, log_msg_func, vote_handler, output_dir="screenshots"):
    """
    投票完成後，逐一查詢各公司投票結果並截圖

    Args:
        driver: Selenium WebDriver 實例
        voting_stats: dict 投票統計 {'total': 總數, 'voted': 已投票, 'failed': 失敗, 'has_unvoted': 是否找到未投票}
        log_msg_func: 用於輸出log的函數
        vote_handler: VoteHandler 實例，用於操作頁面
        output_dir: 截圖保存目錄
    """
    # 投票統計
    log_msg_func("\n" + "=" * 60)
    log_msg_func("【投票完成】統計結果")
    log_msg_func(f"  總計: {voting_stats['total']} 個公司")
    log_msg_func(f"  ✓ 已投票: {voting_stats['voted']} 個")
    log_msg_func(f"  ✗ 失敗: {voting_stats['failed']} 個")

    if not voting_stats.get('has_unvoted', True):
        log_msg_func("  ℹ️  無未投票公司（所有公司已投票）")

    log_msg_func("=" * 60)

    # 建立截圖回調函數
    screenshot_func = create_company_screenshot_callback(driver, log_msg_func, output_dir=output_dir)

    # 逐一查詢並截圖
    vote_handler.screenshot_all_companies_results(log_msg_func, screenshot_func)

    log_msg_func("【截圖完成】所有投票結果已保存到 screenshots/ 目錄")
    log_msg_func(f"文件名格式: 日期_名稱_代號.png")


def create_company_screenshot_callback(driver, log_msg_func, output_dir="screenshots"):
    """
    創建一個公司投票結果截圖回調函數
    截圖範圍：排除頁面 header 及「報告事項」以下的內容
    
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
            # 捲回頁面頂部
            driver.execute_script("window.scrollTo(0, 0);")
            time.sleep(0.5)

            # 取得頁面縮放比例（高 DPI）
            device_pixel_ratio = driver.execute_script("return window.devicePixelRatio;") or 1

            # 取得 header 底部 Y（截圖從此開始）
            crop_top = 0
            for selector in [
                'div[class="c-header_pageInfo"]',
                'header',
                '.c-header',
            ]:
                try:
                    el = driver.find_element(By.CSS_SELECTOR, selector)
                    bottom = el.location["y"] + el.size["height"]
                    if bottom > 0:
                        crop_top = bottom
                        break
                except Exception:
                    pass

            # 取得「報告事項」的頂部 Y（截圖到此結束）
            crop_bottom = None
            try:
                # 找含「報告事項」文字的元素
                els = driver.find_elements(By.XPATH, '//*[contains(text(), "報告事項")]')
                for el in els:
                    if el.is_displayed():
                        crop_bottom = el.location["y"]
                        break
            except Exception:
                pass

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

            log_msg_func(f"   ✓ 已保存截圖: {filename}  (裁剪 y={crop_top}~{crop_bottom})")
        except Exception as e:
            log_msg_func(f"   ❌ 截圖失敗: {str(e)}")

    return screenshot_callback
