# 股東e票通登入模組使用說明

## 📋 概述
此模組提供自動化登入股東e票通的功能，身份證字號存儲在 `.env` 文件中。

## 🔧 設置步驟

### 1. 配置 .env 文件
編輯根目錄的 `.env` 文件，填入以下信息：

```env
# 身份證字號（例如：A123456789）
ID_NUMBER=YOUR_ID_NUMBER_HERE

# 登入頁面URL
LOGIN_URL=https://xxx.xxx.xxx/login

# 登入超時時間（秒）
LOGIN_TIMEOUT=30
```

### 2. 安裝依賴
確保已安裝所需的Python套件：

```bash
pip install selenium webdriver-manager python-dotenv
```

## 📦 模組結構

### src/login_handler.py
主要的登入處理模組，包含 `LoginHandler` 類別

**主要方法：**
- `login()` - 執行登入流程
- `get_id_number()` - 從.env讀取身份證號
- `set_id_number(id_number)` - 設置身份證號
- `verify_login_success()` - 驗證登入是否成功

## 🚀 使用方式

### 方式1：直接運行示例程式
```bash
python login_example.py
```

### 方式2：在自己的代碼中導入使用
```python
from src.login_handler import LoginHandler
from selenium import webdriver

# 建立WebDriver
driver = webdriver.Chrome()

# 初始化登入處理器
login_handler = LoginHandler(driver)

# 執行登入
if login_handler.login():
    print("登入成功")
    login_handler.verify_login_success()
else:
    print("登入失敗")

driver.quit()
```

## 🔍 自定義輸入框選擇器

如果系統找不到身份證輸入框，可以修改 `login_handler.py` 中的 `_find_id_input_field()` 方法，添加新的選擇器：

```python
def _find_id_input_field(self):
    selectors = [
        (By.ID, 'your_id_here'),
        (By.CSS_SELECTOR, 'input[name="yourname"]'),
        # 添加更多選擇器
    ]
```

## ⚠️ 注意事項

1. 身份證字號必須是10個字元（1個英文字母 + 9個數字）
2. 請勿將真實身份証號上傳至公開代碼行儲庫
3. `.env` 文件應該被添加到 `.gitignore` 中
4. 登入超時默認為30秒，可根據需要調整

## 🐛 故障排除

| 問題 | 解決方案 |
|------|--------|
| 找不到身份證輸入框 | 使用開發者工具檢查頁面元素，更新選擇器 |
| 登入超時 | 增加 LOGIN_TIMEOUT 的值 |
| Chrome驅動程式錯誤 | 執行 `pip install --upgrade webdriver-manager` |

## 📝 相關文件
- `.env` - 配置文件
- `src/login_handler.py` - 登入處理模組
- `login_example.py` - 登入示例程式
