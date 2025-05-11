[English README](README.en.md) | [中文說明](README.md)

# 股票資訊爬蟲

## 專案描述
這是一個自動化股票資訊抓取工具，使用Yahoo Finance API來獲取股票、ETF以及加密貨幣的價格和市值數據，並將結果存儲到Excel文件中。

## 功能
- 從Excel文件讀取股票代碼
- 使用Yahoo Finance API獲取當前股票價格和市值
- 智能識別資產類型（股票、ETF、加密貨幣、匯率）
- 根據資產類型選擇性抓取市值
- 將抓取的數據寫入Excel文件，並根據日期建立新列
- 支援外部設定檔，方便配置Excel文件路徑

## 環境需求
- Python 3.6+
- pandas
- yfinance
- openpyxl

## 安裝
1. 複製此程式碼儲存庫或下載原始碼
2. 安裝所需的依賴庫：
   ```
   pip install -r requirements.txt
   ```

## 使用方法
1. 確保您有一個正確格式的Excel文件，包含一個名為"Symbol"的列，其中包含所有股票代碼
2. 配置`config.ini`文件：
   ```ini
   [Excel]
   file_path = 您的Excel文件路徑.xlsx
   sheet_name = 您的工作表名稱
   
   [Output]
   verbose = True
   ```
3. 運行程式：
   ```
   python main.py
   ```
4. 程式將抓取每個股票的收盤價和市值（如果適用），並將結果寫入Excel文件

## 設定檔選項
### [Excel] 部分
- `file_path`: Excel文件的路徑
- `sheet_name`: 包含股票代碼的工作表名稱

### [Output] 部分
- `verbose`: 是否輸出詳細日誌 (True/False)

## 資產類型處理
- 股票（Equity/Stock）：抓取價格和市值
- 加密貨幣（Cryptocurrency）：抓取價格和市值
- ETF/共同基金：只抓取價格，不抓取市值
- 匯率（Currency）：只抓取價格，不抓取市值

## 注意事項
- 確保您的Excel文件在運行程式時未被其他程式開啟
- Yahoo Finance API可能會對過多的請求進行限制，請勿短時間內頻繁運行程式
- 某些股票代碼可能與Yahoo Finance不相容，請確保使用正確的代碼格式
