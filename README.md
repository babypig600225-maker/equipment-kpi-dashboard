# 富強醫材 設備維修 KPI 儀表板

## 功能說明
- 上傳「設備維修統計.xlsx」自動解析，顯示 MTBF / MTTR / 故障次數 / 維修工時 / 停工工時
- 支援機台系列總覽 & 個別機台兩個分頁
- 月份與系列篩選器

## 部署方式

### 方式 A：本機執行
```bash
pip install -r requirements.txt
streamlit run app.py
```
開啟瀏覽器 http://localhost:8501

### 方式 B：內網伺服器部署
```bash
# 安裝套件
pip install -r requirements.txt

# 指定 host 讓內網可存取
streamlit run app.py --server.address 0.0.0.0 --server.port 8501
```
其他人用 http://[伺服器IP]:8501 開啟

### 方式 C：Streamlit Cloud（公開）
1. 把此資料夾推送到 GitHub
2. 至 https://streamlit.io/cloud 登入並 Deploy
3. 選擇 app.py 即完成

## 每月更新流程
1. 準備新的「設備維修統計_1-X月_.xlsx」
2. 開啟儀表板網址
3. 點「↩ 重新上傳」，上傳新檔案
4. 自動更新完成

## 需求格式
Excel 需包含工作表：
- 「關鍵設備績效指標」（必要）
- 各月工作表（如「2026-1月」，用於取得報修件數與完成率）
