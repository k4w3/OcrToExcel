# MACアドレス自動記録ツール（OCR）

印刷されたMACアドレスラベルをPCのインカメラで読み取り、Excelファイルに1行ずつ自動追記するツールです。  

---

## 主な機能

- 印刷されたMACアドレスをOCRで自動認識
- 認識結果をExcelファイル（mac_list.xlsx）に記録
- 重複チェック機能
- 認識成功後、即時Excelに追記

---

## 動作環境

- Python 3.8 以上
- インカメラ

---

## インストール

```
pip install opencv-python pytesseract openpyxl
