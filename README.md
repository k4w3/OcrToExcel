# MACアドレス自動記録ツール（OCR）

印刷されたMACアドレスラベルをPCのインカメラで読み取り、Excelファイルに1行ずつ自動追記するツールです。  

---

## 主な機能

- MACアドレスをOCRで自動認識
- 認識結果をExcelファイル（mac_list.xlsx）に記録
- 重複チェック機能
- 認識成功後、即時Excelに追記

---

## 動作環境

- Windows 10/11（※Mac/Linux は未確認）
- Python 3.8 以上
- Webカメラまたは内蔵カメラ

---

## Tesseract OCR のインストール

下記からインストーラーをダウンロード  
https://github.com/UB-Mannheim/tesseract/wiki

---

## ライブラリのインストール

```
pip install opencv-python pytesseract openpyxl
```

---

## 実行方法

```
python .\main3.py
