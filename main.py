import cv2
import pytesseract
import re
from openpyxl import load_workbook
from datetime import datetime

# pytesseractのパス（必要な場合）
# pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

excel_file = "mac_list.xlsx"
wb = load_workbook(excel_file)
ws = wb.active

# MACアドレスの正規表現
mac_pattern = r"([0-9A-Fa-f]{2}[:-]){5}([0-9A-Fa-f]{2})"

cap = cv2.VideoCapture(0)
print("カメラ起動中。MACアドレスが印刷されたラベルをかざしてください。")

while True:
    ret, frame = cap.read()
    if not ret:
        break

    # OCRで文字抽出
    gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
    text = pytesseract.image_to_string(gray)

    # MACアドレスっぽい文字列を検索
    macs = re.findall(mac_pattern, text)
    macs = [''.join(m[:-1]) + m[-1] for m in macs]  # タプルをMAC文字列に復元

    for mac in macs:
        mac = mac.upper()
        print(f"認識されたMAC: {mac}")

        existing_macs = [cell.value for cell in ws['A'] if cell.value]
        if mac in existing_macs:
            print("→ 重複しています。スキップします。")
            continue

        ws.append([mac, datetime.now().strftime("%Y-%m-%d %H:%M:%S")])
        wb.save(excel_file)
        print("→ Excelに追記しました。")

    cv2.imshow("MAC Scanner (OCR)", frame)
    if cv2.waitKey(1) == ord('q'):
        break

cap.release()
cv2.destroyAllWindows()
