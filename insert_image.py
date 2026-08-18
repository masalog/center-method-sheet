import xlwings as xw
import sys

# --- 設定 ---
excel_path = r"C:\Users\dghy1\center-method-sheet\center_sheet_C-1-2.xlsx"
img_path = r"C:\Users\dghy1\center-method-sheet\images\image_001.png"

# --- 画像サイズ・位置（ポイント単位） ---
IMG_WIDTH = 267.7
IMG_HEIGHT = 642.5
IMG_LEFT = 187.6
IMG_TOP = 150.8

# --- Excelに挿入 ---
app = None
try:
    app = xw.App(visible=False)
    wb = app.books.open(excel_path)
    ws = wb.sheets["C1_2"]

    # 既存の画像のうち、同じ位置にあるものを削除（吹き出し画像は残す）
    for pic in list(ws.pictures):
        if abs(pic.left - IMG_LEFT) < 5 and abs(pic.top - IMG_TOP) < 5:
            name = pic.name
            pic.delete()
            print(f"既存画像を削除: {name}")

    # 新しい画像を挿入
    pic = ws.pictures.add(
        img_path,
        left=IMG_LEFT,
        top=IMG_TOP
    )
    pic.width = IMG_WIDTH
    pic.height = IMG_HEIGHT

    # 保存して閉じる
    wb.save()
    wb.close()
    app.quit()
    print("✅ 画像を差し替えました（既存の吹き出し図形は保持されています）")

except Exception as e:
    print(f"❌ エラー: {e}")
    if app:
        try:
            app.quit()
        except:
            pass
    sys.exit(1)
