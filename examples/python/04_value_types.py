# value-types.py — node 版サンプルと同一検証 (pyosbxl)
import json
import pyosbxl

app = pyosbxl.App()
wb = app.createWorkBook("value-types-py.xlsx")
ws = wb.openWorkSheet("Sheet1")

# JSON.stringify 相当で引用して 表示値/生値/書式 を 1 行で出す
def show(a1):
    r = ws.getRange(f"{a1}:{a1}")
    disp = json.dumps(r.getValue()[a1], ensure_ascii=False)
    raw = json.dumps(r.getValue(True)[a1], ensure_ascii=False)      # True で生値
    fmt = json.dumps(r.getNumberFormat()[a1], ensure_ascii=False)
    print(f"  {a1}: 表示値={disp} 生値={raw} 書式={fmt}")

def num(x):          # 整数値の float は int 表記にする（JS の出力と揃える）
    return int(x) if isinstance(x, float) and x.is_integer() else x

def boolstr(x):      # Python の True/False を JS 風の true/false で表示する
    return str(x).lower()

# --- 日付・時刻 ---
# DateTimeObject は引数なしで生成し、setYMD() / setHMS() で中身を設定します
date = pyosbxl.dto.DateTimeObject()
date.setYMD(2024, 1, 15)

dt = pyosbxl.dto.DateTimeObject()   # ※ 変数名 datetime は標準モジュール名を避けて dt
dt.setYMD(2024, 1, 15)
dt.setHMS(13, 45, 30)

time = pyosbxl.dto.DateTimeObject()
time.setHMS(9, 5, 0)

print("[日付・時刻]")
ws.getRange("A1:A1").setDateValue(date)                       # 日付のみ
ws.getRange("A2:A2").setDateValue(dt)                         # 日付 + 時刻
ws.getRange("A3:A3").setDateValue(time)                       # 時刻のみ
ws.getRange("A4:A4").setDateStringValue("2024-01-15")         # 文字列から
ws.getRange("A5:A5").setDateStringValue("令和6年1月15日")       # 和暦も認識されます
ws.getRange("A6:A6").setDateValue(date, False, "yyyy-mm-dd")   # 表示書式を同時に指定
ws.getRange("A7:A7").setDateValue(date, False, "ggge年m月d日") # 和暦+引用符なし日本語
for a in ["A1", "A2", "A3", "A4", "A5", "A6", "A7"]:
    show(a)

# --- 数値・文字列 ---
print("[数値・文字列]")
ws.getRange("B1:B1").setNumberValue(1234567.891)
ws.getRange("B2:B2").setNumberValue(1234567.891, False, "#,##0.00")
ws.getRange("B3:B3").setNumberValue(0.25, False, "0.0%")
ws.getRange("B4:B4").setValue("00123")       # 先頭 0 は文字列として保持される
ws.getRange("B5:B5").setValue("2024/1/15")   # 汎用 setter は日付として自動認識
ws.getRange("B6:B6").setValue("12345", True) # forceString=True で強制的に文字列扱い
ws.getRange("B7:B7").setNumberValue(1234.5, False, '"円"#,##0')      # 引用符リテラル
ws.getRange("B8:B8").setNumberValue(-1234.5, False, "#,##0;赤-#,##0") # 裸の色トークン（[赤] へ自動変換）
for b in ["B1", "B2", "B3", "B4", "B5", "B6", "B7", "B8"]:
    show(b)

# --- 真偽値 ---
print("[真偽値]")
ws.getRange("C1:C1").setBooleanValue(True)
ws.getRange("C2:C2").setBooleanValue(False)
ws.getRange("C3:C3").setFormula("TRUE")      # 数式でも真偽値になる（"=" は付けない）
for c in ["C1", "C2", "C3"]:
    show(c)

# --- 日付のシリアル値 ---
print("[シリアル値]")
print("  2024/1/15 のシリアル値 =", num(app.getNumericValue(date)))
print("  1904 年形式の場合      =", num(app.getNumericValue(date, True)))
print("  wb.isDate1904()        =", boolstr(wb.isDate1904()))

wb.save()
wb.close()