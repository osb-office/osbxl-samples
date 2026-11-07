// conditional-format.js — 条件付き書式ルールの追加・読み戻し・削除
const { nodeosbxl, enums, dto } = require("nodeosbxl");

const app = new nodeosbxl.App();
const wb = app.createWorkBook("condfmt.xlsx");
const ws = wb.openWorkSheet("Sheet1");

// 色ヘルパー: ColorObject は "#" なしの 6 桁 hex で作る
const color = (hex) => {
    const c = new dto.ColorObject();
    c.setHexColor(hex);
    return c;
};

// ---------- データ表 ----------
// B列: セル値ルール / C列: カラースケール / D列: データバー / E列: 上位3件・平均比較
["月", "セル値ルール", "カラースケール", "データバー", "上位/平均"].forEach((h, i) => {
    const col = "ABCDE"[i];
    ws.getRange(`${col}1:${col}1`).setValue(h);
});
const sales = [45, 78, 23, 91, 56, 34, 87, 12, 65, 49];
sales.forEach((v, i) => {
    ws.getRange(`A${i + 2}:A${i + 2}`).setNumberValue(i + 1);
    for (const col of ["B", "C", "D", "E"]) {
        ws.getRange(`${col}${i + 2}:${col}${i + 2}`).setNumberValue(v);
    }
});

const fc = ws.getFormatConditions();

// ---------- (1) セル値ルール: 60 より大きい → 濃い赤の太字 + 薄赤の塗りつぶし ----------
const cellVal = new dto.CellValueObject();
cellVal.setValueCondition(60, enums.XlFormatConditionOperator.FormatConditionGreater);
cellVal.setA1C1("B2:B11");
const redFont = new dto.FontObject();
redFont.setBold(true);
redFont.setColorObject(color("9C0006"));
cellVal.setFont(redFont);
const pinkFill = new dto.PatternFillObject();
pinkFill.setPattern(enums.XlPatternType.PatternSolid);
pinkFill.setCellColor(color("FFC7CE"));
cellVal.setPatternFill(pinkFill);
fc.setCellValueCondition(1, cellVal);

// ---------- (2) カラースケール: C列に 3 色のグラデーション ----------
const colorScale = new dto.ColorScaleObject();
colorScale.setA1C1("C2:C11");
colorScale.setMinimumCondition(color("F8696B"), enums.XlConditionValueType.ConditionValueTypeMinimum);
colorScale.setMediumCondition(color("FFEB84"), enums.XlConditionValueType.ConditionValueTypePercentile, 50);
colorScale.setMaximumCondition(color("63BE7B"), enums.XlConditionValueType.ConditionValueTypeMaximum);
fc.setColorScaleFormatCondition(2, colorScale);

// ---------- (3) データバー: D列 ----------
const dataBar = new dto.DataBarObject();
dataBar.setA1C1("D2:D11");
dataBar.setDataBarColor(color("638EC6"));
dataBar.setMinimumCondition(enums.XlConditionValueType.ConditionValueTypeMinimum);
dataBar.setMaximumCondition(enums.XlConditionValueType.ConditionValueTypeMaximum);
fc.setDataBarFormatCondition(3, dataBar);

// ---------- (4) 上位 3 件 + 平均より上: E列 ----------
const top3 = new dto.Top10Object();
top3.setRank(3, true);                       // true = 上位（false なら下位）
top3.setA1C1("E2:E11");
const blueFont = new dto.FontObject();
blueFont.setBold(true);
blueFont.setColorObject(color("0070C0"));
top3.setFont(blueFont);
fc.setTop10FormatCondition(4, top3);

const aboveAvg = new dto.AboveAverageObject();
aboveAvg.setAboveBelow(enums.XlAboveBelow.XlAboveAverage);
aboveAvg.setA1C1("E2:E11");
const greenFont = new dto.FontObject();
greenFont.setBold(true);
greenFont.setColorObject(color("006100"));
aboveAvg.setFont(greenFont);
fc.setAboveAverageCondition(5, aboveAvg);

fc.apply();                                  // 全条件の設定後に apply() が必須
console.log("[1] 5 件のルールを登録し apply() しました");

// ---------- ルール一覧の読み戻し ----------
const typeName = Object.fromEntries(
    Object.entries(enums.XlFormatConditionType).map(([name, val]) => [val, name]));
console.log(`[2] ルール数 = ${Object.keys(fc.getConditionTypes()).length}`);
for (const [prio, ranges] of Object.entries(fc.getConditionTypes())) {
    for (const [range, type] of Object.entries(ranges)) {
        console.log(`    priority ${prio}: ${range} → ${typeName[type]} (type=${type})`);
    }
}

// ---------- 各ルールの内容の読み戻し ----------
console.log("[3] ルール内容の読み戻し");
const g1 = fc.getCellValueCondition(1);
console.log(`    セル値  : ${g1.getA1C1()} 値=${g1.getConditionValue()}` +
    ` 演算子=${g1.getConditionOperator()} (Greater=${enums.XlFormatConditionOperator.FormatConditionGreater})` +
    ` 文字色=${g1.getFont().getColorObject().getHexColor()} 太字=${g1.getFont().isBold()}` +
    ` 背景=${g1.getPatternFill().getCellColor().getHexColor()}`);
const g2 = fc.getColorScaleFormatCondition(2);
console.log(`    カラースケール: min=${g2.getMinimumColor().getHexColor()} (type=${g2.getMinimumValueType()})` +
    ` mid=${g2.getMediumColor().getHexColor()} (type=${g2.getMediumValueType()}, 閾値=${g2.getMediumValue()})` +
    ` max=${g2.getMaximumColor().getHexColor()} (type=${g2.getMaximumValueType()})`);
const g3 = fc.getDataBarFormatCondition(3);
console.log(`    データバー: 色=${g3.getDataBarColor().getHexColor()}` +
    ` minType=${g3.getMinimumCondition()} maxType=${g3.getMaximumCondition()}`);
const g4 = fc.getTop10FormatCondition(4);
console.log(`    上位 N 件: rank=${g4.getRank()} isTop=${g4.getTopBottom()}`);
const g5 = fc.getAboveAverageCondition(5);
console.log(`    平均比較: aboveBelow=${g5.getAboveBelow()}` +
    ` (XlAboveAverage=${enums.XlAboveBelow.XlAboveAverage})`);

// ---------- 条件を満たして書式が付いたセル（apply() 後に取得可能）----------
const byRow = (a, b) => Number(a.slice(1)) - Number(b.slice(1));
console.log("[4] 書式が適用されたセル");
console.log(`    B列 (>60)     : ${Object.keys(fc.getAppliedFontObject("B2:B11")).sort(byRow).join(" ")}`);
console.log(`    E列 (上位/平均): ${Object.keys(fc.getAppliedFontObject("E2:E11")).sort(byRow).join(" ")}`);
const appliedBar = fc.getAppliedDataBar("D2:D11");
console.log(`    D列 データバー : ${Object.values(appliedBar).filter(Boolean).length} セル`);

// ---------- 削除 ----------
fc.removeCondition(4);                       // 上位 3 件のルールだけ削除
console.log(`[5] removeCondition(4) 後のルール数 = ${Object.keys(fc.getConditionTypes()).length}`);
fc.removeAllCondition();                     // 全ルールを削除
console.log(`[6] removeAllCondition() 後のルール数 = ${Object.keys(fc.getConditionTypes()).length}`);

// セル値ルールを登録し直して保存（保存ファイルを開いたとき効果が見えるように）
fc.setCellValueCondition(1, cellVal);
fc.apply();

wb.save();
wb.close();
console.log("[done] condfmt.xlsx を保存しました（セル値ルールを再登録済み）");