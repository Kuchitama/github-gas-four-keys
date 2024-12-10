// interface ISheet {
//   readonly sheet: GoogleAppsScript.Spreadsheet.Sheet;
//   readonly sheetName: string;
// }

abstract class Sheet {
  abstract sheet: GoogleAppsScript.Spreadsheet.Sheet;
  public static sheetName: string;
  public static readonly defaultSheet: GoogleAppsScript.Spreadsheet.Spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  protected getOrCreateSheet(sheetName: string): GoogleAppsScript.Spreadsheet.Sheet {
    return Sheet.defaultSheet.getSheetByName(sheetName) || Sheet.defaultSheet.insertSheet(sheetName);
  }
}

export class SettingsSheet extends Sheet {
  public static readonly sheetName: string = "分析設定";
  readonly sheet: GoogleAppsScript.Spreadsheet.Sheet;

  constructor() {
    super();
    this.sheet = this.getOrCreateSheet(SettingsSheet.sheetName);
  }

  public initialize(pullRequestsSheetName: string): void {
    this.sheet.getRange(1, 1, 2, 2)
      .setValues([
        ["集計メンバー", "集計可否(TRUEかFALSEを入力してください)"],
        [`=unique('${pullRequestsSheetName}'!A2:A100000000)`, ""]
        ]);

    this.sheet.getRange(1, 1, 1, 2)
      .setBackgroundRGB(224, 224, 224);

    this.sheet.getRange(1, 4, 1, 2)
      .setValues([["キー名", "値"]])
      .setBackgroundRGB(224, 224, 224);

    this.sheet.getRange(2, 4, 15, 2)
      .setValues([
        ["移動平均のウィンドウ幅(日)", "28"],
        ["修復/巻き直しPRのブランチ名の検索ルール(正規表現)", "hotfix|revert"],
        ["障害修復PRのブランチ名の検索ルール(正規表現)", "hotfix"],
        ["デプロイ頻度のElite判定条件(1日あたり平均N回以上PRがマージされていればElite)", "0.4285714286"],
        ["デプロイ頻度のHigh判定条件(1日あたり平均N回以上PRがマージされていればHigh)", "0.1428571429"],
        ["デプロイ頻度のMedium判定条件(1日あたり平均N回以上PRがマージされていればMedium)", "0.03333333333"],
        ["変更リードタイムのElite判定条件(修復/巻き戻しを除くPRが初コミットからマージされるまで平均N時間以内ならElite)", "24"],
        ["変更リードタイムのHigh判定条件(修復/巻き戻しを除くPRが初コミットからマージされるまで平均N時間以内ならHigh)", "168"],
        ["変更リードタイムのMedium判定条件(修復/巻き戻しを除くPRが初コミットからマージされるまで平均N時間以内ならMedium)", "720"],
        ["変更障害率のElite判定条件(全PRのうち修復/巻き戻しPRの割合がN%以下ならElite)", "0.15"],
        ["変更障害率のHigh判定条件(全PRのうち修復/巻き戻しPRの割合がN%以下ならHigh)", "0.3"],
        ["変更障害率のMedium判定条件(全PRのうち修復/巻き戻しPRの割合がN%以下ならMedium)", "0.45"],
        ["平均修復時間のElite判定条件(修復PRが初コミットからマージされるまで平均N時間以内ならElite)", "24"],
        ["平均修復時間のHigh判定条件(修復PRが初コミットからマージされるまで平均N時間以内ならHigh)", "168"],
        ["平均修復時間のMeidum定条件(修復PRが初コミットからマージされるまで平均N時間以内ならMedium", "720"],
      ]);
  }
}

export class FourKeysSheet extends Sheet {
  public static readonly sheetName: string = "FourKeys計測結果";
  readonly sheet: GoogleAppsScript.Spreadsheet.Sheet;

  constructor() {
    super();
    this.sheet = this.getOrCreateSheet(FourKeysSheet.sheetName);
  }

  public initialize(pullRequestsSheetName: string): void {
    this.sheet.getRange(1, 1, 2, 9)
      .setValues([
        [
          "集計日",
          "デプロイ頻度", "",
          "変更のリードタイム", "",
          "変更失敗率", "",
          "平均復旧時間", ""
        ],
        [
          "",
          "回数", "ランク",
          "時間(hours)", "ランク",
          "割合(%)", "ランク",
          "時間(hours)", "ランク"
        ]])
      .setBackgroundRGB(
        224, 224, 224
      );
    const formatRanges = [
      this.sheet.getRange("C3:C1000"),
      this.sheet.getRange("E3:E1000"),
      this.sheet.getRange("G3:G1000"),
      this.sheet.getRange("I3:I1000")
    ];
    const rules = this.sheet.getConditionalFormatRules();
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo("Elite")
      .setBackground("#b7e1cd")
      .setRanges(formatRanges)
      .build());
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo("High")
      .setBackground("#c9daf8")
      .setRanges(formatRanges)
      .build());
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo("Low")
      .setBackground("#f4cccc")
      .setRanges(formatRanges)
      .build());
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo("Medium")
      .setBackground("#fff2cc")
      .setRanges(formatRanges)
      .build());
    this.sheet.setConditionalFormatRules(rules);

    this.sheet.getRange(3, 2, 1, 8)
      .setValues([[
        `=SUM(MAP('${SettingsSheet.sheetName}'!B$2:B$1000, '${SettingsSheet.sheetName}'!A$2:A$1000, LAMBDA(b, a, IF(OR(b<>FALSE, ISBLANK(b)), COUNTIFS('${pullRequestsSheetName}'!F$2:F$100000000, ">=" & A3-'${SettingsSheet.sheetName}'!E$2,'${pullRequestsSheetName}'!F$2:F$100000000, "<" & A3, '${pullRequestsSheetName}'!I$2:I$100000000, FALSE, '${pullRequestsSheetName}'!A$2:A$100000000, a), 0))))/'${SettingsSheet.sheetName}'!E$2`,
        `=IFS(B3>='${SettingsSheet.sheetName}'!E$5, "Elite", B3>='${SettingsSheet.sheetName}'!E$6, "High", B3>='${SettingsSheet.sheetName}'!E$7, "Medium", TRUE, "Low")`,
        `=IF(B3 > 0, SUM(MAP('${SettingsSheet.sheetName}'!B$2:B$1000, '${SettingsSheet.sheetName}'!A$2:A$1000, LAMBDA(b, a, IF(OR(b<>FALSE, ISBLANK(b)), SUMIFS('${pullRequestsSheetName}'!G$2:G$100000000, '${pullRequestsSheetName}'!F$2:F$100000000, ">=" & A3-'${SettingsSheet.sheetName}'!E$2,'${pullRequestsSheetName}'!F$2:F$100000000, "<" & A3, '${pullRequestsSheetName}'!I$2:I$100000000, FALSE, '${pullRequestsSheetName}'!A$2:A$100000000, a), 0)))) / (B3*'${SettingsSheet.sheetName}'!E$2), 0)`,
        `=IFS(D3<='${SettingsSheet.sheetName}'!E$8, "Elite", D3<='${SettingsSheet.sheetName}'!E$9, "High", D3<='${SettingsSheet.sheetName}'!E$10, "Medium", TRUE, "Low")`,
        `=IF(B3 > 0, SUM(MAP(${SettingsSheet.sheetName}!B$2:B$1000, ${SettingsSheet.sheetName}!A$2:A$1000, LAMBDA(b, a, IF(OR(b<>FALSE, ISBLANK(b)), COUNTIFS('${pullRequestsSheetName}'!F$2:F$100000000, ">=" & A3-'${SettingsSheet.sheetName}'!E$2,'${pullRequestsSheetName}'!F$2:F$100000000, "<" & A3, '${pullRequestsSheetName}'!I$2:I$100000000, TRUE, '${pullRequestsSheetName}'!A$2:A$100000000, a), 0))))/(B3*'${SettingsSheet.sheetName}'!E$2), 0)`,
        `=IFS(F3<='${SettingsSheet.sheetName}'!E$11, "Elite", F3<='${SettingsSheet.sheetName}'!E$12, "High", F3<='${SettingsSheet.sheetName}'!E$13, "Medium", TRUE, "Low")`,
        `=IF(SUM(MAP(${SettingsSheet.sheetName}!B$2:B$1000, ${SettingsSheet.sheetName}!A$2:A$1000, LAMBDA(b, a, IF(OR(b<>FALSE, ISBLANK(b)), COUNTIFS('${pullRequestsSheetName}'!F$2:F$100000000, ">=" & A3-'${SettingsSheet.sheetName}'!E$2,'${pullRequestsSheetName}'!F$2:F$100000000, "<" & A3, '${pullRequestsSheetName}'!J$2:J$100000000, TRUE, '${pullRequestsSheetName}'!A$2:A$100000000, a))))) > 0, SUM(MAP(${SettingsSheet.sheetName}!B$2:B$1000, ${SettingsSheet.sheetName}!A$2:A$1000, LAMBDA(b, a, IF(OR(b<>FALSE, ISBLANK(b)), SUMIFS('${pullRequestsSheetName}'!G$2:G$100000000, '${pullRequestsSheetName}'!F$2:F$100000000, ">=" & A3-'${SettingsSheet.sheetName}'!E$2,'${pullRequestsSheetName}'!F$2:F$100000000, "<" & A3, '${pullRequestsSheetName}'!J$2:J$100000000, TRUE, '${pullRequestsSheetName}'!A$2:A$100000000, a)))))/SUM(MAP(${SettingsSheet.sheetName}!B$2:B$1000, ${SettingsSheet.sheetName}!A$2:A$1000, LAMBDA(b, a, IF(OR(b<>FALSE, ISBLANK(b)), COUNTIFS('${pullRequestsSheetName}'!F$2:F$100000000, ">=" & A3-'${SettingsSheet.sheetName}'!E$2,'${pullRequestsSheetName}'!F$2:F$100000000, "<" & A3, '${pullRequestsSheetName}'!J$2:J$100000000, TRUE, '${pullRequestsSheetName}'!A$2:A$100000000, a))))), 0)`,
        `=IFS(H3<='${SettingsSheet.sheetName}'!E$14, "Elite", H3<='${SettingsSheet.sheetName}'!E$15, "High", H3<='${SettingsSheet.sheetName}'!E$16, "Medium", TRUE, "Low")`
      ]])
      .setNumberFormats([[
        "#,##0.00", "@",
        "#,##0.00", "@",
        "0.00%", "@",
        "#,##0.00", "@",
      ]]);
    this.sheet.getRange(3, 2, 1, 8).copyTo(
      this.sheet.getRange(4, 2, 4, 8)
    );
    const today = new Date();
    this.sheet.getRange(3, 1, 5, 1).setValues(
      [4,3,2,1,0].map((numberOfTwoWeek) => [Utilities.formatDate(new Date(new Date().setDate(today.getDate() - 14*numberOfTwoWeek)), "JST", "yyyy-MM-dd")])
    );
    this.sheet.getRange(8, 1, 1, 1).setValues([
      ["移行のデータの統計値を取得する場合はB~I列を上からペーストしA列は任意の値を入力てください。"]
    ]);

    const dfLtChart = this.sheet.newChart()
      .addRange(this.sheet.getRange("A1:A1000"))
      .addRange(this.sheet.getRange("B1:B1000"))
      .addRange(this.sheet.getRange("D1:D1000"))
      .setChartType(Charts.ChartType.COMBO)
      .setNumHeaders(1)
      .setOption("title", "デプロイ頻度と変更リードタイム")
      .setOption("series",[
        { targetAxisIndex: 0, legend: "デプロイ頻度(1日あたり)"},
        { targetAxisIndex: 1, legend: "変更リードタイム"},
      ])
      .setOption("vAxes", [
        {
          title: "回数",
          minValue: 0
        },
        {
          title: "時間",
          minValue: 0
        },
      ])
      .setOption("hAxes", {
        title: "Week"
      })
      .setPosition(1, 10, 0, 0)
      .build();
    this.sheet.insertChart(dfLtChart);

    const incidentChart = this.sheet.newChart()
      .addRange(this.sheet.getRange("A1:A1000"))
      .addRange(this.sheet.getRange("F1:F1000"))
      .addRange(this.sheet.getRange("H1:H1000"))
      .setChartType(Charts.ChartType.COMBO)
      .setNumHeaders(1)
      .setOption("title", "変更障害率と平均復旧時間")
      .setOption("series",[
        { targetAxisIndex: 0, legend: "変更障害率"},
        { targetAxisIndex: 1, legend: "平均修復時間"},
      ])
      .setOption("vAxes", [
        {
          title: "％",
          minValue: 0
        },
        {
          title: "時間",
          minValue: 0
        },
      ])
      .setOption("hAxes", {
        title: "Week"
      })
      .setPosition(20, 10, 0, 0)
      .build();
    this.sheet.insertChart(incidentChart);
  }

}
