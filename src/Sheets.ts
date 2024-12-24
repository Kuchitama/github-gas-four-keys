import { PullRequest } from "../types/main";

type Values = any[];

export abstract class Sheet {
  abstract sheet: GoogleAppsScript.Spreadsheet.Sheet;
  public static sheetName: string;
  public static readonly defaultSheet: GoogleAppsScript.Spreadsheet.Spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  public static readonly existValueFilter = (val: any) => val !== null && val !== undefined && val !== "";

  protected getOrCreateSheet(sheetName: string): GoogleAppsScript.Spreadsheet.Sheet {
    return Sheet.defaultSheet.getSheetByName(sheetName) || Sheet.defaultSheet.insertSheet(sheetName);
  }

  /**
  * get vertical values array from range(Value[n][0]) 
  * 
  * @param colChar column index. e.g. 'A' 
  * @param opt Row range option. The range starts the first row if head is 0. The range ends the last row if last is 0. e.g. {head: 0, last: 0} 
  * @returns Values array of the range value
  */
  getVerticalValues(colChar, opt={head: 0, last:0}): Values {
    const head = opt.head || 0;
    const last = opt.last || 0;
    const start = head === 0 ? colChar : `${colChar}${head}`;
    const end = last <= 0 ? colChar : `${colChar}${last}`

    return this.sheet.getRange(`${start}:${end}`).getValues().map((vs, _) => vs[0]);
  }
}

export class SettingsSheet extends Sheet {
  public static readonly sheetName: string = "分析設定";
  readonly sheet: GoogleAppsScript.Spreadsheet.Sheet;

  constructor() {
    super();
    this.sheet = this.getOrCreateSheet(SettingsSheet.sheetName);
  }

  public initialize(pullRequestsSheetName: string, repos: string[]): void {
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

    this.sheet.getRange(1, 7, 1, 2)
      .setValues([
        ["リポジトリ", "最終更新日時"],
      ])
      .setBackgroundRGB(224, 224, 224);

    this.sheet.getRange(2, 7, repos.length, 2)
      .setValues(repos.map((repo, i) => [repo, `=IF(MAXIFS('プルリク情報'!K2:K, 'プルリク情報'!H2:H, G${i+2})=0, "", TEXT(MAXIFS('プルリク情報'!K2:K, 'プルリク情報'!H2:H, G${i+2}), "yyyy-mm-dd HH:MM:SS"))`]))
  }

  getLatestUpdatedAt(repo: string): Date | null {
    const repoIdx = this.getVerticalValues('G', {head: 2, last: 0}).findIndex((val, _) => val === repo);
    const dateStr = this.sheet.getRange(repoIdx + 2, 8).getValue();
    if (!Sheet.existValueFilter(dateStr)) {
      return null;
    }
    return new Date(dateStr);
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

export class PullRequestsSheet extends Sheet {
  public static readonly sheetName: string = "プルリク情報";
  readonly sheet: GoogleAppsScript.Spreadsheet.Sheet;

  constructor() {
    super();
    this.sheet = this.getOrCreateSheet(PullRequestsSheet.sheetName);
  }

  public initialize(): void {
    this.sheet.getRange(1, 1, 1, 12)
      .setValues([[
        "メンバー名",
        "ブランチ名",
        "PR本文",
        "マージ済",
        "初コミット日時",
        "マージ日時",
        "マージまでの時間(hours)",
        "リポジトリ",
        "障害発生判定",
        "障害対応PR",
        "更新日時",
        "id",
      ]])
      .setBackgroundRGB(224, 224, 224);   
  } 

  upsertPullRequest(repositoryName: string, pullRequest: PullRequest) {
    const row = this.findRowByPullRequest(repositoryName, pullRequest);

    this.writePullRequestData(pullRequest, row, repositoryName);
  }

  private writePullRequestData(pullRequest: PullRequest, row: number, repositoryName: string) {
    let firstCommitDate: Date | null = null;
    if (pullRequest.commits.nodes[0].commit.committedDate) {
      firstCommitDate = new Date(pullRequest.commits.nodes[0].commit.committedDate);
    }
    let mergedAt: Date | null = null;
    if (pullRequest.mergedAt) {
      mergedAt = new Date(pullRequest.mergedAt);
    }

    const updatedAt = new Date(pullRequest.updatedAt);
    this.sheet.getRange(row, 1).setValue(pullRequest.author.login);
    this.sheet.getRange(row, 2).setValue(pullRequest.headRefName);
    this.sheet.getRange(row, 3).setValue(pullRequest.bodyText);
    this.sheet.getRange(row, 4).setValue(pullRequest.merged);
    this.sheet.getRange(row, 5).setValue(!!firstCommitDate ? this.formatDate(firstCommitDate) : "");
    this.sheet.getRange(row, 6).setValue(!!mergedAt ? this.formatDate(mergedAt) : "");
    this.sheet.getRange(row, 7).setValue(
      (!!firstCommitDate && !!mergedAt) ?
        (mergedAt.getTime() - firstCommitDate.getTime()) / 60 / 60 / 1000 :
        "");
    this.sheet.getRange(row, 8).setValue(repositoryName);
    this.sheet.getRange(row, 9).setValue(`=REGEXMATCH(B${row}, '分析設定'!$E$3)`);
    this.sheet.getRange(row, 10).setValue(`=REGEXMATCH(B${row}, '分析設定'!$E$4)`);
    this.sheet.getRange(row, 11).setValue(this.formatDate(updatedAt));
    this.sheet.getRange(row, 12).setValue(`${repositoryName}/${pullRequest.number}`);
  }

  /**
   * 
   * @param repositoryName The name of the repository
   * @param pullRequest The pull request object
   * @returns The row number of the pull request. If the pull request is not found, it returns the least row number of empty.
   */
  private findRowByPullRequest(repositoryName: string, pullRequest: PullRequest): number {
    const prIds = this.getVerticalValues('L', {head: 2, last: 0}).filter(Sheet.existValueFilter);

    const index = prIds.findIndex((id, _) => id === `${repositoryName}/${pullRequest.number}`);
    return index > 0 ? index + 2: prIds.length + 2;
  }

  /**
   * 日付フォーマットヘルパ関数
   * 
   * @param date 
   * @returns formated string ${year}-${month}-${day} ${hour}:${minute}:${second} 
   */
  private formatDate(date: Date): string {
    const year = date.getFullYear();
    const month = (date.getMonth() + 1).toString().padStart(2, '0');
    const day = date.getDate().toString().padStart(2, '0');
    const hour = date.getHours().toString().padStart(2, '0');
    const minute = date.getMinutes().toString().padStart(2, '0');
    const second = date.getSeconds().toString().padStart(2, '0');
    return `${year}-${month}-${day} ${hour}:${minute}:${second}`;
  }
} 
