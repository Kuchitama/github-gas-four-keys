interface ISheet {
  readonly sheet: GoogleAppsScript.Spreadsheet.Sheet;
}

abstract class Sheet {
  public static readonly defaultSheet: GoogleAppsScript.Spreadsheet.Spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  protected getOrCreateSheet(sheetName: string): GoogleAppsScript.Spreadsheet.Sheet {
    return Sheet.defaultSheet.getSheetByName(sheetName) || Sheet.defaultSheet.insertSheet(sheetName);
  }
}

export class SettingsSheet extends Sheet implements ISheet {
  sheetName: string;
  readonly sheet: GoogleAppsScript.Spreadsheet.Sheet;

  constructor() {
    super();
    this.sheetName = "分析設定";
    this.sheet = this.getOrCreateSheet(this.sheetName);
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
