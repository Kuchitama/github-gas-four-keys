import { jest } from '@jest/globals';

const mockSpreadsheetApp: GoogleAppsScript.Spreadsheet.SpreadsheetApp = {
  getActiveSpreadsheet: jest.fn().mockReturnValue({
    getSheetByName: jest.fn(),
    insertSheet: jest.fn(),
  }),
  newConditionalFormatRule: jest.fn().mockReturnValue({
    whenTextEqualTo: jest.fn().mockReturnThis(),
    setBackground: jest.fn().mockReturnThis(),
    setRanges: jest.fn().mockReturnThis(),
    build: jest.fn()
  }),
} as unknown as GoogleAppsScript.Spreadsheet.SpreadsheetApp;

const mockUtilities = {
  formatDate: jest.fn(),
} as unknown as GoogleAppsScript.Utilities.Utilities

const mockCharts = {
  ChartType: jest.fn().mockReturnValue({
    COMBO: jest.fn(),
  }),
} as unknown as GoogleAppsScript.Charts.Charts

global.SpreadsheetApp = mockSpreadsheetApp;
global.Utilities = mockUtilities;
global.Charts = mockCharts;

import { FourKeysSheet } from '../src/Sheets';

describe('FourKeysSheet', () => {
  let fourKeysSheet: FourKeysSheet;
  let mockSheet: GoogleAppsScript.Spreadsheet.Sheet;

  const mockGetRange = {
    setValues: jest.fn().mockReturnThis(),
    setBackgroundRGB: jest.fn().mockReturnThis(),
    setNumberFormats: jest.fn().mockReturnThis(),
    copyTo: jest.fn().mockReturnThis(),
  }
  beforeEach(() => {
    jest.clearAllMocks();

    mockSheet = {
      getRange: jest.fn().mockReturnValue(mockGetRange),
      newChart: jest.fn().mockReturnValue({
        addRange: jest.fn().mockReturnThis(),
        setChartType: jest.fn().mockReturnThis(),
        setNumHeaders: jest.fn().mockReturnThis(),
        setOption: jest.fn().mockReturnThis(),
        setPosition: jest.fn().mockReturnThis(),
        build: jest.fn()
      }),
      insertChart: jest.fn(),
      getConditionalFormatRules: jest.fn().mockReturnValue({
        push: jest.fn(),
      }),
      setConditionalFormatRules: jest.fn(),
    } as unknown as GoogleAppsScript.Spreadsheet.Sheet;
    (mockSpreadsheetApp.getActiveSpreadsheet().getSheetByName as jest.Mock).mockReturnValue(mockSheet);
    fourKeysSheet = new FourKeysSheet();
  });

  test('should initialize the FourKeys sheet correctly', () => {
    fourKeysSheet.initialize('PullRequests');

    expect(mockSheet.getRange).toHaveBeenCalledWith(1, 1, 2, 9);
    expect(mockGetRange.setValues).toHaveBeenCalledWith([
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
      ]
    ]);
    expect(mockGetRange.setBackgroundRGB).toHaveBeenCalled();
    expect(mockSheet.getRange(3, 2, 1, 8).setValues).toHaveBeenCalled();
    expect(mockSheet.getRange(3, 2, 1, 8).setNumberFormats).toHaveBeenCalled();
    expect(mockSheet.getRange(3, 2, 1, 8).copyTo).toHaveBeenCalled();
    expect(mockSheet.getRange(3, 1, 5, 1).setValues).toHaveBeenCalled();
    expect(mockSheet.getRange(8, 1, 1, 1).setValues).toHaveBeenCalled();

    expect(mockSheet.newChart().addRange).toHaveBeenCalled();
    expect(mockSheet.newChart().setChartType).toHaveBeenCalled();
    expect(mockSheet.newChart().setNumHeaders).toHaveBeenCalled();
    expect(mockSheet.newChart().setOption).toHaveBeenCalled();
    expect(mockSheet.newChart().setPosition).toHaveBeenCalled();
    expect(mockSheet.insertChart).toHaveBeenCalled();
  });
});
