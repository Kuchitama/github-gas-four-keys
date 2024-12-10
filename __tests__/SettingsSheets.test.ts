import { jest } from '@jest/globals';

const mockSpreadsheetApp: GoogleAppsScript.Spreadsheet.SpreadsheetApp = {
  getActiveSpreadsheet: jest.fn().mockReturnValue({
    getSheetByName: jest.fn(),
    insertSheet: jest.fn(),
  }),
} as unknown as GoogleAppsScript.Spreadsheet.SpreadsheetApp;

const mockGetRange = {
  setValues: jest.fn().mockReturnThis(),
  setBackgroundRGB: jest.fn().mockReturnThis(),
}

const mockSheet = {
  getRange: jest.fn().mockReturnValue(mockGetRange),
} as unknown as GoogleAppsScript.Spreadsheet.Sheet;

global.SpreadsheetApp = mockSpreadsheetApp;

import { SettingsSheet } from '../src/Sheets';

describe('SettingsSheet', () => {
  beforeEach(() => {
    jest.clearAllMocks();
  });

  test('should create or get the sheet on initialization', () => {
    (mockSpreadsheetApp.getActiveSpreadsheet().getSheetByName as jest.Mock).mockReturnValue(mockSheet);

    const settingsSheet = new SettingsSheet();

    expect(mockSpreadsheetApp.getActiveSpreadsheet().getSheetByName).toHaveBeenCalledWith('分析設定');
    expect(mockSpreadsheetApp.getActiveSpreadsheet().insertSheet).not.toHaveBeenCalled();
    expect(settingsSheet.sheet).toBe(mockSheet);
  });

  test('should insert a new sheet if it does not exist', () => {
    (mockSpreadsheetApp.getActiveSpreadsheet().getSheetByName as jest.Mock).mockReturnValue(null);
    (mockSpreadsheetApp.getActiveSpreadsheet().insertSheet as jest.Mock).mockReturnValue(mockSheet);

    const settingsSheet = new SettingsSheet();

    expect(mockSpreadsheetApp.getActiveSpreadsheet().getSheetByName).toHaveBeenCalledWith('分析設定');
    expect(mockSpreadsheetApp.getActiveSpreadsheet().insertSheet).toHaveBeenCalledWith('分析設定');
    expect(settingsSheet.sheet).toBe(mockSheet);
  });

  test('should initialize the sheet with correct values', () => {
    (mockSpreadsheetApp.getActiveSpreadsheet().getSheetByName as jest.Mock).mockReturnValue(mockSheet);

    const settingsSheet = new SettingsSheet();
    settingsSheet.initialize('PullRequests');

    expect(mockGetRange.setValues).toHaveBeenCalled();
  });
});
