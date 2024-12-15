import { jest } from '@jest/globals';

const mockSpreadsheetApp: GoogleAppsScript.Spreadsheet.SpreadsheetApp = {
  getActiveSpreadsheet: jest.fn().mockReturnValue({
    getSheetByName: jest.fn(),
    insertSheet: jest.fn(),
  }),
} as unknown as GoogleAppsScript.Spreadsheet.SpreadsheetApp;

global.SpreadsheetApp = mockSpreadsheetApp;

import { PullRequestsSheet } from '../src/Sheets';

describe('PullRequestsSheet', () => {
  let pullRequestsSheet: PullRequestsSheet;
  let mockSheet: GoogleAppsScript.Spreadsheet.Sheet;

  const mockGetRange = {
    setValues: jest.fn().mockReturnThis(),
    setValue: jest.fn().mockReturnThis(),
    getValues: jest.fn().mockReturnValue([
      ['1'], [], [],
    ]),
    setBackgroundRGB: jest.fn().mockReturnThis(),
  }

  beforeEach(() => {
    jest.clearAllMocks();
    process.env.TZ = 'UTC';
    mockSheet = {
      getRange: jest.fn().mockReturnValue(mockGetRange),
    } as unknown as GoogleAppsScript.Spreadsheet.Sheet;

    (mockSpreadsheetApp.getActiveSpreadsheet().getSheetByName as jest.Mock).mockReturnValue(mockSheet);
    (mockSpreadsheetApp.getActiveSpreadsheet().insertSheet as jest.Mock).mockReturnValue(mockSheet);

    pullRequestsSheet = new PullRequestsSheet();
  });

  test('should initialize the sheet with correct headers', () => {
    pullRequestsSheet.initialize();

    expect(mockSheet.getRange).toHaveBeenCalledWith(1, 1, 1, 12);
    expect(mockGetRange.setValues).toHaveBeenCalled();
  });

  test('should upsert pull request data correctly', () => {
    const mockPullRequest = {
      number: 1,
      author: { login: 'testUser' },
      headRefName: 'feature/test',
      bodyText: 'Test PR',
      merged: true,
      mergedAt: '2024-01-01T00:00:00Z',
      commits: {
        nodes: [{
          commit: {
            committedDate: '2024-01-01T00:00:00Z'
          }
        }]
      },
      updatedAt: '2024-01-01T00:00:00Z',
    };

    pullRequestsSheet.upsertPullRequest('test-repo', mockPullRequest);

    expect(mockSheet.getRange).toHaveBeenCalledTimes(13);
    expect(mockGetRange.setValue).toHaveBeenCalledTimes(12);
  });
});
