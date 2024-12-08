import { jest } from '@jest/globals';
import { PullRequest } from '../types/main';

const mockSpreadsheetApp: GoogleAppsScript.Spreadsheet.SpreadsheetApp = {
  getActiveSpreadsheet: jest.fn().mockReturnValue({
  }),
  newConditionalFormatRule: jest.fn().mockReturnValue({
    whenTextEqualTo: jest.fn().mockReturnThis(),
    setBackground: jest.fn().mockReturnThis(),
    setRanges: jest.fn().mockReturnThis(),
    build: jest.fn()
  }),
  WeekDay: {
    MONDAY: 1
  }
} as unknown as GoogleAppsScript.Spreadsheet.SpreadsheetApp;

const mockUrlFetchApp: GoogleAppsScript.URL_Fetch.UrlFetchApp = {
	fetch: jest.fn()
} as unknown as GoogleAppsScript.URL_Fetch.UrlFetchApp;


const mockProperty = {
	deleteAllProperties: jest.fn(),
	deleteProperty: jest.fn(),
	getKeys: jest.fn(),
	getProperties: jest.fn(),
	getProperty: jest.fn(),
	setProperties: jest.fn(),
	setProperty: jest.fn(),
} as GoogleAppsScript.Properties.Properties;

const mockPropertiesService = {
	getScriptProperties: jest.fn().mockReturnValue({
		getProperty: jest.fn((key: string) => {
			const properties: { [key: string]: string } = {
				GITHUB_REPO_NAMES: '["repo1", "repo2"]',
				GITHUB_REPO_OWNER: 'testOwner',
				GITHUB_API_TOKEN: 'test-token'
			};
			return properties[key] || null;
		})
	}),
	getUserProperties: jest.fn(),
	getDocumentProperties: jest.fn(),
} as unknown as GoogleAppsScript.Properties.PropertiesService;

global.SpreadsheetApp = mockSpreadsheetApp;
global.PropertiesService = mockPropertiesService;
global.UrlFetchApp = mockUrlFetchApp;


// テスト対象のコードをインポート
import {
  getPullRequests,
//   getAllRepos,
} from '../main';

describe('getPullRequests', () => {
  beforeEach(() => {
    jest.clearAllMocks();
  });

  test('正常にPRを取得できる場合', async () => {
    const mockPRData: PullRequest = {
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
      }
    };

    const mockResponse = {
      getContentText: () => JSON.stringify({
        data: {
          repository: {
            pullRequests: {
              pageInfo: {
                hasNextPage: false
              },
              nodes: [mockPRData]
            }
          }
        }
      })
    };

    (mockUrlFetchApp.fetch as jest.Mock).mockReturnValue(mockResponse);

    const result = getPullRequests('test-repo');

    expect(result).toHaveLength(1);
    expect(result[0].author.login).toBe('testUser');
    expect(mockUrlFetchApp.fetch).toHaveBeenCalledWith(
      'https://api.github.com/graphql',
      expect.any(Object)
    );
  });
});

// describe('getAllRepos', () => {
//   beforeEach(() => {
//     jest.clearAllMocks();
//   });

//   test('全リポジトリのPRを取得してスプレッドシートに書き込む', () => {
//     const mockSheet = {
//       getRange: jest.fn().mockReturnValue({
//         setValue: jest.fn()
//       })
//     };
//     (mockSpreadsheetApp.getActiveSpreadsheet().getSheetByName as jest.Mock).mockReturnValue(mockSheet);

//     const mockPRResponse = {
//       getContentText: () => JSON.stringify({
//         data: {
//           repository: {
//             pullRequests: {
//               pageInfo: {
//                 hasNextPage: false
//               },
//               nodes: [{
//                 author: { login: 'testUser' },
//                 headRefName: 'feature/test',
//                 bodyText: 'Test PR',
//                 merged: true,
//                 mergedAt: '2024-01-01T00:00:00Z',
//                 commits: {
//                   nodes: [{
//                     commit: {
//                       committedDate: '2024-01-01T00:00:00Z'
//                     }
//                   }]
//                 }
//               }]
//             }
//           }
//         }
//       })
//     };

//     (mockUrlFetchApp.fetch as jest.Mock).mockReturnValue(mockPRResponse);

//     getAllRepos();

//     expect(mockSheet.getRange).toHaveBeenCalled();
//     expect(mockUrlFetchApp.fetch).toHaveBeenCalled();
//   });
// });
