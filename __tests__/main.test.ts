import { jest } from '@jest/globals';
import { PullRequest } from '../types/main';

const mockSpreadsheetApp: GoogleAppsScript.Spreadsheet.SpreadsheetApp = {
  getActiveSpreadsheet: jest.fn().mockReturnValue({
		getSheetByName: jest.fn(),
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
  getAllRepos,
} from '../main';

describe('getPullRequests', () => {
  beforeEach(() => {
    jest.clearAllMocks();
  });

  test('正常にPRを取得できる場合', async () => {
    const mockPRData: PullRequest = {
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

	test('100件以上のPRを取得する場合', async () => {
    const mockPRData: PullRequest[] = Array.from({length: 100}).map((_, i: number) => {
			const updatedDate = new Date()
			updatedDate.setDate(updatedDate.getDate() - i);

			return {
        number: 101 - i,
				author: { login: 'testUser' },
				headRefName: 'feature/test',
				bodyText: 'Test PR',
				merged: true,
				mergedAt: updatedDate.toISOString(),
				commits: {
					nodes: [{
						commit: {
							committedDate: updatedDate.toISOString()
						}
					}]
				},
        updatedAt: '2024-01-01T00:00:00Z',
			}
    });

    const mockResponse1 = {
      getContentText: () => JSON.stringify({
        data: {
          repository: {
            pullRequests: {
              pageInfo: {
                hasNextPage: true,
								endCursor: 100,
              },
              nodes: mockPRData
            }
          }
        }
      })
    };
		const mockResponse2 = {
      getContentText: () => JSON.stringify({
        data: {
          repository: {
            pullRequests: {
              pageInfo: {
                hasNextPage: false,
              },
              nodes: [{
								author: { login: 'testUser' },
								headRefName: 'feature/test',
								bodyText: 'Test PR',
								merged: true,
								mergedAt: new Date().toISOString(),
								commits: {
									nodes: [{
										commit: {
											committedDate: new Date().toISOString()
										}
									}]
								},
								updatedAt: new Date().toISOString(),
							}] 
            }
          }
        }
      })
    };


		const fetchFn = (mockUrlFetchApp.fetch as jest.Mock)
    fetchFn.mockReturnValueOnce(mockResponse1);
		fetchFn.mockReturnValueOnce(mockResponse2);

    const result = getPullRequests('test-repo');


		// Check the number of fetch calls
		expect(fetchFn.mock.calls).toHaveLength(2);

		expect((fetchFn.mock.calls[0][1] as any).payload).not.toContain('after:');
		expect((fetchFn.mock.calls[1][1] as any).payload).toContain('after: \\\"100\\\"');

    expect(result).toHaveLength(101);
    expect(result[0].author.login).toBe('testUser');
    expect(mockUrlFetchApp.fetch).toHaveBeenCalledWith(
      'https://api.github.com/graphql',
      expect.any(Object)
    );
  });

  test('updatedFrom以降のPRを返却する', () => {
    const mockResponse = {
      getContentText: () => JSON.stringify({
        data: {
          repository: {
            pullRequests: {
              pageInfo: {
                hasNextPage: false
              },
              nodes: [{
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
              }, {
                number: 2,
                author: { login: 'testUser' },
                headRefName: 'feature/test',
                bodyText: 'Test PR',
                merged: true,
                mergedAt: '2024-01-02T00:00:00Z',
                commits: {
                  nodes: [{
                    commit: {
                      committedDate: '2024-01-02T00:00:00Z'
                    }
                  }]
                },
                updatedAt: '2024-01-02T00:00:00Z',
              }] as PullRequest[]
            }
          }
        }
      })
    };

    const fetchFn = (mockUrlFetchApp.fetch as jest.Mock);
    fetchFn.mockReturnValue(mockResponse);

    const updatedFrom = new Date('2024-01-01T00:00:00Z');

    const result = getPullRequests('test-repo', updatedFrom);

    expect(result).toHaveLength(1);
    expect(result[0].author.login).toBe('testUser');
    expect(new Date(result[0].updatedAt).getTime()).toBeGreaterThan(updatedFrom.getTime());
  })

  test('100件以上のPRがあるが、updatedFromで取得件数を絞る場合', async () => {
    const mockPRData: PullRequest[] = Array.from({length: 100}).map((_, i: number) => {
			const updatedDate = new Date()
			updatedDate.setDate(updatedDate.getDate() - i);

			return {
        number: 101 - i,
				author: { login: 'testUser' },
				headRefName: 'feature/test',
				bodyText: 'Test PR',
				merged: true,
				mergedAt: updatedDate.toISOString(),
				commits: {
					nodes: [{
						commit: {
							committedDate: updatedDate.toISOString()
						}
					}]
				},
				updatedAt: updatedDate.toISOString(),
			}
    });

    const mockResponse1 = {
      getContentText: () => JSON.stringify({
        data: {
          repository: {
            pullRequests: {
              pageInfo: {
                hasNextPage: true, // There are more than 100 PRs
								endCursor: 100,
              },
              nodes: mockPRData
            }
          }
        }
      })
    };

		const fetchFn = (mockUrlFetchApp.fetch as jest.Mock)
    fetchFn.mockReturnValueOnce(mockResponse1);

    // get from 3days ago
    const updatedFrom = new Date();
    updatedFrom.setDate(updatedFrom.getDate() - 3);

    const result = getPullRequests('test-repo', updatedFrom);

		// Check the number of fetch calls
		expect(fetchFn.mock.calls).toHaveLength(1);
    expect(result).toHaveLength(3);
  });
});

describe('getAllRepos', () => {
  const mockSetValue = jest.fn();
  const mockValues = jest.fn().mockReturnValue([['2023-01-01T00:00:00Z','2023-12-31T00:00:00Z','2023-06-01T00:00:00Z']]);
  const mockSheet = {
      getRange: jest.fn().mockReturnValue({
        setValue: mockSetValue,
        getValues: mockValues,
      })
    };
  beforeEach(() => {
    jest.clearAllMocks();
  });

  test('全リポジトリのPRを取得してスプレッドシートに書き込む', () => {
    (mockSpreadsheetApp.getActiveSpreadsheet().getSheetByName as jest.Mock).mockReturnValue(mockSheet);

    const mockPRResponse = {
      getContentText: () => JSON.stringify({
        data: {
          repository: {
            pullRequests: {
              pageInfo: {
                hasNextPage: false
              },
              nodes: [{
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
              }]
            }
          }
        }
      })
    };

		const fetchFn = mockUrlFetchApp.fetch as jest.Mock;
    fetchFn.mockReturnValue(mockPRResponse);

    getAllRepos();

    expect(mockUrlFetchApp.fetch).toHaveBeenCalled();
		expect((fetchFn.mock.calls[0][1] as any).payload).toContain('repository(name: \\\"repo1\\\"');
		expect((fetchFn.mock.calls[1][1] as any).payload).toContain('repository(name: \\\"repo2\\\"');

    expect((mockSheet.getRange as jest.Mock).mock.calls.length).toBe(28);
    expect(mockSetValue.mock.calls.length).toBe(24);
  });

  test('PRの最終更新日時を考慮してPRを取り込む', () => {
    (mockSpreadsheetApp.getActiveSpreadsheet().getSheetByName as jest.Mock).mockReturnValue(mockSheet);

    const mockPRResponse = {
      getContentText: () => JSON.stringify({
        data: {
          repository: {
            pullRequests: {
              pageInfo: {
                hasNextPage: false
              },
              nodes: [{
                author: { login: 'testUser' },
                headRefName: 'feature/test',
                bodyText: 'Test PR',
                merged: true,
                mergedAt: '2024-01-02T00:00:00Z',
                commits: {
                  nodes: [{
                    commit: {
                      committedDate: '2024-01-02T00:00:00Z'
                    }
                  }]
                },
                updatedAt: '2024-01-02T00:00:00Z',
              }, {
                author: { login: 'testUser' },
                headRefName: 'feature/test',
                bodyText: 'Test PR',
                merged: true,
                mergedAt: '2023-01-01T00:00:00Z',
                commits: {
                  nodes: [{
                    commit: {
                      committedDate: '2023-01-01T00:00:00Z'
                    }
                  }]
                },
                updatedAt: '2023-01-01T00:00:00Z',
              }]
            }
          }
        }
      })
    };

		const fetchFn = mockUrlFetchApp.fetch as jest.Mock;
    fetchFn.mockReturnValue(mockPRResponse);

    getAllRepos();

    expect(mockUrlFetchApp.fetch).toHaveBeenCalled();
		expect((fetchFn.mock.calls[0][1] as any).payload).toContain('repository(name: \\\"repo1\\\"');

    // 最終更新日時='2024-01-01T00:00:00Z' として、それ以降の1件のみ更新する
    expect(mockSetValue.mock.calls.length).toBe(24);
  });
});
