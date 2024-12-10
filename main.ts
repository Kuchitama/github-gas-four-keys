import { PullRequest } from "./types/main";
import { FourKeysSheet, SettingsSheet } from "./src/Sheets";

const githubEndpoint = "https://api.github.com/graphql";

const repositoryNames = JSON.parse(PropertiesService.getScriptProperties().getProperty("GITHUB_REPO_NAMES") || "");
const repositoryOwner = PropertiesService.getScriptProperties().getProperty("GITHUB_REPO_OWNER");
const githubAPIKey = PropertiesService.getScriptProperties().getProperty("GITHUB_API_TOKEN");

const defaultSheet = SpreadsheetApp.getActiveSpreadsheet();

function initialize() {

  const pullRequestsSheetName = "プルリク情報";
  const pullRequestsSheet = getOrCreateSheet(pullRequestsSheetName);
  pullRequestsSheet.getRange(1, 1, 1, 12)
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

  new SettingsSheet().initialize(pullRequestsSheetName);
  new FourKeysSheet().initialize(pullRequestsSheetName);
  
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === "getAllRepos").forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger("getAllRepos")
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(0)
    .create();
}

function getAllRepos() {

  let i = 0;
  repositoryNames.forEach((repositoryName) => {
    // ToDo: Get latest updatedAt
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(`プルリク情報`);
    if (sheet === null) {
      throw new Error("Cannot find プルリク情報")
    }

    const updates = getVerticalValues(sheet, 'K', { head: 2 , last: 0});

    const latestUpdated = (updates.length > 0) ? new Date(updates.sort((a, b) => new Date(b).getTime() - new Date(a).getTime())[0]) : null; 
    if ( latestUpdated === null ) {
      console.log(`get all PRs`);
    } else {
      console.log(`get PRs from ${latestUpdated.toISOString()}`);
    }
    const pullRequests =  getPullRequests(repositoryName, latestUpdated);

    for (const pullRequest of pullRequests) {
      // Upsert rows by pullRequest
      const prIds = sheet.getRange("L2:L").getValues().map((vs, _) => vs[0]);

      const last = prIds.filter((id, _) => id !== null && id !== undefined && id !== "").length;

      const index = prIds.findIndex((id, _) => id === `${repositoryName}/${pullRequest.number}`);
      const row = index > 0 ? index + 2: last + 2;

      upsertPullRequestData(pullRequest, sheet, row, repositoryName);

      i++;
    }
  });
}

/** 
 * get vertical values array from range(Value[n][0]) 
 * args 
 *    sheet: GoogleAppsScript.Spreadsheet.SpreadsheetApp
 *    colChar: column index. e.g. 'A'
 *    opt: Row range option.
 * return Value[]
*/
function getVerticalValues(sheet, colChar, opt={head: 0, last:0}) {
  const head = opt.head || 0;
  const last = opt.last || 0;
  const start = head === 0 ? colChar : `${colChar}${head}`;
  const end = last <= 0 ? colChar : `${colChar}${last}`

  return sheet.getRange(`${start}:${end}`).getValues().map((vs, _) => vs[0]);
}

function upsertPullRequestData(pullRequest, sheet, row, repositoryName) {
  let firstCommitDate: Date | null = null;
  if (pullRequest.commits.nodes[0].commit.committedDate) {
    firstCommitDate = new Date(pullRequest.commits.nodes[0].commit.committedDate);
  }
  let mergedAt: Date | null = null;
  if (pullRequest.mergedAt) {
    mergedAt = new Date(pullRequest.mergedAt);
  }

  const updatedAt = new Date(pullRequest.updatedAt);
  sheet.getRange(row, 1).setValue(pullRequest.author.login);
  sheet.getRange(row, 2).setValue(pullRequest.headRefName);
  sheet.getRange(row, 3).setValue(pullRequest.bodyText);
  sheet.getRange(row, 4).setValue(pullRequest.merged);
  sheet.getRange(row, 5).setValue(!!firstCommitDate ? formatDate(firstCommitDate) : "");
  sheet.getRange(row, 6).setValue(!!mergedAt ? formatDate(mergedAt) : "");
  sheet.getRange(row, 7).setValue(
    (!!firstCommitDate && !!mergedAt) ?
      (mergedAt.getTime() - firstCommitDate.getTime()) / 60 / 60 / 1000 :
      "");
  sheet.getRange(row, 8).setValue(repositoryName);
  sheet.getRange(row, 9).setValue(`=REGEXMATCH(B${row}, '分析設定'!$E$3)`);
  sheet.getRange(row, 10).setValue(`=REGEXMATCH(B${row}, '分析設定'!$E$4)`);
  sheet.getRange(row, 11).setValue(formatDate(updatedAt));
  sheet.getRange(row, 12).setValue(`${repositoryName}/${pullRequest.number}`);
}

function getPullRequests(repositoryName: string, updatedFrom: Date | null = null): PullRequest[] {
  const  fetchSizeLimit = 100
  /*
   指定したリポジトリから全てのPRを抽出する.
  */
  let resultPullRequests = [];
  let after = "";
  while (true) {
    const graphql = `query{
      repository(name: "${repositoryName}", owner: "${repositoryOwner}"){
        pullRequests (first: ${fetchSizeLimit} ${after}, orderBy: {field: UPDATED_AT, direction: DESC}) {
          pageInfo {
            startCursor
            hasNextPage
            endCursor
          }
          nodes {
            number 
            author {
              login
            }
            headRefName
            bodyText
            merged
            mergedAt
            commits (first: 1) {
              nodes {
                commit {
                  committedDate
                }
              }
            }
            updatedAt
          }
        }
      }
    }`;
    const option: GoogleAppsScript.URL_Fetch.URLFetchRequestOptions = {
      method: 'post',
      contentType: 'application/json',
      headers: {
        Authorization: 'bearer ' + githubAPIKey
      },
      payload: JSON.stringify({query: graphql}),
    } ;
    const res = UrlFetchApp.fetch(githubEndpoint, option);
    const json = JSON.parse(res.getContentText());
    // filter by updatedAt with updatedFrom
    let updatedRequests = json.data.repository.pullRequests.nodes; 
    
    if (updatedFrom !== null) {
      updatedRequests = updatedRequests.filter((node, i) => (new Date(node.updatedAt).getTime() - updatedFrom.getTime()) > 0);
    }

    resultPullRequests = resultPullRequests.concat(updatedRequests);
    if (updatedRequests.length < fetchSizeLimit || !json.data.repository.pullRequests.pageInfo.hasNextPage) {
      break;
    }
    after = ', after: "' + json.data.repository.pullRequests.pageInfo.endCursor + '"';
  }

  return resultPullRequests.reverse(); // The reverse() occurs javascript runtime error in GAS debug mode.
}

function formatDate(date) {
  /*
   日付フォーマットヘルパ関数
  */
  const year = date.getFullYear();
  const month = (date.getMonth() + 1).toString().padStart(2, '0');
  const day = date.getDate().toString().padStart(2, '0');
  const hour = date.getHours().toString().padStart(2, '0');
  const minute = date.getMinutes().toString().padStart(2, '0');
  const second = date.getSeconds().toString().padStart(2, '0');
  return `${year}-${month}-${day} ${hour}:${minute}:${second}`;
}

function getOrCreateSheet(sheetName) {
  return defaultSheet.getSheetByName(sheetName) || defaultSheet.insertSheet(sheetName);
}

// Export functions for testing
export { getAllRepos, getPullRequests };
