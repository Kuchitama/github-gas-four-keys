import { PullRequest } from "./types/main";
import { Sheet, FourKeysSheet, PullRequestsSheet, SettingsSheet } from "./src/Sheets";

const githubEndpoint = "https://api.github.com/graphql";

const repositoryNames = JSON.parse(PropertiesService.getScriptProperties().getProperty("GITHUB_REPO_NAMES") || "");
const repositoryOwner = PropertiesService.getScriptProperties().getProperty("GITHUB_REPO_OWNER");
const githubAPIKey = PropertiesService.getScriptProperties().getProperty("GITHUB_API_TOKEN");

function initialize() {

  new PullRequestsSheet().initialize();
  new SettingsSheet().initialize(PullRequestsSheet.sheetName, repositoryNames);
  new FourKeysSheet().initialize(PullRequestsSheet.sheetName);
  
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === "getAllRepos").forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger("getAllRepos")
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(0)
    .create();
}

function getAllRepos() {
  // Get latest updatedAt
  const pullRequestsSheet = new PullRequestsSheet();
  const settingsSheet = new SettingsSheet();

  repositoryNames.forEach((repositoryName) => {
    const latestUpdated = settingsSheet.getLatestUpdatedAt(repositoryName);
    if ( latestUpdated === null ) {
      console.log(`get all PRs`);
    } else {
      console.log(`get PRs from ${latestUpdated.toISOString()}`);
    }

    const pullRequests: PullRequest[] =  getPullRequests(repositoryName, latestUpdated);

    pullRequests.forEach(pullRequest => pullRequestsSheet.upsertPullRequest(repositoryName, pullRequest));
  });
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
      updatedRequests = updatedRequests.filter((node, _) => (new Date(node.updatedAt).getTime() - updatedFrom.getTime()) > 0);
    }

    resultPullRequests = resultPullRequests.concat(updatedRequests);
    if (updatedRequests.length < fetchSizeLimit || !json.data.repository.pullRequests.pageInfo.hasNextPage) {
      break;
    }
    after = ', after: "' + json.data.repository.pullRequests.pageInfo.endCursor + '"';
  }

  return resultPullRequests.reverse(); // The reverse() occurs javascript runtime error in GAS debug mode.
}

// Export functions for testing
export { getAllRepos, getPullRequests };
