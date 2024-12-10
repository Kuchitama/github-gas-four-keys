interface PullRequest {
  id: int
  author: {
    login: string;
  };
  headRefName: string;
  bodyText: string;
  merged: boolean;
  mergedAt: string | null;
  commits: {
    nodes: Array<{
      commit: {
        committedDate: string;
      };
    }>;
  };
  updatedAt: string;
}

export { PullRequest };
