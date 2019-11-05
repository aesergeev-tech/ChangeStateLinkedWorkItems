import * as tl from "azure-pipelines-task-lib/task";
import * as request from "request-promise-native";

const collectionUrl = process.env["SYSTEM_TEAMFOUNDATIONCOLLECTIONURI"];
const teamProject = process.env["SYSTEM_TEAMPROJECT"];
const accessToken = tl.getEndpointAuthorization("SystemVssConnection", true)
  .parameters.AccessToken;

async function run() {
  try {
    const pipelineType = tl.getInput("pipelineType");
    const workItemsData =
      pipelineType === "Build"
        ? await getWorkItemsFromBuild()
        : await getWorkItemsFromRelease();
    workItemsData.forEach(async (workItem: any) => {
      await changeStateOfWorkItem(workItem);
    });
  } catch (err) {
    tl.logIssue(tl.IssueType.Error, err.message);
    tl.setResult(tl.TaskResult.Failed, err.message);
  }
}

async function getWorkItemsFromBuild() {
  const buildId = process.env["BUILD_BUILDID"];
  const uri = `${collectionUrl}/${teamProject}/_apis/build/builds/${buildId}/workitems`;
  const options = createGetRequestOptions(uri);
  const result = await request.get(options);
  return result.value;
}

async function getWorkItemsFromRelease() {
  const releaseId = process.env["RELEASE_RELEASEID"];
  const uri = `${collectionUrl}/${teamProject}/_apis/release/releases/${releaseId}/workitems`;
  const options = createGetRequestOptions(uri);
  const result = await request.get(options);
  return result.value;
}

async function changeStateOfWorkItem(workItem: any) {
  const desiredState = tl.getInput("desiredState");
  const uri = workItem.url + "?api-version=2.0";
  const patchOptions = getPatchRequestOptions(uri, desiredState);
  await request.patch(patchOptions);
}

function createGetRequestOptions(uri: string): any {
  let options = {
    uri: uri,
    headers: {
      authorization: `Bearer ${accessToken}`,
      "content-type": "application/json"
    },
    json: true
  };
  return options;
}

function getPatchRequestOptions(uri: string, desiredState: string): any {
  const options = {
    uri: uri,
    headers: {
      authorization: `Bearer ${accessToken}`,
      "content-type": "application/json-patch+json"
    },
    body: [
      {
        op: "add",
        path: "/fields/System.State",
        value: desiredState
      }
    ],
    json: true
  };
  return options;
}

run();
