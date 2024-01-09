const axios = require("axios");
const XLSX = require("xlsx");
const jsonData = require("../input.json");
const moment = require("moment");

const encodedToken = Buffer.from(`:${jsonData.personalAccessToken}`).toString(
  "base64"
);

const headers = {
  Authorization: `Basic ${encodedToken}`,
  "content-type": "application/json",
};

async function generateData() {
  const workItemForExcel = [];
  const iterationData = await axios.post(
    jsonData.uri,
    { query: jsonData.query },
    { headers }
  );
  const workItems = iterationData.data.workItems;
  for (let i = 0; i < workItems.length; i++) {
    const userStoryItem = workItems[i];
    const userStoryResponse = await axios.get(userStoryItem.url, { headers });
    const userStoryResponseData = userStoryResponse.data;

    const userStoryRevesionId =
      userStoryResponseData._links.workItemRevisions.href;
    const userStoryRevesionResponseHistory = await axios.get(
      userStoryRevesionId,
      { headers }
    );

    const formedData = {
      sNo: i + 1,
      id: userStoryResponseData.id,
      title: userStoryResponseData.fields["System.Title"],
      storyPoints:
        userStoryResponseData.fields["Microsoft.VSTS.Scheduling.StoryPoints"] ??
        0,
      status: userStoryResponseData.fields["System.State"],
      assignedTo:
        userStoryResponseData.fields["System.AssignedTo"]["displayName"],
      history: userStoryRevesionResponseHistory.data.value.map(
        (e) => `title:  ${e.fields["System.Title"]},
      status: ${e.fields["System.State"]},
      assignedTo: ${
        e.fields["System.AssignedTo"] == undefined
          ? ""
          : e.fields["System.AssignedTo"]["displayName"]
      },
      changeDate : ${moment(new Date(e.fields["System.ChangedDate"])).format(
        "MMMM Do YYYY, h:mm:ss a"
      )},
      changedBy: ${e.fields["System.ChangedBy"]["displayName"]},
      sprintPath: ${e.fields["System.IterationPath"]},
      comments: ${e.fields["System.History"]?.toString() ?? ""}`
      ).join("\n\n"),
    };

    //get linked Work items to the story
    const linkedWorkItemsResponse = await axios.post(
      jsonData.uri,
      {
        query: `SELECT [System.Id] FROM WorkItemLinks WHERE ([Source].[System.Id] = ${userStoryResponseData.id}) AND ([System.Links.LinkType] = 'System.LinkTypes.Hierarchy-Forward')`,
      },
      { headers }
    );
    //pick only proper work items where source is null
    const workLinkedList = linkedWorkItemsResponse.data["workItemRelations"];
    const workLinkedListFiltered = workLinkedList
      .filter((e) => e.source != null)
      .map((e) => e.target.url);
    for (let j = 0; j < workLinkedListFiltered.length; j++) {
      const linkedItemUrl = workLinkedListFiltered[j];
      const subWorkItemResponse = await axios.get(linkedItemUrl, { headers });
      formedData["tasks"] = `${formedData["tasks"] ?? ""}${
        subWorkItemResponse.data.fields["System.Title"]
      } ,`;
      //get history of work items for single staks or story
      const revisionsResponse = await axios.get(`${linkedItemUrl}/revisions`, {
        headers,
      });
      const some = revisionsResponse.data.value.filter(
        (e) => e.fields["Microsoft.VSTS.Scheduling.RemainingWork"] != null
      );
      const firstLoggedData = some.sort(
        (a, b) =>
          new Date(a["Microsoft.VSTS.Common.StateChangeDate"]) <
          new Date(b[["Microsoft.VSTS.Common.StateChangeDate"]])
      );
      const firstLoggedDataD = firstLoggedData.map(
        (e) => e.fields["Microsoft.VSTS.Scheduling.RemainingWork"]
      )[0];
      formedData["totalWorkHours"] = `${
        parseInt(formedData["totalWorkHours"] ?? "0") +
        parseInt(firstLoggedDataD)
      }`;
    }
    workItemForExcel.push(formedData);
  }

  const worksheet = XLSX.utils.json_to_sheet(workItemForExcel);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, "Azure Report");
  XLSX.writeFile(workbook, `azure_report.xlsx`);
}

generateData();
