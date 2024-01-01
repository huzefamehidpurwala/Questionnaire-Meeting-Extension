import { BearerTokenAuthProvider, createApiClient } from "@microsoft/teamsfx";
import config from "./config";
import { executeDeepLink } from "@microsoft/teams-js";
import axios from "axios";

const getQuestionsFunc = config.apiName || "getQuestions";
const getMeetingInfoFunc = config.meetingInfoApiName || "getMeetingInfo";
const postAnswersFunc = config.postAnswersApiName || "postAnswers";
const patchQuestionnaireRootListFunc =
  config.patchQuestionnaireRootListApiName || "patchQuestionnaireRootList";
const createQuestionnaireFunc =
  config.createQuestionnaireApiName || "createQuestionnaire";

// *custom hook
/* export const useGetTeamsUserCredential = () => {
  const teamsUserCredential = useContext(TeamsFxContext).teamsUserCredential;
  return teamsUserCredential;
}; */

export async function patchQuestionnaireRootList(
  teamsUserCredential,
  quesRowId,
  exactDateTime
) {
  if (!teamsUserCredential) {
    throw new Error("TeamsFx SDK is not initialized.");
  }
  try {
    const apiBaseUrl = config.apiEndpoint + "/api/";
    // createApiClient(...) creates an Axios instance which uses BearerTokenAuthProvider to inject token to request header
    const apiClient = createApiClient(
      apiBaseUrl,
      new BearerTokenAuthProvider(
        async () => (await teamsUserCredential.getToken("")).token
      )
    );
    const response = await apiClient.post(patchQuestionnaireRootListFunc, {
      quesRowId,
      exactDateTime,
    });
    // config.questions = response.graphClientMessage;
    return response.data;
  } catch (err) {
    let funcErrorMsg = "";
    if (err?.response?.status === 404) {
      funcErrorMsg = `There may be a problem with the deployment of Azure Function App, please deploy Azure Function (Run command palette "Teams: Deploy") first before running this App`;
    } else if (err.message === "Network Error") {
      funcErrorMsg =
        "Cannot call Azure Function due to network error, please check your network connection status and ";
      if (err.config.url.indexOf("localhost") >= 0) {
        funcErrorMsg += `make sure to start Azure Function locally (Run "npm run start" command inside api folder from terminal) first before running this App`;
      } else {
        funcErrorMsg += `make sure to provision and deploy Azure Function (Run command palette "Teams: Provision" and "Teams: Deploy") first before running this App`;
      }
    } else {
      funcErrorMsg = err.message;
      if (err.response?.data?.error) {
        funcErrorMsg += ": " + err.response.data.error;
      }
    }
    throw new Error(funcErrorMsg);
  }
}

export async function postAnswers(teamsUserCredential, answersArr) {
  if (!teamsUserCredential) {
    throw new Error("TeamsFx SDK is not initialized.");
  }
  try {
    const apiBaseUrl = config.apiEndpoint + "/api/";
    // createApiClient(...) creates an Axios instance which uses BearerTokenAuthProvider to inject token to request header
    const apiClient = createApiClient(
      apiBaseUrl,
      new BearerTokenAuthProvider(
        async () => (await teamsUserCredential.getToken("")).token
      )
    );
    const response = await apiClient.post(postAnswersFunc, { answersArr });
    // config.questions = response.graphClientMessage;
    return response.data;
  } catch (err) {
    let funcErrorMsg = "";
    if (err?.response?.status === 404) {
      funcErrorMsg = `There may be a problem with the deployment of Azure Function App, please deploy Azure Function (Run command palette "Teams: Deploy") first before running this App`;
    } else if (err.message === "Network Error") {
      funcErrorMsg =
        "Cannot call Azure Function due to network error, please check your network connection status and ";
      if (err.config.url.indexOf("localhost") >= 0) {
        funcErrorMsg += `make sure to start Azure Function locally (Run "npm run start" command inside api folder from terminal) first before running this App`;
      } else {
        funcErrorMsg += `make sure to provision and deploy Azure Function (Run command palette "Teams: Provision" and "Teams: Deploy") first before running this App`;
      }
    } else {
      funcErrorMsg = err.message;
      if (err.response?.data?.error) {
        funcErrorMsg += ": " + err.response.data.error;
      }
    }
    throw new Error(funcErrorMsg);
  }
}

export async function getMeetingInfo(teamsUserCredential, chatId) {
  if (!teamsUserCredential) {
    throw new Error("TeamsFx SDK is not initialized.");
  }
  try {
    const apiBaseUrl = config.apiEndpoint + "/api/";
    // createApiClient(...) creates an Axios instance which uses BearerTokenAuthProvider to inject token to request header
    const apiClient = createApiClient(
      apiBaseUrl,
      new BearerTokenAuthProvider(
        async () => (await teamsUserCredential.getToken("")).token
      )
    );
    const response = await apiClient.post(getMeetingInfoFunc, { chatId });
    // config.questions = response.graphClientMessage;
    return response.data;
  } catch (err) {
    let funcErrorMsg = "";
    if (err?.response?.status === 404) {
      funcErrorMsg = `There may be a problem with the deployment of Azure Function App, please deploy Azure Function (Run command palette "Teams: Deploy") first before running this App`;
    } else if (err.message === "Network Error") {
      funcErrorMsg =
        "Cannot call Azure Function due to network error, please check your network connection status and ";
      if (err.config.url.indexOf("localhost") >= 0) {
        funcErrorMsg += `make sure to start Azure Function locally (Run "npm run start" command inside api folder from terminal) first before running this App`;
      } else {
        funcErrorMsg += `make sure to provision and deploy Azure Function (Run command palette "Teams: Provision" and "Teams: Deploy") first before running this App`;
      }
    } else {
      funcErrorMsg = err.message;
      if (err.response?.data?.error) {
        funcErrorMsg += ": " + err.response.data.error;
      }
    }
    throw new Error(funcErrorMsg);
  }
}

export async function getListItems(teamsUserCredential, listId, sort = "") {
  if (!teamsUserCredential) {
    throw new Error("TeamsFx SDK is not initialized.");
  }
  try {
    const apiBaseUrl = config.apiEndpoint + "/api/";
    // createApiClient(...) creates an Axios instance which uses BearerTokenAuthProvider to inject token to request header
    const apiClient = createApiClient(
      apiBaseUrl,
      new BearerTokenAuthProvider(
        async () => (await teamsUserCredential.getToken("")).token
      )
    );
    const response = listId
      ? await apiClient.post(getQuestionsFunc, { listId, sort })
      : await apiClient.get(getQuestionsFunc);
    // config.questions = response.graphClientMessage;
    return response.data;
  } catch (err) {
    let funcErrorMsg = "";
    if (err?.response?.status === 404) {
      funcErrorMsg = `There may be a problem with the deployment of Azure Function App, please deploy Azure Function (Run command palette "Teams: Deploy") first before running this App`;
    } else if (err.message === "Network Error") {
      funcErrorMsg =
        "Cannot call Azure Function due to network error, please check your network connection status and ";
      if (err.config.url.indexOf("localhost") >= 0) {
        funcErrorMsg += `make sure to start Azure Function locally (Run "npm run start" command inside api folder from terminal) first before running this App`;
      } else {
        funcErrorMsg += `make sure to provision and deploy Azure Function (Run command palette "Teams: Provision" and "Teams: Deploy") first before running this App`;
      }
    } else {
      funcErrorMsg = err.message;
      if (err.response?.data?.error) {
        funcErrorMsg += ": " + err.response.data.error;
      }
    }
    throw new Error(funcErrorMsg);
  }
}

export async function createQuestionnaire(
  teamsUserCredential,
  newList,
  listFields
) {
  if (!teamsUserCredential) {
    throw new Error("TeamsFx SDK is not initialized.");
  }
  try {
    const apiBaseUrl = config.apiEndpoint + "/api/";
    // createApiClient(...) creates an Axios instance which uses BearerTokenAuthProvider to inject token to request header
    const apiClient = createApiClient(
      apiBaseUrl,
      new BearerTokenAuthProvider(
        async () => (await teamsUserCredential.getToken("")).token
      )
    );
    const response = await apiClient.post(createQuestionnaireFunc, {
      newList,
      listFields,
    });
    // config.questions = response.graphClientMessage;
    return response.data;
  } catch (err) {
    let funcErrorMsg = "";
    if (err?.response?.status === 404) {
      funcErrorMsg = `There may be a problem with the deployment of Azure Function App, please deploy Azure Function (Run command palette "Teams: Deploy") first before running this App`;
    } else if (err.message === "Network Error") {
      funcErrorMsg =
        "Cannot call Azure Function due to network error, please check your network connection status and ";
      if (err.config.url.indexOf("localhost") >= 0) {
        funcErrorMsg += `make sure to start Azure Function locally (Run "npm run start" command inside api folder from terminal) first before running this App`;
      } else {
        funcErrorMsg += `make sure to provision and deploy Azure Function (Run command palette "Teams: Provision" and "Teams: Deploy") first before running this App`;
      }
    } else {
      funcErrorMsg = err.message;
      if (err.response?.data?.error) {
        funcErrorMsg += ": " + err.response.data.error;
      }
    }
    throw new Error(funcErrorMsg);
  }
}

export async function customPostAnswers(ansArr, accessToken) {
  for (const ansObj of ansArr) {
    await axios
      .post(
        "https://graph.microsoft.com/v1.0/sites/29151005-7f34-4489-a82f-ff4a19cb537a/lists/d26a4a06-27e1-47cf-9782-155f265f5984/items",
        { fields: ansObj },
        { headers: { Authorization: `Bearer ${accessToken}` } }
      )
      .then((response) => {
        console.log("in utils customPostAnswers", response);
        // if (response.status === 201) {
        //   // setShowEdit(!showEdit);
        //   // BtnClick();
        //   alert("successful create");
        // } else {
        //   alert("create failed");
        // }
      });
  }
}

export const colNames = [
  "Title",
  "option1", // "field_1",
  "option2", // "field_2",
  "option3", // "field_3",
  "option4", // "field_4",
  "correctOption", // "field_5",
];

export const propsOfStateObj = [
  "idOfQues", // 0
  "question", // 1
  "option1", // 2
  "option2", // 3
  "option3", // 4
  "option4", // 5
  "correctOption", // 6
];

export const encryptObj = (objToEncrypt) => {
  const stringifiedVar = JSON.stringify(objToEncrypt);

  let encryptedString = [];
  for (let i = 0; i < stringifiedVar.length; i++) {
    let code = stringifiedVar.charCodeAt(i);
    encryptedString.push(code + 3);
  }

  return JSON.stringify(encryptedString);
  /* try {
    sessionStorage.setItem(sessionKey, JSON.stringify(encryptedString));
    return true;
  } catch (error) {
    console.error("falied to save in session Storage", error);
    return false;
  } */
};

export const decryptStr = (strToDecode) => {
  // const strToDecode = sessionStorage.getItem(sessionKey);
  const encryptedArr = JSON.parse(strToDecode);

  let decryptedString = "";
  for (let i = 0; i < encryptedArr.length; i++) {
    let code = String.fromCharCode(encryptedArr[i] - 3);
    decryptedString += code.toString();
  }

  return JSON.parse(decryptedString);
};

export const updateListFields = (objToUpdate) => {
  const updated = objToUpdate;
  for (const obj of updated) {
    obj.Title = obj[propsOfStateObj[1]];
    delete obj[propsOfStateObj[0]];
    delete obj[propsOfStateObj[1]];
  }

  return updated;
};

export function toTitleCase(str) {
  // Convert underscores, hyphens, and camelCase to spaces
  str = str
    .replace(/_/g, " ")
    .replace(/-/g, " ")
    .replace(/([a-z])([A-Z])/g, "$1 $2");

  // Add a space before a capital letter if it's preceded by a lowercase letter
  str = str.replace(/([a-z])([A-Z])/g, "$1 $2");

  return str.replace(
    /\w\S*/g,
    (txt) => txt.charAt(0).toUpperCase() + txt.substr(1).toLowerCase()
  );
}

export function compareObjects(object1, object2) {
  const keys1 = Object.keys(object1);
  const keys2 = Object.keys(object2);

  if (keys1.length !== keys2.length) {
    return false;
  }

  for (let key of keys1) {
    if (object1[key] !== object2[key]) {
      return false;
    }
  }

  return true;
}

export function convertDate(str) {
  var date = new Date(str),
    mnth = ("0" + (date.getMonth() + 1)).slice(-2),
    day = ("0" + date.getDate()).slice(-2);
  return [day, mnth, date.getFullYear()].join("-");
}

export function convertDateTime(inputDateString) {
  // Parse the input date string
  const inputDate = new Date(inputDateString);

  // Format the date
  const day = inputDate.getDate().toString().padStart(2, "0");
  const month = (inputDate.getMonth() + 1).toString().padStart(2, "0");
  const year = inputDate.getFullYear();

  // Format the time
  const hours = inputDate.getHours() % 12 || 12; // Convert to 12-hour format
  const minutes = inputDate.getMinutes().toString().padStart(2, "0");
  // const seconds = inputDate.getSeconds().toString(); /* .padStart(2, "0") */
  const period = inputDate.getHours() < 12 ? "am" : "pm";

  // Assemble the formatted string
  const formattedString = `${day}-${month}-${year} ${hours}:${minutes}${period}`;

  return formattedString;
}

export const redirectUsingDeeplink = (pathName) => {
  // var encodedContext = `{"subEntityId":  ${data}}`;

  const webUrl = `https://teams.microsoft.com/l/entity/${config.teamsAppId}${pathName}`;
  executeDeepLink(webUrl);
};

export const handleStringSort = (a, b, desc = false) => {
  const x = a.toLowerCase();
  const y = b.toLowerCase();
  if (x < y) return desc ? 1 : -1;
  if (x > y) return desc ? -1 : 1;
  return 0;
};
