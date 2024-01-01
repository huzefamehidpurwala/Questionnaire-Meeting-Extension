const config = {
  initiateLoginEndpoint: process.env.REACT_APP_START_LOGIN_PAGE_URL,
  clientId: process.env.REACT_APP_CLIENT_ID,
  apiEndpoint: process.env.REACT_APP_FUNC_ENDPOINT,
  apiName: process.env.REACT_APP_FUNC_NAME,
  meetingInfoApiName: process.env.REACT_APP_MEETING_INFO_FUNC_NAME,
  postAnswersApiName: process.env.REACT_APP_POST_ANSWERS_FUNC_NAME,
  createQuestionnaireApiName: process.env.REACT_APP_CREATE_QUESTIONNAIRE_FUNC_NAME,
  patchQuestionnaireRootListApiName: process.env.REACT_APP_PATCH_QUESTIONNAIREROOTLIST_FUNC_NAME,
  questionnaireRootListId: process.env.REACT_APP_QUESTIONNAIRE_ROOTLIST
};

export default config;
