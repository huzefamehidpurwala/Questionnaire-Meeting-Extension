import { useData } from "@microsoft/teamsfx-react";
import { TeamsFxContext } from "../Context";
import { useContext, useState, useEffect } from "react";
import {
  compareObjects,
  getListItems,
  handleStringSort,
  toTitleCase,
} from "../../lib/utils";
import SmallPopUp from "../SmallPopUp";
import { Button, Checkbox, Text } from "@fluentui/react-components";
import HChart from "./HChart";
import config from "../../lib/config";

const Analysis = () => {
  const teamsUserCredential = useContext(TeamsFxContext).teamsUserCredential;

  const [needConsent, setNeedConsent] = useState(false);
  const [attendeeNameArr, setAttendeeNameArr] = useState([]);
  const [selectedAttendeeNameArr, setSelectedAttendeeNameArr] = useState([]);
  const [questionnaireObjArr, setQuestionnaireObjArr] = useState([]);
  const [selectedQuestionnaireArr, setSelectedQuestionnaireArr] = useState([]);
  const [initatedquestionnaireIdArr, setInitatedQuestionnaireIdArr] = useState(
    []
  );

  const {
    loading: analyticsOfQuestionnaireLoading,
    data: analyticsOfQuestionnaireData,
    error: analyticsOfQuestionnaireError,
    reload: analyticsOfQuestionnaireReload,
  } = useData(async () => {
    if (!teamsUserCredential) {
      throw new Error("TeamsFx SDK is not initialized.");
    }
    if (needConsent) {
      await teamsUserCredential.login(["Sites.Read.All"]); // "Sites.FullControl.All",
      setNeedConsent(false);
    }
    try {
      // *listId of 'Analytics Of Questionnaire'
      const analyticsOfQuestionnaire = await getListItems(
        teamsUserCredential,
        "d26a4a06-27e1-47cf-9782-155f265f5984"
      );

      return analyticsOfQuestionnaire;
    } catch (error) {
      console.error("error in useData analytics api", error);
      if (error.message.includes("Access Denied")) {
        setNeedConsent(true);
      }
    }
  });

  const {
    loading: questionnaireRootListLoading,
    data: questionnaireRootListData,
    error: questionnaireRootListError,
    reload: questionnaireRootListReload,
  } = useData(async () => {
    if (!teamsUserCredential) {
      throw new Error("TeamsFx SDK is not initialized.");
    }
    if (needConsent) {
      await teamsUserCredential.login(["Sites.Read.All"]); // "Sites.FullControl.All",
      setNeedConsent(false);
    }
    try {
      // *listId of 'questionnaireRootList'
      const questionnaireRootList = await getListItems(
        teamsUserCredential,
        config.questionnaireRootListId
      );

      return questionnaireRootList;
    } catch (error) {
      console.error("error in useData ques api in analytics component", error);
      if (error.message.includes("Access Denied")) {
        setNeedConsent(true);
      }
    }
  });

  const updateQuestionnaireObjArr = () => {
    let tempSet = new Set();
    questionnaireRootListData.graphClientMessage.value.forEach((row) =>
      initatedquestionnaireIdArr.includes(row.fields.idOfLists)
        ? tempSet.add({
            questionnaireName: row.fields.Title,
            questionnaireId: row.fields.idOfLists,
          })
        : ""
    );
    setQuestionnaireObjArr(
      [...tempSet].sort((a, b) =>
        handleStringSort(a.questionnaireName, b.questionnaireName)
      )
    );
  };

  const updateAttendeeNameArr = () => {
    let tempSetAtt = new Set();
    let tempSetId = new Set();
    analyticsOfQuestionnaireData.graphClientMessage.value.forEach((row) => {
      tempSetAtt.add(row.fields.attendeeName);
      tempSetId.add(row.fields.questionnaireListId);
    });

    setAttendeeNameArr([...tempSetAtt].sort(handleStringSort));
    setInitatedQuestionnaireIdArr([...tempSetId]);
  };

  useEffect(() => {
    analyticsOfQuestionnaireData && updateAttendeeNameArr();
    return;
    // eslint-disable-next-line
  }, [analyticsOfQuestionnaireData]);

  useEffect(() => {
    questionnaireRootListData &&
      !!initatedquestionnaireIdArr.length &&
      updateQuestionnaireObjArr();
    return;
    // eslint-disable-next-line
  }, [questionnaireRootListData, initatedquestionnaireIdArr]);

  const checkPresenceOfObj = (objArr, toCheckObj) => {
    for (let i = 0; i < objArr.length; i++) {
      if (compareObjects(objArr[i], toCheckObj)) {
        return true;
      }
    }
    return false;
  };

  const getIndex = (objArr, toCheckObj) => {
    for (let i = 0; i < objArr.length; i++) {
      if (compareObjects(objArr[i], toCheckObj)) {
        return i;
      }
    }
    return -1;
  };

  return (
    <>
      {!!questionnaireRootListError ||
      !!analyticsOfQuestionnaireError ||
      needConsent ? (
        <SmallPopUp
          open={
            !!questionnaireRootListError ||
            !!analyticsOfQuestionnaireError ||
            needConsent
          }
          activeActions={false}
          spinner={false}
          modalType="alert"
        >
          <div className="error">
            <Text size={800}>Fetching Data Failed</Text>
            <br />
            <br />
            <Button
              appearance="primary"
              disabled={questionnaireRootListLoading}
              onClick={questionnaireRootListReload}
            >
              Authorize and call Azure Function
            </Button>{" "}
            <Button
              appearance="primary"
              disabled={analyticsOfQuestionnaireLoading}
              onClick={analyticsOfQuestionnaireReload}
            >
              Authorize and call Azure Function
            </Button>
          </div>
        </SmallPopUp>
      ) : (
        <>
          {/* no data exists */}
          <SmallPopUp
            open={
              !(
                analyticsOfQuestionnaireLoading || questionnaireRootListLoading
              ) &&
              !analyticsOfQuestionnaireData?.graphClientMessage.value.length
            }
            activeActions={false}
            spinner={false}
            modalType="alert"
          >
            <div className="error">
              <Text size={800}>No Data Exists!</Text>
            </div>
          </SmallPopUp>

          <SmallPopUp
            className="loading"
            msg={"Preparing things..."}
            open={
              analyticsOfQuestionnaireLoading || questionnaireRootListLoading
            }
            spinner={true}
            activeActions={false}
            modalType="alert"
          />

          {!!analyticsOfQuestionnaireData?.graphClientMessage.value.length &&
            !!questionnaireRootListData &&
            !!analyticsOfQuestionnaireData && (
              <>
                <div className="analytics-flex-container">
                  <div className="sub-container1">
                    <div>
                      <Button
                        onClick={(e) => {
                          setSelectedAttendeeNameArr([]);
                          setSelectedQuestionnaireArr([]);
                        }}
                      >
                        Clear Selection
                      </Button>
                    </div>

                    <div className="checkbox-container">
                      <Text size={400} weight="bold">
                        Select Attendee Name:
                      </Text>
                      {attendeeNameArr.length > 1 && (
                        <Checkbox
                          label={<Text size={200}>Select All Attedees</Text>}
                          checked={
                            !!selectedAttendeeNameArr.length
                              ? attendeeNameArr.length >
                                selectedAttendeeNameArr.length
                                ? "mixed"
                                : true
                              : false
                          }
                          onChange={(e, data) => {
                            data.checked
                              ? setSelectedAttendeeNameArr([...attendeeNameArr])
                              : setSelectedAttendeeNameArr([]);
                          }}
                        />
                      )}
                      {attendeeNameArr.map((attendeeName, key) => (
                        <Checkbox
                          label={attendeeName}
                          key={key}
                          checked={selectedAttendeeNameArr.includes(
                            attendeeName
                          )}
                          onChange={(e, data) => {
                            data.checked
                              ? setSelectedAttendeeNameArr((t) => [
                                  ...t,
                                  attendeeName,
                                ])
                              : setSelectedAttendeeNameArr((t) => {
                                  t.splice(t.indexOf(attendeeName), 1);
                                  return [...t];
                                });
                          }}
                        />
                      ))}
                    </div>

                    <div className="checkbox-container">
                      <Text size={400} weight="bold">
                        Select Questionnaire:
                      </Text>
                      {questionnaireObjArr.length > 1 && (
                        <Checkbox
                          label={
                            <Text size={200}>Select All Questionnaires</Text>
                          }
                          checked={
                            !!selectedQuestionnaireArr.length
                              ? questionnaireObjArr.length >
                                selectedQuestionnaireArr.length
                                ? "mixed"
                                : true
                              : false
                          }
                          onChange={(e, data) => {
                            data.checked
                              ? setSelectedQuestionnaireArr([
                                  ...questionnaireObjArr,
                                ])
                              : setSelectedQuestionnaireArr([]);
                          }}
                        />
                      )}
                      {questionnaireObjArr.map((obj, key) => (
                        <Checkbox
                          label={toTitleCase(obj.questionnaireName)}
                          key={key}
                          checked={checkPresenceOfObj(
                            selectedQuestionnaireArr,
                            obj
                          )}
                          onChange={(e, data) => {
                            const newObj = {
                              questionnaireName: obj.questionnaireName, // e.target.value
                              questionnaireId: obj.questionnaireId, // e.target.id
                            };
                            data.checked
                              ? setSelectedQuestionnaireArr((t) => [
                                  ...t,
                                  newObj,
                                ])
                              : setSelectedQuestionnaireArr((t) => {
                                  t.splice(getIndex(t, newObj), 1);
                                  return [...t];
                                });
                          }}
                        />
                      ))}
                    </div>

                    {/* // *date below */}
                  </div>
                  <div className="sub-container2">
                    {!!selectedAttendeeNameArr.length &&
                      !!selectedQuestionnaireArr.length &&
                      selectedQuestionnaireArr
                        .sort((a, b) =>
                          handleStringSort(
                            a.questionnaireName,
                            b.questionnaireName
                          )
                        )
                        .map((obj) => {
                          return (
                            <HChart
                              key={obj.questionnaireId}
                              questionnaireId={obj.questionnaireId}
                              questionnaireName={toTitleCase(
                                obj.questionnaireName
                              )}
                              selectedAttendeeNameArr={selectedAttendeeNameArr}
                              chartType={"column"}
                              analyticsOfQuestionnaireData={
                                analyticsOfQuestionnaireData
                              }
                            />
                          );
                        })}
                  </div>
                </div>
              </>
            )}
        </>
      )}
    </>
  );
};

export default Analysis;
