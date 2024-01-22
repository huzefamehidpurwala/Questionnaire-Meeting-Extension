import { useData } from "@microsoft/teamsfx-react";
import { TeamsFxContext } from "../Context";
import { useContext, useState, useEffect } from "react";
import { compareObjects, getListItems, toTitleCase } from "../../lib/utils";
import SmallPopUp from "../SmallPopUp";
import { Button, Checkbox, Text } from "@fluentui/react-components";
import HChart from "./HChart";
import config from "../../lib/config";

const Analysis = () => {
  const teamsUserCredential = useContext(TeamsFxContext).teamsUserCredential;

  const [needConsent, setNeedConsent] = useState(false);
  const [attendeeNameArr, setAttendeeNameArr] = useState([]);
  // const [dateTimeArr, setDateTimeArr] = useState([]);
  // const [selectedDateTime, setSelectedDateTime] = useState("");
  const [selectedAttendeeNameArr, setSelectedAttendeeNameArr] = useState([]);
  const [questionnaireObjArr, setQuestionnaireObjArr] = useState([]);
  const [selectedQuestionnaireArr, setSelectedQuestionnaireArr] = useState([]);
  // const [selectedQuestionnaireIdArr, setSelectedQuestionnaireIdArr] = useState([]);
  // const [selectedQuestionnaireNameArr, setSelectedQuestionnaireNameArr] = useState([]);
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
      console.error("error in useData analysis api", error);
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
      console.error("error in useData ques api in analysis component", error);
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
          }) // tempSet.add(row.fields.Title)
        : ""
    );
    // console.log("tempSet", tempSet);
    setQuestionnaireObjArr([...tempSet]);
  };

  const updateAttendeeNameArr = () => {
    let tempSetAtt = new Set();
    let tempSetId = new Set();
    analyticsOfQuestionnaireData.graphClientMessage.value.forEach((row) => {
      tempSetAtt.add(row.fields.attendeeName);
      tempSetId.add(row.fields.questionnaireListId);
    });
    // console.log("bhag bhiya", tempSetAtt, tempSetId);
    setAttendeeNameArr([...tempSetAtt]);
    setInitatedQuestionnaireIdArr([...tempSetId]);
  };

  // const updateDateTimeArr = () => {
  //   let tempSetDate = new Set();
  //   analyticsOfQuestionnaireData.graphClientMessage.value.forEach((row) => {
  //     const id = row.fields.questionnaireListId;
  //     const date = row.fields.dateTheAttendeeGaveAns;
  //     JSON.stringify(selectedQuestionnaireArr).includes(id) &&
  //       tempSetDate.add(date);
  //   });
  //   // console.log("bhag bhiya", tempSetDate, tempSetId);
  //   setDateTimeArr([...tempSetDate]);
  // };

  // useEffect(() => {
  //   !!selectedQuestionnaireArr.length && updateDateTimeArr();
  //   return;
  //   // eslint-disable-next-line
  // }, [selectedQuestionnaireArr]);

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

  // console.log("in analytics global==", dateTimeArr);
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
          {!analyticsOfQuestionnaireData?.graphClientMessage.value.length ? (
            <SmallPopUp
              open={
                !analyticsOfQuestionnaireData?.graphClientMessage.value.length
              }
              // onOpenChange={(e, data) => setSuccessCreate(data.open)}
              activeActions={false}
              spinner={false}
              modalType="alert"
            >
              <div className="error">
                <Text size={800}>No Data Exists!</Text>
              </div>
            </SmallPopUp>
          ) : (
            !!questionnaireRootListData &&
            !!analyticsOfQuestionnaireData && (
              <>
                <div className="analytics-flex-container">
                  <div className="sub-container1">
                    <div>
                      <Button
                        // appearance="transparent"
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
                          // size="large"
                          checked={
                            !!selectedAttendeeNameArr.length
                              ? attendeeNameArr.length >
                                selectedAttendeeNameArr.length
                                ? "mixed"
                                : true
                              : false
                          }
                          onChange={(e, data) => {
                            // console.log("onChange == ", data.checked/* , "---", e */);
                            data.checked
                              ? setSelectedAttendeeNameArr([...attendeeNameArr])
                              : setSelectedAttendeeNameArr([]);
                          }}
                        />
                      )}
                      {attendeeNameArr.map((attendeeName, key) => (
                        <Checkbox
                          label={attendeeName} // <Text size={400}>{attendeeName}</Text>
                          // value={attendeeName}
                          key={key}
                          checked={selectedAttendeeNameArr.includes(
                            attendeeName
                          )}
                          onChange={(e, data) => {
                            // console.log("onChange == ", data.checked/* , "---", e */);
                            data.checked
                              ? setSelectedAttendeeNameArr((t) => [
                                  ...t,
                                  attendeeName,
                                ])
                              : setSelectedAttendeeNameArr((t) => {
                                  t.splice(t.indexOf(attendeeName), 1);
                                  // console.log("in state", t);
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
                          // size="large"
                          checked={
                            !!selectedQuestionnaireArr.length
                              ? questionnaireObjArr.length >
                                selectedQuestionnaireArr.length
                                ? "mixed"
                                : true
                              : false
                          }
                          onChange={(e, data) => {
                            // console.log("onChange == ", data.checked/* , "---", e */);
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
                          label={toTitleCase(obj.questionnaireName)} // <Text size={400}>{toTitleCase(obj.questionnaireName)}</Text>
                          // value={obj.questionnaireName}
                          // id={obj.questionnaireId}
                          key={key}
                          checked={checkPresenceOfObj(
                            selectedQuestionnaireArr,
                            obj
                          )}
                          onChange={(e, data) => {
                            // console.log("onChange == ", data.checked/* , "---", e */);
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
                                  // console.log("in state", t);
                                  return [...t];
                                });
                          }}
                        />
                      ))}
                    </div>

                    {/* // *date */}
                    <>
                      {/* {!!selectedQuestionnaireArr.length && (
                      <div className="checkbox-container">
                        <Text size={400} weight="bold">
                          Select Date:
                        </Text>
                        {dateTimeArr.length > 1 && (
                          <Checkbox
                            label={<Text size={200}>Select All Dates</Text>}
                            // size="large"
                            // checked={
                            //   !!selectedAttendeeNameArr.length
                            //     ? attendeeNameArr.length >
                            //       selectedAttendeeNameArr.length
                            //       ? "mixed"
                            //       : true
                            //     : false
                            // }
                            // onChange={(e, data) => {
                            //   // console.log("onChange == ", data.checked);
                            //   data.checked
                            //     ? setSelectedAttendeeNameArr([...attendeeNameArr])
                            //     : setSelectedAttendeeNameArr([]);
                            // }}
                          />
                        )}
                        {dateTimeArr.map((date, key) => (
                          <Checkbox
                            label={convertDateTime(date)} // <Text size={400}>{toTitleCase(obj.questionnaireName)}</Text>
                            key={key}
                            checked={selectedDateTime.includes(date)}
                            onChange={(e, data) => {
                              data.checked
                                ? setSelectedDateTime(date)
                                : setSelectedDateTime("");
                            }}
                          />
                        ))}
                      </div>
                    )} */}
                    </>
                  </div>
                  <div className="sub-container2">
                    {!!selectedAttendeeNameArr.length &&
                      !!selectedQuestionnaireArr.length &&
                      // selectedDateTime &&
                      selectedQuestionnaireArr.map((obj, key) => (
                        <HChart
                          key={key}
                          questionnaireId={obj.questionnaireId}
                          questionnaireName={toTitleCase(obj.questionnaireName)}
                          selectedAttendeeNameArr={selectedAttendeeNameArr}
                          chartType={"column"}
                          // selectedDateTime={selectedDateTime}
                          // questionnaireRootListData={questionnaireRootListData}
                          analyticsOfQuestionnaireData={
                            analyticsOfQuestionnaireData
                          }
                        />
                      ))}
                  </div>
                </div>
              </>
            )
          )}

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
        </>
      )}
    </>
  );
};

export default Analysis;
