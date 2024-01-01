import { Button, Text } from "@fluentui/react-components";
import React, { useState, useEffect, useRef, useContext } from "react";
import {
  LiveShareClient,
  LiveState,
  UserMeetingRole,
} from "@microsoft/live-share";
import { LiveShareHost, app } from "@microsoft/teams-js";
import { useData } from "@microsoft/teamsfx-react";
import { TeamsFxContext } from "../Context";
import { colNames, customPostAnswers, getListItems } from "../../lib/utils";
import SmallPopUp from "../SmallPopUp";
import { useSearchParams } from "react-router-dom";
import QuestionnaireStage from "./QuestionnaireStage";
import config from "../../lib/config";
import { TeamsUserCredential } from "@microsoft/teamsfx";

// const counterKey = "indexOfQuestion";
const containerSchema = {
  initialObjects: {
    indOfQues: LiveState, // SharedMap / LiveState
    // toggleLiveState: LiveState,
  },
};
const exactDateTime = new Date();
const allowedRoles = [UserMeetingRole.organizer, UserMeetingRole.presenter];
const counterInitialValue = 10;

const Questionnaire = () => {
  window.onbeforeunload = function () {
    return "Your saved answers will be lost!";
  };

  const [searchParams] = useSearchParams();
  const questionnaireListId = searchParams.get("listId");

  const [indexOfQuestion, setIndexOfQuestion] = useState(0);
  const [questionObj, setQuestionObj] = useState(null);
  const [pageLoading, setPageLoading] = useState(true);
  const [needConsent, setNeedConsent] = useState(false);
  const [startQuiz, setStartQuiz] = useState(false);
  const [ansArr, setAnsArr] = useState([]);
  const [noListIdFound, setNoListIdFound] = useState(false);
  const [counter, setCounter] = useState(counterInitialValue);
  const [counterInterValue, setCounterInterValue] = useState(undefined);
  const [quesInterValue, setQuesInterValue] = useState(undefined);
  const [showQuitPopUp, setShowQuitPopUp] = useState("");
  const [userRole, setUserRole] = useState(
    sessionStorage.getItem("userMeetingRole")
  );

  const indOfQues = useRef(undefined);
  // const toggleLiveState = useRef(undefined);

  const teamsUserCredential = useContext(TeamsFxContext).teamsUserCredential;
  const { loading, data, error, reload } = useData(async () => {
    // error, reload
    if (!teamsUserCredential) {
      throw new Error("TeamsFx SDK is not initialized.");
    }
    if (needConsent) {
      await teamsUserCredential.login(["Sites.Read.All"]); // "Sites.FullControl.All",
      setNeedConsent(false);
    }
    try {
      const functionRes =
        userRole === UserMeetingRole.organizer
          ? questionnaireListId
            ? await getListItems(teamsUserCredential, questionnaireListId)
            : setNoListIdFound(true)
          : { graphClientMessage: "" };
      return functionRes.graphClientMessage;
    } catch (error) {
      console.error("error in useData ques", error);
      if (error.message.includes("Access Denied")) {
        setNeedConsent(true);
      }
    }
  });

  const handleStartRestart = async (method) => {
    if (method !== "start" && method !== "restart") return;

    const quesVal =
      method === "start" ? data.value[0] : method === "restart" && null;

    method === "restart" && clearIntervals();

    try {
      setIndexOfQuestion(0);
      await indOfQues.current.set(quesVal);
    } catch (error) {
      console.error("error occurred in setting the fluid container \n", error);
    }
  };

  /* const handleQuesNav = async (method) => {
    if (method !== "add" && method !== "sub") return;

    setIndexOfQuestion((prev) => {
      method === "add" ? ++prev : method === "sub" && --prev;
      indOfQues.current.set(data?.value[prev]);
      // await indOfQues.current.set(data?.value[prev]);
      return prev;
    });
  }; */

  const handleExit = async () => {
    setStartQuiz(false);
    setIndexOfQuestion(0);

    const authConfig = {
      initiateLoginEndpoint: config.initiateLoginEndpoint,
      clientId: config.clientId,
    };

    const credential = new TeamsUserCredential(authConfig);

    const token = (await credential.getToken(["User.Read"])).token;

    await indOfQues.current.set({ accessToken: token });
  };

  const updateArrOfAnsGiven = async (selectedOption) => {
    const fields = questionObj.fields;
    const userInfo = await teamsUserCredential.getUserInfo();
    const currentAttendeeMailId = userInfo.preferredUserName;
    const attendeeName = userInfo.displayName;

    // console.log("checking in func", selectedOption);

    setAnsArr((prev) => [
      ...prev,
      {
        Title: currentAttendeeMailId,
        attendeeName,
        questionOfQuestionnaire: fields[colNames[0]],
        questionCorrectAns: fields[colNames[5]],
        ansGivenByAttendee: selectedOption,
        ansIsCorrect: selectedOption === fields[colNames[5]]?.toString(),
        dateTheAttendeeGaveAns: exactDateTime,
        questionnaireListId,
      },
    ]);
  };

  const intervals = () => {
    const counterInterValueInd = setInterval(() => {
      setCounter((prev) => --prev);
    }, 1000);
    setCounterInterValue(counterInterValueInd);

    const quesInterValueInd = setInterval(() => {
      setCounter(counterInitialValue);
      userRole === UserMeetingRole.organizer &&
        setIndexOfQuestion((prev) => {
          indOfQues.current.set(data?.value[++prev]);
          // await indOfQues.current.set(data?.value[++prev]);
          return prev;
        });
    }, (counterInitialValue + 1) * 1000);
    setQuesInterValue(quesInterValueInd);
  };

  const clearIntervals = () => {
    // console.log("i m in clearing func", counterInterValue, quesInterValue);
    clearInterval(counterInterValue);
    clearInterval(quesInterValue);
    setCounter(counterInitialValue);
    setCounterInterValue(undefined);
    setQuesInterValue(undefined);
  };

  useEffect(() => {
    clearIntervals();

    if (questionObj) {
      if ("accessToken" in questionObj) {
        const postAns = async () => {
          try {
            if (ansArr.length > data?.value?.length)
              // eslint-disable-next-line
              throw { name: "customMsg", msg: "you answered a question twice" };

            if (!ansArr.length)
              // eslint-disable-next-line
              throw {
                name: "customMsg",
                msg: "you didn't answer any question",
              };

            await customPostAnswers(ansArr, questionObj.accessToken);
            setAnsArr([]);
            setShowQuitPopUp(
              `Quiz Completed${
                userRole !== UserMeetingRole.organizer
                  ? " and answers submitted"
                  : ""
              }!`
            );
          } catch (error) {
            console.error("error in customPostAnswer", error);
            error.name === "customMsg"
              ? setShowQuitPopUp(`Quiz Completed but ${error.msg}!`)
              : setShowQuitPopUp(
                  "Quiz Completed but answers were not posted. Some error occured!"
                );
          }
        };

        userRole !== UserMeetingRole.organizer
          ? postAns()
          : setShowQuitPopUp("Quiz Completed!");
      }
      !("accessToken" in questionObj) && intervals();
    }
    // eslint-disable-next-line
  }, [questionObj]);

  useEffect(() => {
    setPageLoading(true);

    sessionStorage.removeItem("answers");
    const user = sessionStorage.getItem("userMeetingRole");
    setUserRole(user);

    (data || user !== UserMeetingRole.organizer) &&
      app.initialize().then(async () => {
        try {
          const host = LiveShareHost.create();
          const client = new LiveShareClient(host);

          const { container } = await client.joinContainer(containerSchema);

          // ({ indOfQues: indOfQues.current, toggleLiveState: toggleLiveState.current } = container.initialObjects);
          indOfQues.current = container.initialObjects.indOfQues;
          // console.log("checking this current", container.initialObjects);

          /* toggleLiveState.current.on("stateChanged", (val) => {
            console.log("timre changed!!", val);
            if (val) {
              intervals();
            } else {
              clearIntervals();
            }
          }); */

          indOfQues.current.on("stateChanged", (quest) => {
            setQuestionObj(quest); // indOfQues.current.state
          });

          // You can optionally declare what roles you want to be able to change state
          // await toggleLiveState.current.initialize(true, allowedRoles);
          await indOfQues.current.initialize(questionObj, allowedRoles);

          // console.log("checking initializing ==", indOfQues.current.initializeState);
          // toggleLiveState.current.isInitialized &&
          indOfQues.current.isInitialized && setPageLoading(false);
        } catch (error) {
          console.error("eror occured", error);
          setPageLoading(false);
        }

        app.notifySuccess();
      });
    // eslint-disable-next-line
  }, [data]);

  useEffect(() => {
    if (data) {
      startQuiz ? handleStartRestart("start") : handleStartRestart("restart");
    }
    // eslint-disable-next-line
  }, [startQuiz]);

  userRole === UserMeetingRole.organizer &&
    data?.value.length - 1 < indexOfQuestion &&
    handleExit();

  // console.log("global console in questionnaire ==", ansArr);

  return (
    <>
      {!!error || needConsent ? (
        <SmallPopUp
          open={!!error || needConsent}
          activeActions={false}
          spinner={false}
          modalType="alert"
        >
          <div className="error">
            <Text size={800}>Fetching Data Failed</Text>
            <br />
            <br />
            <Button appearance="primary" disabled={loading} onClick={reload}>
              Authorize and call Azure Function
            </Button>
          </div>
        </SmallPopUp>
      ) : (
        <>
          <SmallPopUp
            className="loading"
            msg={"Fetching Questions..."}
            open={pageLoading || loading}
            spinner={true}
            activeActions={false}
            modalType="alert"
          />

          <SmallPopUp
            title={showQuitPopUp}
            open={showQuitPopUp}
            onOpenChange={(e, data) => setShowQuitPopUp(data.open)}
            spinner={false}
            activeActions={true}
            modalType="alert"
          />

          {/* popup on ListId not found */}
          <SmallPopUp
            open={noListIdFound}
            activeActions={false}
            spinner={false}
            modalType="alert"
          >
            <div className="error">
              <Text size={800}>No List ID Found</Text>
            </div>
          </SmallPopUp>

          <div className="grid-center-box gap-4 relative min-h-screen min-w-full w-auto app-bg-dark-color">
            {questionObj && !("accessToken" in questionObj) && (
              <Text block font="numeric" size={800}>
                {counter} sec left
                {/* {(counter / 100).toFixed(2)} */}
              </Text>
            )}

            {userRole === UserMeetingRole.organizer && (
              <div className="flex gap-4">
                {/* // * backWard btn */}
                <>
                  {/* <Button
                    shape="circular"
                    appearance="subtle"
                    size="large"
                    icon={<ArrowLeft48Filled />}
                    disabled={!questionObj || indexOfQuestion < 1}
                    onClick={(e) => handleQuesNav("sub")}
                  /> */}
                </>

                {/* //! need to find solution if questionnaire has only 1 question. */}
                {data?.value.length - 1 < indexOfQuestion + 1 ? (
                  <Button size="large" onClick={handleExit}>
                    Exit
                  </Button>
                ) : (
                  <Button size="large" onClick={(e) => setStartQuiz((t) => !t)}>
                    {`${startQuiz ? "Restart" : "Start"}`}
                  </Button>
                )}

                {/* // * forward btn */}
                <>
                  {/* <Button
                    shape="circular"
                    appearance="subtle"
                    size="large"
                    icon={<ArrowRight48Filled />}
                    disabled={
                      !questionObj ||
                      data?.value.length - 1 < indexOfQuestion + 1
                    }
                    onClick={(e) => handleQuesNav("add")}
                  /> */}
                </>
              </div>
            )}

            <div className="relative min-w-[90vw] min-h-[80vh] rounded-3xl overflow-auto border-4 border-slate-600">
              <div className="absolute bg-teams-bg-1 w-full h-full opacity-50"></div>
              {questionObj && !("accessToken" in questionObj) && (
                <QuestionnaireStage
                  fields={questionObj?.fields}
                  userRole={userRole}
                  updateArrOfAnsGiven={updateArrOfAnsGiven}
                  ansArr={ansArr}
                />
              )}
            </div>
          </div>
        </>
      )}
    </>
  );
};

export default Questionnaire;
