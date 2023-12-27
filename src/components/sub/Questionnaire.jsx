import { Button, Text, Tooltip } from "@fluentui/react-components";
import React, { useState, useEffect, useRef, useContext } from "react";
import {
  ArrowRight48Filled,
  ArrowLeft48Filled,
  ArrowClockwise48Filled,
} from "@fluentui/react-icons";
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
    // timer: LiveTimer,
  },
};
const exactDateTime = new Date();

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
  // const [toggleState, setToggleState] = useState(false);
  const [userRole, setUserRole] = useState(
    sessionStorage.getItem("userMeetingRole")
  );

  const indOfQues = useRef(undefined);

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

  const handleQuesNav = (method) => {
    if (method === "add") {
      indOfQues.current.set({ ...data?.value[indexOfQuestion + 1] });
      setIndexOfQuestion((t) => ++t);
    } else if (method === "sub") {
      indOfQues.current.set({ ...data?.value[indexOfQuestion - 1] });
      setIndexOfQuestion((t) => --t);
    }
  };

  const handleExit = async () => {
    setStartQuiz(false);

    const authConfig = {
      initiateLoginEndpoint: config.initiateLoginEndpoint,
      clientId: config.clientId,
    };

    const credential = new TeamsUserCredential(authConfig);

    const token = (await credential.getToken(["User.Read"])).token;

    indOfQues.current.set({ accessToken: token });
  };

  const updateArrOfAnsGiven = async (selectedOption) => {
    const fields = questionObj.fields;
    const userInfo = await teamsUserCredential.getUserInfo();
    const currentAttendeeMailId = userInfo.preferredUserName;
    const attendeeName = userInfo.displayName;

    // console.log("checking in func", selectedOption);

    setAnsArr((t) => [
      ...t,
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

  useEffect(() => {
    if (questionObj && ansArr.length && "accessToken" in questionObj) {
      (async () => {
        try {
          await customPostAnswers(ansArr, questionObj.accessToken);
        } catch (error) {
          console.error("error in customPostAnswer", error);
        }
      })();
    }
    /* else if (ansArr.length && "restart" in questionObj) {
      setAnsArr([]);
    } */
    // eslint-disable-next-line
  }, [questionObj]);

  useEffect(() => {
    setPageLoading(true);

    sessionStorage.removeItem("answers");
    setUserRole(sessionStorage.getItem("userMeetingRole"));

    app.initialize().then(async () => {
      try {
        const host = LiveShareHost.create();
        const client = new LiveShareClient(host);

        const { container } = await client.joinContainer(containerSchema);

        indOfQues.current = container.initialObjects.indOfQues;

        // You can optionally declare what roles you want to be able to change state
        const allowedRoles = [
          UserMeetingRole.organizer,
          UserMeetingRole.presenter,
        ];
        await indOfQues.current.initialize(questionObj, allowedRoles);

        indOfQues.current.on("stateChanged", () =>
          setQuestionObj(indOfQues.current.state)
        );

        setPageLoading(false);
      } catch (error) {
        console.error("eror occured", error);
      }

      app.notifySuccess();
    });
    setPageLoading(false);
    // eslint-disable-next-line
  }, []);

  useEffect(() => {
    if (data) {
      if (startQuiz) {
        try {
          setQuestionObj({ ...data.value[0] });
          // setIndexOfQuestion(indOfQues.current.state);
          indOfQues.current.set({ ...data.value[0] });
        } catch (error) {
          console.error(
            "error occurred in setting the fluid container \n",
            error
          );
        }
      } else {
        setQuestionObj(null);
        setIndexOfQuestion(0);
        // setAnsArr([]); // !not possible as this will only effect for organiser
        indOfQues.current?.set(null);
      }
    }
    // eslint-disable-next-line
  }, [data, startQuiz]);

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
            title={"Quiz Completed!"}
            // open={showQuitPopUp}
            // onOpenChange={(e, data) => setShowQuitPopUp(data.open)}
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

          <>
            {/* {userRole !== UserMeetingRole.organizer && (
              <div className="fixed top-4 right-4">
                <Tooltip content="Refresh" positioning="below-end" withArrow>
                  <Button
                    appearance="subtle"
                    size="large"
                    shape="circular"
                    icon={<ArrowClockwise48Filled />}
                    onClick={(e) => window.location.reload()}
                  />
                </Tooltip>
              </div>
            )} */}
          </>

          <div className="grid-center-box gap-4 min-h-screen app-bg-dark-color">
            {userRole === UserMeetingRole.organizer && (
              <div className="flex gap-4">
                <Button
                  shape="circular"
                  appearance="subtle"
                  size="large"
                  icon={<ArrowLeft48Filled />}
                  disabled={!questionObj || indexOfQuestion < 1}
                  // onClick={(e) => setIndexOfQuestion((t) => --t % 10)}
                  onClick={(e) => handleQuesNav("sub")}
                />

                {data?.value.length - 1 < indexOfQuestion + 1 ? (
                  <Button size="large" onClick={handleExit}>
                    Exit
                  </Button>
                ) : (
                  <Button size="large" onClick={(e) => setStartQuiz((t) => !t)}>
                    {`${startQuiz ? "Restart" : "Start"}`}
                  </Button>
                )}

                <Button
                  shape="circular"
                  appearance="subtle"
                  size="large"
                  icon={<ArrowRight48Filled />}
                  disabled={
                    !questionObj || data?.value.length - 1 < indexOfQuestion + 1
                  }
                  // onClick={(e) => setIndexOfQuestion((t) => ++t % 10)}
                  onClick={(e) => handleQuesNav("add")}
                />
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
