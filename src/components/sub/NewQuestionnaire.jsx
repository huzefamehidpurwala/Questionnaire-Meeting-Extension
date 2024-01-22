import {
  Button,
  LargeTitle,
  Text,
  mergeClasses,
} from "@fluentui/react-components";
import React, { useState, useEffect, useRef, useContext } from "react";
import {
  LiveShareClient,
  LiveState,
  UserMeetingRole,
} from "@microsoft/live-share";
import { LiveShareHost, app } from "@microsoft/teams-js";
import { useData } from "@microsoft/teamsfx-react";
import { TeamsFxContext } from "../Context";
// eslint-disable-next-line
import { allowedRoles, colNames, customPostAnswers, getListItems, isAdmin } from "../../lib/utils";
import SmallPopUp from "../SmallPopUp";
import { useSearchParams } from "react-router-dom";
import NewQuestionnaireStage from "./NewQuestionnaireStage";
import config from "../../lib/config";
import { TeamsUserCredential } from "@microsoft/teamsfx";
// eslint-disable-next-line
import { ArrowLeft48Filled, ArrowRight48Filled } from "@fluentui/react-icons";
import { useGeneralCss, useQuestionnaireCss } from "../../styles";
import "../../Questionnaire.css";

const containerSchema = { initialObjects: { indOfQues: LiveState } };
const exactDateTime = new Date();
const counterInitialValue = 10;

const NewQuestionnaire = () => {
  window.onbeforeunload = function () {
    return "Your saved answers will be lost!";
  };

  const questStyles = useQuestionnaireCss();
  const generalStyles = useGeneralCss();

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

  const teamsUserCredential = useContext(TeamsFxContext).teamsUserCredential;
  const { loading, data, error, reload } = useData(async () => {
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

  // eslint-disable-next-line
  const handleQuesNav = (method) => {
    if (method !== "add" && method !== "sub") return;

    setIndexOfQuestion((prev) => {
      method === "add" ? ++prev : method === "sub" && --prev;
      indOfQues.current.set(data?.value[prev]);
      // await indOfQues.current.set(data?.value[prev]);
      return prev;
    });
  };

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

    let userInfo = undefined;
    try {
      userInfo = await teamsUserCredential.getUserInfo();
    } catch (error) {
      console.error("error in teamsUserCredential", error);
    }

    const currentAttendeeMailId =
      userInfo?.preferredUserName || "error in teamsUserCredential";
    const attendeeName =
      userInfo?.displayName || "error in teamsUserCredential";

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
        setAnsArr([]);
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

            await indOfQues.current.set(null);
            // console.log("console in useEffect ==", userRole,ansArr, questionObj);
            await customPostAnswers(ansArr, questionObj.accessToken);
            // setAnsArr([]); // ! if this is enabled and above setAnsArr is commented then this api runs twice
            setShowQuitPopUp(
              `Quiz Completed${userRole !== UserMeetingRole.organizer
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

          indOfQues.current = container.initialObjects.indOfQues;

          indOfQues.current.on("stateChanged", (quest) => {
            setQuestionObj(quest); // indOfQues.current.state
          });

          // You can optionally declare what roles you want to be able to change state
          await indOfQues.current.initialize(questionObj, allowedRoles);

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

  // * Exit automatically
  isAdmin(userRole) && data?.value.length - 1 < indexOfQuestion && handleExit();

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
            open={!!showQuitPopUp}
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

          <div className={`min-h-screen w-screen bg-[] flex flex-col items-center justify-center gap-4 p-4 select-none`}>
            {questionObj && !("accessToken" in questionObj) ? (
              <>
                {/* // * counter and exit/restatrt btns */}
                <>
                  <div
                    className={`min-w-[75vw] flex ${isAdmin(userRole) ? "justify-between" : "justify-end"
                      } items-center`}
                  >
                    <div className="flex gap-4 items-center">
                      <Text size={500}>Seconds left</Text>
                      <Text
                        font="numeric"
                        size={500}
                        className={mergeClasses(
                          generalStyles.flexCenterBox,
                          questStyles.counter,
                          "rounded-full"
                        )}
                      >
                        {counter}
                      </Text>
                    </div>

                    {isAdmin(userRole) && (
                      <div className="flex gap-4 items-center">
                        <Button
                          className={mergeClasses(
                            generalStyles.border_0,
                            questStyles.exitBtn
                          )}
                          onClick={handleExit}
                        >
                          <Text size={500}>Exit</Text>
                        </Button>
                        <Button
                          className={mergeClasses(
                            generalStyles.border_0,
                            questStyles.startBtn
                          )}
                          onClick={(e) => setStartQuiz(false)}
                        >
                          <Text size={500}>Restart</Text>
                        </Button>
                      </div>
                    )}
                  </div>
                </>

                {/* // * stage */}
                <>
                  <NewQuestionnaireStage
                    fields={questionObj?.fields}
                    userRole={userRole}
                    updateArrOfAnsGiven={updateArrOfAnsGiven}
                    ansArr={ansArr}
                  />
                </>

                {/* // * back-n-for btns */}
                {isAdmin(userRole) && (
                  <>
                    {/* <div className="min-w-[75vw] flex justify-center gap-8 items-center">
                      <Button
                        className={mergeClasses(
                          generalStyles.border_0,
                          questStyles.prevBtn
                        )}
                        disabled={!questionObj || indexOfQuestion < 1}
                        onClick={(e) => handleQuesNav("sub")}
                      >
                        <Text size={500}>Previous</Text>
                      </Button>
                      <Button
                        className={mergeClasses(
                          generalStyles.border_0,
                          questStyles.nextBtn
                        )}
                        disabled={
                          !questionObj ||
                          data?.value.length - 1 < indexOfQuestion + 1
                        }
                        onClick={(e) => handleQuesNav("add")}
                      >
                        <Text size={500}>Next</Text>
                      </Button>
                    </div> */}
                  </>
                )}
              </>
            ) : (
              <>
                {/* // * welcome page */}
                <div className="grid items-center justify-center gap-y-4">
                  <LargeTitle>Welcome to the Quiz..!</LargeTitle>

                  {isAdmin(userRole) && (
                    <Button
                      size="large"
                      className={questStyles.startBtn}
                      onClick={(e) => setStartQuiz(true)}
                    >
                      <LargeTitle>Start</LargeTitle>
                    </Button>
                  )}
                </div>
              </>
            )}
          </div>
        </>
      )}
    </>
  );

  // return (
  //   <>
  //     {!!error || needConsent ? (
  //       <SmallPopUp
  //         open={!!error || needConsent}
  //         activeActions={false}
  //         spinner={false}
  //         modalType="alert"
  //       >
  //         <div className="error">
  //           <Text size={800}>Fetching Data Failed</Text>
  //           <br />
  //           <br />
  //           <Button appearance="primary" disabled={loading} onClick={reload}>
  //             Authorize and call Azure Function
  //           </Button>
  //         </div>
  //       </SmallPopUp>
  //     ) : (
  //       <>
  //         <SmallPopUp
  //           className="loading"
  //           msg={"Fetching Questions..."}
  //           open={pageLoading || loading}
  //           spinner={true}
  //           activeActions={false}
  //           modalType="alert"
  //         />

  //         <SmallPopUp
  //           title={showQuitPopUp}
  //           open={!!showQuitPopUp}
  //           onOpenChange={(e, data) => setShowQuitPopUp(data.open)}
  //           spinner={false}
  //           activeActions={true}
  //           modalType="alert"
  //         />

  //         {/* popup on ListId not found */}
  //         <SmallPopUp
  //           open={noListIdFound}
  //           activeActions={false}
  //           spinner={false}
  //           modalType="alert"
  //         >
  //           <div className="error">
  //             <Text size={800}>No List ID Found</Text>
  //           </div>
  //         </SmallPopUp>

  //         <div className="wrapper">
  //           <div className="flex flex-col justify-center items-center gap-y-4 mb-5">
  //             {questionObj && !("accessToken" in questionObj) && (
  //               <div className="time-laps">
  //                 <div className="label">Second Left</div>
  //                 <div className="count">{counter}</div>
  //               </div>
  //             )}

  //             {userRole === UserMeetingRole.organizer && (
  //               <div className="flex gap-4">
  //                 {/* // * backWard btn */}
  //                 <>
  //                   {/* <Button
  //                     shape="circular"
  //                     appearance="subtle"
  //                     size="large"
  //                     icon={<ArrowLeft48Filled />}
  //                     disabled={!questionObj || indexOfQuestion < 1}
  //                     onClick={(e) => handleQuesNav("sub")}
  //                   /> */}
  //                 </>

  //                 {/* //! need to find solution if questionnaire has only 1 question. */}
  //                 {data?.value.length - 1 < indexOfQuestion + 1 ? (
  //                   <Button size="large" onClick={handleExit}>
  //                     Exit
  //                   </Button>
  //                 ) : (
  //                   <Button
  //                     size="large"
  //                     onClick={(e) => setStartQuiz((t) => !t)}
  //                   >
  //                     {`${startQuiz ? "Restart" : "Start"}`}
  //                   </Button>
  //                 )}

  //                 {/* // * forward btn */}
  //                 <>
  //                   {/* <Button
  //                     shape="circular"
  //                     appearance="subtle"
  //                     size="large"
  //                     icon={<ArrowRight48Filled />}
  //                     disabled={
  //                       !questionObj ||
  //                       data?.value.length - 1 < indexOfQuestion + 1
  //                     }
  //                     onClick={(e) => handleQuesNav("add")}
  //                   /> */}
  //                 </>
  //               </div>
  //             )}
  //           </div>

  //           <div className="ms-Grid">
  //             <div className="ms-Grid-row">
  //               <div className="ms-Grid-col ms-sm12 ms-m12">
  //                 <div className="card">
  //                   {questionObj && !("accessToken" in questionObj) && (
  //                     <NewQuestionnaireStage
  //                       fields={questionObj?.fields}
  //                       userRole={userRole}
  //                       updateArrOfAnsGiven={updateArrOfAnsGiven}
  //                       ansArr={ansArr}
  //                     />
  //                   )}
  //                 </div>
  //               </div>
  //             </div>
  //           </div>
  //         </div>
  //       </>
  //     )}
  //   </>
  // );
};

export default NewQuestionnaire;
