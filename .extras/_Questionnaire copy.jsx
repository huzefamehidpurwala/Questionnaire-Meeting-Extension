import { useContext, useEffect, useState } from "react";
import { TeamsFxContext } from "../Context";
import { useData } from "@microsoft/teamsfx-react";
import { colNames, getListItems, postAnswers } from "../../lib/utils";
import { Button, Text } from "@fluentui/react-components";
import SmallPopUp from "../SmallPopUp";
import { Navigate, useSearchParams } from "react-router-dom";
import { ArrowLeft48Filled, ArrowRight48Filled } from "@fluentui/react-icons";
import { UserMeetingRole } from "@microsoft/live-share";

const optionChars = ["A", "B", "C", "D"];
const exactDateTime = new Date();

const Questionnaire = () => {
  window.onbeforeunload = function () {
    return "Your saved answers will be lost!";
  };

  const [searchParams] = useSearchParams();
  const questionnaireListId = searchParams.get("listId");
  /* useEffect(() => {
    const [listId, rowIds] = searchParams;
    console.log("check searchParams", rowIds, listId);
  }, []); */

  // const [pageLoading, setPageLoading] = useState(false);
  const [needConsent, setNeedConsent] = useState(false);
  const [counter, setCounter] = useState(10);
  const [currentQues, setCurrentQues] = useState(0);
  const [showQuitPopUp, setShowQuitPopUp] = useState(false);
  const [showQuitContent, setShowQuitContent] = useState(false);
  const [counterInterValue, setCounterInterValue] = useState("");
  const [quesInterValue, setQuesInterValue] = useState("");
  const [userMeetingRole, setUserMeetingRole] = useState("");
  const [noListIdFound, setNoListIdFound] = useState(false);
  // const [isAnsCorrect, setIsAnsCorrect] = useState("");
  const [isAnsGiven, setIsAnsGiven] = useState(false);
  const [arrOfAnsGiven, setArrOfAnsGiven] = useState([]);
  const [selectedOptionColValue, setSelectedOptionColValue] = useState("");

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
      // console.log("check list id url query", searchParams.get("listId"));
      const functionRes = questionnaireListId
        ? await getListItems(teamsUserCredential, questionnaireListId)
        : setNoListIdFound(true);
      return functionRes;
    } catch (error) {
      console.error("error in useData ques", error);
      if (error.message.includes("Access Denied")) {
        setNeedConsent(true);
      }
    }
  });

  const lengthOfData = data?.graphClientMessage?.value?.length;
  // console.log(lengthOfData);

  const intervals = (value) => {
    const timer = value || counter;
    const counterInterValueInd = setInterval(() => {
      setCounter((t) => t - 1);
    }, 1000);
    setCounterInterValue(counterInterValueInd);

    const quesInterValueInd = setInterval(() => {
      setCounter(10);
      setIsAnsGiven(false);
      setSelectedOptionColValue("");
      // setIsAnsCorrect("");
      setCurrentQues((t) => {
        if (t < 0) return lengthOfData;
        else if (t > lengthOfData) return 0;
        else return t + 1;
      });
    }, (timer + 1) * 1000);
    setQuesInterValue(quesInterValueInd);
    // console.log("huzefa is great", timer);
  };

  /* const containerSchema = { initialObjects: { currentLiveQues: LiveState } };
  const currentLiveQues = useRef(null);
  
  useEffect(() => {
    app.initialize().then(async () => {
      // (async () => {
      const host = LiveShareHost.create();
      const client = new LiveShareClient(host);
      const { container } = await client.joinContainer(containerSchema);
      currentLiveQues.current = container.initialObjects.currentLiveQues;
      await currentLiveQues.current.initialize({ currentQues, counter });
      // liveState.on("update", );
      // console.log("second console in ques", liveState.state);
      // })();

      app.notifySuccess();
    });
  }, []); */

  /* useEffect(() => {
    // currentLiveQues. = container.initialObjects.currentLiveQues;
    currentLiveQues.current?.set({ currentQues, counter });
    console.log("first console in ques", currentLiveQues.current?.state);
  }, [currentQues, counter]); */

  useEffect(() => {
    if (data) {
      intervals(0);
    }
    sessionStorage.removeItem("answers");
    const temp = sessionStorage.getItem("userMeetingRole");
    setUserMeetingRole(temp);
    // (async () => {
    //   const userInfo = await teamsUserCredential.getUserInfo();
    //   console.log("sdfsa", userInfo);
    // })();
    // eslint-disable-next-line
  }, [data]);

  /* useEffect(() => {
    app.initialize().then(async () => {
      meeting.getAppContentStageSharingState((err, result) => {
        console.log("wowwewa", result);
        console.warn("wowwewa", err);
         if (result.isAppSharing) {
        // Indicates if app is sharing content on the meeting stage.
      } 
      });
    });
  }, []); */

  let question = null;
  if (data) {
    if (currentQues >= lengthOfData) {
      setCurrentQues(0); // this is imp to make this if condition false
      clearInterval(counterInterValue);
      clearInterval(quesInterValue);
      setShowQuitPopUp(true);
      setShowQuitContent(true);
    } else question = data?.graphClientMessage?.value[currentQues];
  }

  const handleNav = (id) => {
    if (id) {
      const counterInterValueInd = clearInterval(counterInterValue);
      const quesInterValueInd = clearInterval(quesInterValue);
      setCounterInterValue(counterInterValueInd);
      setQuesInterValue(quesInterValueInd);
      setIsAnsGiven(false);
      setSelectedOptionColValue("");
      // setIsAnsCorrect("");
      setCounter(10);
      id === "right"
        ? setCurrentQues((t) => {
            if (t < 0) return lengthOfData;
            else if (t > lengthOfData) return 0;
            else return t + 1;
          })
        : id === "left" &&
          setCurrentQues((t) => {
            if (t < 0) return lengthOfData;
            else if (t > lengthOfData) return 0;
            else return t - 1;
          });
      intervals(10);
    }
  };

  const updateArrOfAnsGiven = async (selectedOptionColValueArgv) => {
    // const quesNum = quesNumArgv ? quesNumArgv - 1 : currentQues;
    const currentQuestionFields =
      data.graphClientMessage.value[currentQues].fields;
    // console.log("ho bhiya", selectedOptionColValueArgv, "===", currentQuestionFields[colNames[5]]?.toString().replace(/\s/gm,''));
    // const checkCorrectAns = selectedOptionColValueArgv === currentQuestionFields[colNames[5]];
    // setIsAnsCorrect(checkCorrectAns ? "correct-ans" : "incorrect-ans");

    // let currentAttendeeMailId = "";
    // let attendeeName = "";
    // if (teamsUserCredential) {
    const userInfo = await teamsUserCredential.getUserInfo();
    const currentAttendeeMailId = userInfo.preferredUserName;
    const attendeeName = userInfo.displayName;
    // }

    setArrOfAnsGiven((t) => [
      ...t,
      {
        Title: currentAttendeeMailId,
        attendeeName,
        questionOfQuestionnaire: currentQuestionFields[colNames[0]],
        questionCorrectAns: currentQuestionFields[colNames[5]],
        ansGivenByAttendee: selectedOptionColValueArgv,
        ansIsCorrect:
          selectedOptionColValueArgv ===
          currentQuestionFields[
            colNames[5]
          ]?.toString() /* .replace(/\s/gm, "") */,
        dateTheAttendeeGaveAns: exactDateTime,
        questionnaireListId,
      },
    ]);
  };

  useEffect(() => {
    if (showQuitContent && arrOfAnsGiven?.length !== 0) {
      // console.log("yo yo0", arrOfAnsGiven);

      // postAnswers(teamsUserCredential, arrOfAnsGiven);
      // setArrOfAnsGiven([]);
      (async () => {
        try {
          const reply = await postAnswers(teamsUserCredential, arrOfAnsGiven);
          console.log("pst success", reply);
        } catch (error) {
          console.error("error in posting answers", error);
        }
        setArrOfAnsGiven([]);
      })();
      return;
    }
    // eslint-disable-next-line
  }, [showQuitContent]);

  // console.log("kyu bhiya", { error: !!error, needConsent });
  // console.log("global console huzef===", question?.fields);
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
            open={loading}
            spinner={true}
            activeActions={false}
            modalType="alert"
          />

          <SmallPopUp
            title={"Quiz Completed!"}
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

          {/* popup on ans given */}
          <>
            {/* <SmallPopUp
              open={isAnsGiven}
              // onOpenChange={(e, data) => setIsAnsGiven(data.open)}
              activeActions={false}
              spinner={false}
              modalType="alert"
            >
              <div className={`error ${isAnsCorrect}`}>
                <Text size={800}>{toTitleCase(isAnsCorrect)}</Text>
              </div>
            </SmallPopUp> */}
          </>

          {/* {console.log("check ansert status", selectedOptionColValue === question?.fields[colNames[5]])} */}
          {question &&
            (!showQuitContent ? (
              <div className="flex-container question-container">
                {currentQues > 0 &&
                (userMeetingRole === UserMeetingRole.organizer ||
                  userMeetingRole === UserMeetingRole.presenter) ? (
                  <div
                    className="flex-container question-navigation-btn"
                    id="left"
                    onClick={(e) => handleNav(e.target.id)}
                  >
                    <ArrowLeft48Filled id="left" />
                  </div>
                ) : (
                  <div
                    style={{
                      minWidth: "6rem",
                      minHeight: "6rem",
                      margin: "1rem",
                    }}
                  ></div>
                )}
                <div className="flex-container question-area">
                  {isAnsGiven && (
                    <div
                      className={
                        selectedOptionColValue ===
                        question?.fields[colNames[5]]?.toString()
                          ? /* .replace(/\s/gm, "") */
                            "correct-ans"
                          : "incorrect-ans"
                      }
                    >
                      <Text size={800}>
                        {selectedOptionColValue ===
                        question?.fields[colNames[5]]?.toString()
                          ? /* .replace(/\s/gm, "") */
                            "Correct Answer!"
                          : "Wrong Answer!"}
                      </Text>
                    </div>
                  )}

                  <div className="flex-container timer-box">
                    <Text size={900} weight="semibold">
                      {(counter / 100).toFixed(2)}
                    </Text>
                  </div>

                  <div className="question">
                    <Text size={800} weight="bold">
                      <strong>{`Q.${question?.fields.id})`}</strong>{" "}
                      {question?.fields[colNames[0]]}
                    </Text>
                  </div>

                  <div className="grid-container">
                    {optionChars.map((char, index) => {
                      const currOptionIsCorrectAns = question?.fields[
                        colNames[5]
                      ]
                        ?.toString()
                        /* .replace(/\s/gm, "") */
                        .includes(index + 1);
                      const currOptionIsSelected =
                        selectedOptionColValue.includes(index + 1);
                      return (
                        <div
                          className={
                            !isAnsGiven
                              ? "flex-container grid-item"
                              : `flex-container grid-item disable-option ${
                                  currOptionIsCorrectAns
                                    ? "correct-ans"
                                    : "incorrect-ans"
                                } ${currOptionIsSelected && "selected-option"}`
                          }
                          key={index}
                          onClick={(e) => {
                            // userMeetingRole !== UserMeetingRole.organizer && userMeetingRole !== UserMeetingRole.presenter &&
                            if (!isAnsGiven) {
                              setIsAnsGiven(true);
                              setSelectedOptionColValue(colNames[index + 1]);
                              updateArrOfAnsGiven(colNames[index + 1]); // /* question?.fields.id, */ colNames[index + 1]
                            }
                          }}
                        >
                          <Text size={700}>
                            <strong>{char}.</strong>{" "}
                            {question?.fields[colNames[index + 1]]}
                          </Text>
                        </div>
                      );
                    })}
                  </div>
                </div>

                {userMeetingRole === UserMeetingRole.organizer ||
                userMeetingRole === UserMeetingRole.presenter ? (
                  <div
                    className="flex-container question-navigation-btn"
                    id="right"
                    onClick={(e) => handleNav(e.target.id)}
                  >
                    {currentQues < lengthOfData - 1 ? (
                      <ArrowRight48Filled id="right" />
                    ) : (
                      <Text size={600}>Exit</Text>
                    )}
                  </div>
                ) : (
                  <div
                    style={{
                      minWidth: "6rem",
                      minHeight: "6rem",
                      margin: "1rem",
                    }}
                  ></div>
                )}
              </div>
            ) : (
              !showQuitPopUp && <Navigate to="/analytics" />
            ))}
        </>
      )}
    </>
  );
};

export default Questionnaire;
