import { useContext, useEffect, useRef, useState } from "react";
import { TeamsFxContext } from "../Context";
import { useData } from "@microsoft/teamsfx-react";
import { colNames, getListItems } from "../../lib/utils";
import { Button, Text } from "@fluentui/react-components";
import SmallPopUp from "../SmallPopUp";
import { Navigate } from "react-router-dom";
import { ArrowLeft48Filled, ArrowRight48Filled } from "@fluentui/react-icons";
import {
  LiveShareClient,
  LiveState,
  UserMeetingRole,
} from "@microsoft/live-share";
import { LiveShareHost, app } from "@microsoft/teams-js";
import config from "../../lib/config";

const Questionnaire = () => {
  // const [pageLoading, setPageLoading] = useState(true);
  const [needConsent, setNeedConsent] = useState(false);
  // const [counter, setCounter] = useState(10);
  const [currentQues, setCurrentQues] = useState(5);
  const [showQuitPopUp, setShowQuitPopUp] = useState(false);
  const [showQuitContent, setShowQuitContent] = useState(false);
  // const [counterInterValue, setCounterInterValue] = useState("");
  // const [quesInterValue, setQuesInterValue] = useState("");
  const [userMeetingRole, setUserMeetingRole] = useState("");
  const [currentLiveQues, setCurrentLiveQues] = useState(null);

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
      const functionRes = await getListItems(
        teamsUserCredential,
        config.questionnaireRootListId
      );
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

  // *intervals
  /* const intervals = (value) => {
    const timer = value || counter;
    const counterInterValueInd = setInterval(() => {
      setCounter((t) => t - 1);
    }, 1000);
    setCounterInterValue(counterInterValueInd);

    const quesInterValueInd = setInterval(() => {
      setCounter(10);
      setCurrentQues((t) => {
        if (t < 0) return lengthOfData;
        else if (t > lengthOfData) return 0;
        else return t + 1;
      });
    }, (timer + 1) * 1000);
    setQuesInterValue(quesInterValueInd);
    // console.log("huzefa is great", timer);
  }; */

  const containerSchema = {
    initialObjects: { currentLiveQues: LiveState },
  };

  async function joinContainer() {
    // Are we running in teams?
    const host = LiveShareHost.create();

    // Create client
    const client = new LiveShareClient(host);

    // Join container
    return await client.joinContainer(containerSchema);
  }

  const handleUpdate = () => {
    const diceValue = currentLiveQues.state;
    setCurrentQues(diceValue);
    // setPageLoading((t) => !t);
  };

  // const currentLiveQues = useRef(null);
  useEffect(() => {
    // if (data) {
    //   intervals(0);
    // }
    sessionStorage["userMeetingRole"] &&
      setUserMeetingRole(sessionStorage["userMeetingRole"]);

    // if (sessionStorage["stageQuestionnaire"]) {
    // const temp = JSON.parse(sessionStorage["stageQuestionnaire"]);
    // lengthOfData.current = temp.length;
    // console.log("checking session", temp);
    // setQuestions(temp);
    app.initialize().then(async () => {
      try {
        const { container } = await joinContainer();
        const temp = container.initialObjects.currentLiveQues;
        // console.log("jab tujhe", temp);
        setCurrentLiveQues(temp);
        // You can optionally declare what roles you want to be able to change state
        /* const allowedRoles = [
              UserMeetingRole.organizer,
              UserMeetingRole.presenter,
          ]; */
        // Initialize currentLiveQues with initial state of 1 and allowed roles
        // await currentLiveQues.current.initialize(10);
        // handleUpdate();

        // // Use the changed event to trigger the rerender whenever the value changes.
        // currentLiveQues.current.on("stateChanged", handleUpdate);
        // renderStage(diceState, root);
      } catch (error) {
        console.error("error in copy file in initializing", error);
      }
      // console.log("success", currentLiveQues.current);

      app.notifySuccess();
    });
    // }
    // eslint-disable-next-line
  }, []);

  useEffect(() => {
    if (currentLiveQues) {
      (async () => {
        await currentLiveQues.initialize(10);
        handleUpdate();
        console.log("useEffect running");
      })();
    }
  }, [currentLiveQues]);

  let question = null;
  if (data) {
    if (currentQues >= lengthOfData) {
      setCurrentQues(0); // this is imp to make this if condition false
      // clearInterval(counterInterValue);
      // clearInterval(quesInterValue);
      setShowQuitPopUp(true);
      setShowQuitContent(true);
    } else
      question =
        currentLiveQues &&
        data?.graphClientMessage?.value[currentLiveQues?.state];
  }

  const handleNav = (id) => {
    if (id) {
      console.log("in handleNav");
      const ValOfQuesNum = currentLiveQues.state;
      // const counterInterValueInd = clearInterval(counterInterValue);
      // const quesInterValueInd = clearInterval(quesInterValue);
      // setCounterInterValue(counterInterValueInd);
      // setQuesInterValue(quesInterValueInd);
      // setCounter(10);
      id === "right" && currentLiveQues.set(ValOfQuesNum + 1);
      id === "left" && currentLiveQues.set(ValOfQuesNum - 1);

      // handleUpdate();

      currentLiveQues.on("stateChanged", handleUpdate);
      // intervals(10);
    }
  };
  // console.log("kyu bhiya", { error: !!error, needConsent });
  // console.log("kyu bhiya", currentLiveQues, question);
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
            open={loading /*  || pageLoading */}
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

          {question &&
            (!showQuitContent ? (
              <div className="container question-container">
                {currentQues > 0 &&
                (userMeetingRole === UserMeetingRole.organizer ||
                  userMeetingRole === UserMeetingRole.presenter) ? (
                  <div
                    className="container question-navigation-btn"
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
                <div className="container question-area">
                  <div className="container timer-box">
                    <Text size={900} weight="semibold">
                      {/* {(counter / 100).toFixed(2)} */}Timer
                    </Text>
                  </div>

                  <div className="question">
                    <Text size={800} weight="bold">
                      <strong>{`Q.${question?.fields.id})`}</strong>{" "}
                      {question?.fields[colNames[0]]}
                    </Text>
                  </div>

                  <div className="grid-container">
                    <div className="container grid-item">
                      <Text size={700}>
                        <strong>{`A)`}</strong> {question?.fields[colNames[1]]}
                      </Text>
                    </div>
                    <div className="container grid-item">
                      <Text size={700}>
                        <strong>{`B)`}</strong> {question?.fields[colNames[2]]}
                      </Text>
                    </div>
                    <div className="container grid-item">
                      <Text size={700}>
                        <strong>{`C)`}</strong> {question?.fields[colNames[3]]}
                      </Text>
                    </div>
                    <div className="container grid-item">
                      <Text size={700}>
                        <strong>{`D)`}</strong> {question?.fields[colNames[4]]}
                      </Text>
                    </div>
                  </div>
                </div>

                {userMeetingRole === UserMeetingRole.organizer ||
                userMeetingRole === UserMeetingRole.presenter ? (
                  <div
                    className="container question-navigation-btn"
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
              !showQuitPopUp && <Navigate to="/termsofuse" />
            ))}
        </>
      )}
    </>
  );
};

export default Questionnaire;
