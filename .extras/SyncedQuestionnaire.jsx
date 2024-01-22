import { Button, Text } from "@fluentui/react-components";
import SmallPopUp from "../SmallPopUp";
import { useEffect, useRef, useState } from "react";
import { colNames } from "../../lib/utils";
import { ArrowLeft48Filled, ArrowRight48Filled } from "@fluentui/react-icons";
import {
  LiveShareClient,
  LiveState,
  UserMeetingRole,
} from "@microsoft/live-share";
import { LiveShareHost, app } from "@microsoft/teams-js";

const SyncedQuestionnaire = () => {
  // const [needConsent, setNeedConsent] = useState(false);
  const [showQuitPopUp, setShowQuitPopUp] = useState(false);
  const [showQuitContent, setShowQuitContent] = useState(false);
  const [userMeetingRole, setUserMeetingRole] = useState("");
  // const [currentQuesNum, setCurrentQuesNum] = useState(0);
  const [questions, setQuestions] = useState(null);
  const [pageLoading, setPageLoading] = useState(true);

  // const lengthOfData = sessionStorage["stageQuestionnaire"]?.length;
  const lengthOfData = useRef(0);

  const handleNav = (id) => {
    console.log("button clicked");
    // if (id) {
    //   id === "right"
    //     ? setCurrentQuesNum((t) => {
    //         if (t < 0) return lengthOfData.current;
    //         else if (t > lengthOfData.current) return 0;
    //         else return t + 1;
    //       })
    //     : id === "left" &&
    //       setCurrentQuesNum((t) => {
    //         if (t < 0) return lengthOfData.current;
    //         else if (t > lengthOfData.current) return 0;
    //         else return t - 1;
    //       });
    //   // intervals(10);
    // }
  };

  const currentLiveQues = useRef(null);
  useEffect(() => {
    // console.log("bhag saale", lengthOfData);
    sessionStorage["userMeetingRole"] &&
      setUserMeetingRole(sessionStorage["userMeetingRole"]);

    if (sessionStorage["stageQuestionnaire"]) {
      const temp = JSON.parse(sessionStorage["stageQuestionnaire"]);
      lengthOfData.current = temp.length;
      // console.log("checking session", temp);
      setQuestions(temp);
      app.initialize().then(async () => {
        const quesContainerSchema = {
          initialObjects: { currentLiveQues: LiveState },
        };

        const host = LiveShareHost.create();
        const client = new LiveShareClient(host);

        const { container } = await client.joinContainer(quesContainerSchema);
        // currentLiveQues.current = quesContainer.initialObjects.currentLiveQues;
        currentLiveQues.current = container.initialObjects.currentLiveQues;
        await currentLiveQues.current.initialize(temp[0]);
        setPageLoading(false);

        // console.log("success", currentLiveQues.current);

        app.notifySuccess();
      });
    }
  }, []);

  console.log("open console in syncedQues", currentLiveQues.current?.state);
  // return <h1>In Stage</h1>;

  return (
    <>
      <>
        {/* <SmallPopUp
          className="loading"
          msg={"Fetching Questions..."}
          open={loading}
          spinner={true}
          activeActions={false}
          modalType="alert"
        /> */}

        <SmallPopUp
          title={"Quiz Completed!"}
          open={showQuitPopUp}
          onOpenChange={(e, data) => setShowQuitPopUp(data.open)}
          spinner={false}
          activeActions={true}
          modalType="alert"
        />

        <div className="container question-container">
          {currentLiveQues.current?.state.id - 1 > 0 &&
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
                {/* {(counter / 100).toFixed(2)} */}
              </Text>
            </div>

            <div className="question">
              <Text size={800} weight="bold">
                <strong>{`Q.${currentLiveQues.current?.state.id})`}</strong>{" "}
                {currentLiveQues.current?.state[colNames[0]]}
              </Text>
            </div>

            <div className="grid-container">
              <div className="container grid-item">
                <Text size={700}>
                  <strong>{`A)`}</strong>{" "}
                  {currentLiveQues.current?.state[colNames[1]]}
                </Text>
              </div>
              <div className="container grid-item">
                <Text size={700}>
                  <strong>{`B)`}</strong>{" "}
                  {currentLiveQues.current?.state[colNames[2]]}
                </Text>
              </div>
              <div className="container grid-item">
                <Text size={700}>
                  <strong>{`C)`}</strong>{" "}
                  {currentLiveQues.current?.state[colNames[3]]}
                </Text>
              </div>
              <div className="container grid-item">
                <Text size={700}>
                  <strong>{`D)`}</strong>{" "}
                  {currentLiveQues.current?.state[colNames[4]]}
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
              {currentLiveQues.current?.state.id < lengthOfData.current ? (
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
      </>
    </>
  );
};

export default SyncedQuestionnaire;
