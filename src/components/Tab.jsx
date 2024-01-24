import { useContext, useEffect, useRef, useState } from "react";
import { TeamsFxContext } from "./Context";
import { getMeetingInfo, isAdmin } from "../lib/utils";
import { Display, Text, mergeClasses } from "@fluentui/react-components";
import { FrameContexts, app } from "@microsoft/teams-js";
import { UserMeetingRole } from "@microsoft/live-share";
import SmallPopUp from "./SmallPopUp";
import MeetingStarted from "./sub/MeetingStarted";
import CreateQuestionnaireNew from "./sub/CreateQuestionnaireNew";
import AdminSidePanel from "./sub/AdminSidePanel";
import { Navigate } from "react-router-dom";
import axios from "axios";
import config from "../lib/config";
import { TeamsUserCredential } from "@microsoft/teamsfx";

const currentTime = new Date();

export default function Tab() {
  const { themeString, teamsUserCredential } = useContext(TeamsFxContext);

  const [currentUserRole, setCurrentUserRole] = useState("");
  const [meetingStartDateTime, setMeetingStartDateTime] = useState("");
  const [meetingEndDateTime, setMeetingEndDateTime] = useState("");
  const [persnolTab, setPersnolTab] = useState(false);

  const setCurrentUserRoleFunc = (currentUserId, participantsObj) => {
    if (participantsObj) {
      if (participantsObj.organizer?.identity?.user.id === currentUserId) {
        setCurrentUserRole(UserMeetingRole.organizer);
        sessionStorage.setItem("userMeetingRole", UserMeetingRole.organizer);
      } else {
        participantsObj?.attendees.forEach((people) => {
          if (people?.identity.user?.id === currentUserId) {
            Object.keys(UserMeetingRole).forEach((role) => {
              if (people.role === role) {
                setCurrentUserRole(UserMeetingRole[role]);
                sessionStorage.setItem(
                  "userMeetingRole",
                  UserMeetingRole[role]
                );
              }
            });
          }
        });
      }
    }
  };

  useEffect(() => {
    // Initialize teams app
    app
      .initialize()
      .then(async () => {
        // Get our frameContext from context of our app in Teams
        app.getContext().then(async (context) => {
          if (context.chat?.id && context.meeting?.id) {
            setPersnolTab(false);
            const currentUserId = context.user.id;
            const recieved = await getMeetingInfo(
              teamsUserCredential,
              context.chat?.id
            );

            const participantsObj =
              recieved?.graphClientMessage?.value[0]?.participants;
            const tempMeetingStartDateTime = new Date(
              recieved?.graphClientMessage?.value[0]?.startDateTime
            );
            const tempMeetingEndDateTime = new Date(
              recieved?.graphClientMessage?.value[0]?.endDateTime
            );

            setMeetingStartDateTime(tempMeetingStartDateTime);
            setMeetingEndDateTime(tempMeetingEndDateTime);

            setCurrentUserRoleFunc(currentUserId, participantsObj);
          } else {
            setPersnolTab(true);
            setCurrentUserRole("Not in Meeting Exp");
          }
        });

        app.notifySuccess();
      })
      .catch((err) => setCurrentUserRole("Not in Teams Env")); // eslint-disable-next-line
  }, []);

  // *create site in sharepoint
  // const [isQuestionnaireSitePresent, setIsQuestionnaireSitePresent] =
  //   useState(true);
  // // console.log("hello app.jsx", isQuestionnaireSitePresent);
  // const teamsPageType = useRef("");
  // useEffect(() => {
  //   // Initialize teams app
  //   app.initialize().then(async () => {
  //     const authConfig = {
  //       initiateLoginEndpoint: config.initiateLoginEndpoint,
  //       clientId: config.clientId,
  //     };
  //     const credential = new TeamsUserCredential(authConfig);
  //     const token = await credential.getToken([
  //       "https://ygr11.sharepoint.com/Sites.ReadWrite.All",
  //     ]);
  //     // console.log("TOKEN------------->", token.token);

  //     *api to create site in sharepoint
  //     const res = axios.post(
  //       "https://ygr11.sharepoint.com/_api/SPSiteManager/create", // * /create /delete  https://learn.microsoft.com/en-us/sharepoint/dev/apis/site-creation-rest
  //       {
  //         request: {
  //           Title: "New Hidden Teams Site by API", // *name of the site
  //           Description: "Site created b API", // *Desc of the site
  //           WebTemplate: "STS#3", // *teams site -> "WebTemplate":"STS#3" (or) communication site -> "WebTemplate":"SITEPAGEPUBLISHING#0"
  //           Url: "https://ygr11.sharepoint.com/sites/hteams1", // *url for that site
  //         },
  //       },
  //       { headers: { Authorization: `Bearer ${token.token}` } }
  //     );

  //     app.notifySuccess();
  //   });
  // }, []);

  return (
    <div
      className={mergeClasses(
        themeString === "default"
          ? "light"
          : themeString === "dark"
            ? "dark"
            : "contrast",
        "flex-container"
      )}
    >
      <SmallPopUp
        className="loading"
        msg={"Getting things ready..."}
        open={!currentUserRole}
        spinner={true}
        activeActions={false}
        modalType="alert"
      />

      {currentUserRole &&
        (() => {
          switch (app.getFrameContext()) {
            case FrameContexts.content:
              return persnolTab ? (
                <CreateQuestionnaireNew persnolTab={persnolTab} />
              ) : meetingEndDateTime && meetingStartDateTime ? (
                currentTime > meetingEndDateTime ? (
                  <Navigate to="/analytics" />
                ) : currentTime < meetingStartDateTime ? (
                  isAdmin() ? (
                    <CreateQuestionnaireNew persnolTab={persnolTab} />
                  ) : (
                    <h1>You can not create Questionnaire</h1>
                  )
                ) : meetingStartDateTime < currentTime < meetingEndDateTime ? (
                  isAdmin() ? (
                    <CreateQuestionnaireNew />
                  ) : (
                    <MeetingStarted />
                  )
                ) : null
              ) : null;

            case FrameContexts.sidePanel:
              return isAdmin() ? (
                <AdminSidePanel />
              ) : (
                <Text>
                  You have to answer the questions when the timer starts
                </Text>
              );

            case FrameContexts.meetingStage:
              return (
                <Display>
                  No Questionnaire was selected or an unauthorized person tried
                  to shared the application
                </Display>
              );

            default:
              return <Display>You are not in MS Teams Env</Display>;
          }
        })()}
    </div>
  );
}
