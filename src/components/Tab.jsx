import { useContext, useEffect, useState } from "react";
import { TeamsFxContext } from "./Context";
import { getMeetingInfo } from "../lib/utils";
import { Text, mergeClasses } from "@fluentui/react-components";
import { FrameContexts, app } from "@microsoft/teams-js";
import { UserMeetingRole } from "@microsoft/live-share";
import SmallPopUp from "./SmallPopUp";
import MeetingStarted from "./sub/MeetingStarted";
import CreateQuestionnaire from "./sub/CreateQuestionnaire";
import AdminSidePanel from "./sub/AdminSidePanel";
import { Navigate } from "react-router-dom";

export default function Tab() {
  const {
    themeString,
    // teamsPageType: { current: teamsPageType },
  } = useContext(TeamsFxContext);
  // console.log("keseho", teamsPageType);

  // const teamsPageType = useRef("");
  // const [teamsPageType, setTeamsPageType] = useState(app.getFrameContext());

  const [currentUserRole, setCurrentUserRole] = useState("");
  const [meetingStartDateTime, setMeetingStartDateTime] = useState("");
  const [meetingEndDateTime, setMeetingEndDateTime] = useState("");
  const [persnolTab, setPersnolTab] = useState(false);

  const teamsUserCredential = useContext(TeamsFxContext).teamsUserCredential;

  const setCurrentUserRoleFunc = (currentUserId, participantsObj) => {
    if (participantsObj) {
      if (participantsObj.organizer?.identity?.user.id === currentUserId) {
        // console.log("meeting info tab.jsx", participantsObj.organizer?.identity?.user.id);
        setCurrentUserRole(UserMeetingRole.organizer);
        // config.userMeetingRole = UserMeetingRole.organizer;
        sessionStorage.setItem("userMeetingRole", UserMeetingRole.organizer);
        // return;
      } else {
        participantsObj?.attendees.forEach((people) => {
          if (people?.identity.user?.id === currentUserId) {
            Object.keys(UserMeetingRole).forEach((role) => {
              // console.log("meeting info tab.jsx", role);
              if (people.role === role) {
                setCurrentUserRole(UserMeetingRole[role]);
                // config.userMeetingRole = UserMeetingRole[role];
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
    app.initialize().then(async () => {
      // Get our frameContext from context of our app in Teams
      app.getContext().then(async (context) => {
        // console.log("kyyubfaskjf", context);
        if (context.chat?.id && context.meeting?.id) {
          setPersnolTab(false);
          const currentUserId = context.user.id;
          // console.log("adsfljasl", currentUserId);
          const recieved = await getMeetingInfo(
            teamsUserCredential,
            context.chat?.id
          );
          // console.log("first tab", recieved);

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
    }); // eslint-disable-next-line
  }, []);
  // console.warn(currentUserRole, "kesho");
  // console.log(!!currentUserRole, "kesho");

  // const [pageLoading, setPageLoading] = useState(false);
  // const [needConsent, setNeedConsent] = useState(false);

  // const teamsUserCredential = useContext(TeamsFxContext).teamsUserCredential;
  // const { loading, data, error, reload } = useData(async () => {
  //   if (!teamsUserCredential) {
  //     throw new Error("TeamsFx SDK is not initialized.");
  //   }
  //   if (needConsent) {
  //     await teamsUserCredential.login(["User.Read"]);
  //     setNeedConsent(false);
  //   }
  //   try {
  //     const functionRes = await getListItems(teamsUserCredential);
  //     return functionRes;
  //   } catch (error) {
  //     if (error.message.includes("The application may not be authorized.")) {
  //       setNeedConsent(true);
  //     }
  //   }
  // });

  // scopes: ["OnlineMeetings.Read", "Chat.Read"]
  // const chat = await Axios.get<Chat>(`https://graph.microsoft.com/v1.0/chats/${chatId}`, authHeader);
  // chat scopes = [Chat.Read, Chat.ReadBasic, Chat.ReadWrite]
  // meeting scopes = [OnlineMeeting Artifact.Read.All, OnlineMeetings.Read, OnlineMeetings.ReadWrite]
  // const onlineMeetings = await Axios.get(`https://graph.microsoft.com/v1.0/me/onlineMeetings?$filter=JoinWebUrl eq '${chat.data.onlineMeetingInfo?.joinWebUrl}'`, authHeader);
  // above api will also provide the roles of the users: organizer, presenter, attendee
  /* const startTime = new Date (onlineMeeting.startDateTime as string) ;
  const mainContentElement = (startTime.getTime () < Date.now ()) */

  let component = <></>;
  const currentTime = new Date();
  switch (/* "meetingStage" */ app.getFrameContext()) {
    case FrameContexts.content:
      component = persnolTab ? (
        // <Questionnaire />
        <CreateQuestionnaire persnolTab={persnolTab} />
      ) : // <>
      //   <Text size={700}>Persnol Tab Exp</Text>
      //   <Analysis />
      // </>
      meetingEndDateTime && meetingStartDateTime ? (
        currentTime > meetingEndDateTime ? (
          <Navigate to="/analytics" />
        ) : currentTime < meetingStartDateTime ? (
          currentUserRole === UserMeetingRole.organizer ||
          currentUserRole === UserMeetingRole.presenter ? (
            <CreateQuestionnaire persnolTab={persnolTab} />
          ) : (
            <h1>You can not create Questionnaire</h1>
          )
        ) : meetingStartDateTime < currentTime < meetingEndDateTime ? (
          currentUserRole === UserMeetingRole.organizer ||
          currentUserRole === UserMeetingRole.presenter ? (
            <CreateQuestionnaire />
          ) : (
            <MeetingStarted />
          )
        ) : null
      ) : null;
      break;

    case FrameContexts.sidePanel:
      component =
        currentUserRole === UserMeetingRole.organizer ||
        currentUserRole === UserMeetingRole.presenter ? (
          <AdminSidePanel />
        ) : (
          <Text>You have to answer the questions when the timer starts</Text>
        );
      break;

    case FrameContexts.meetingStage:
      // component = <Questionnaire />;
      component = (
        <h1>
          No Questionnaire was selected or an unauthorized person tried to
          shared the application
        </h1>
      );
      break;

    default:
      component = <h1>You are not in MS Teams Env</h1>;
      break;
  }

  /* console.log(
    "<MeetingStarted />",
    meetingStartDateTime < currentTime < meetingEndDateTime
  );
  console.log("<CreateQuestionnaire />", currentTime < meetingStartDateTime);
  console.log("<Analysis />", currentTime > meetingEndDateTime);
  console.log("<Persnol Tab Exp />", persnolTab); */

  return (
    <div
      className={mergeClasses(
        themeString === "default"
          ? "light"
          : themeString === "dark"
          ? "dark"
          : "contrast",
        "custom-container"
      )}
    >
      <SmallPopUp
        className="loading"
        msg={"Getting things ready..."}
        open={!currentUserRole} //  || !!component
        spinner={true}
        activeActions={false}
        modalType="alert"
      />

      {currentUserRole && component}
    </div>
  );
}
