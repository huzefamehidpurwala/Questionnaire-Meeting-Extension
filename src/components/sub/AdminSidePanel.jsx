import { meeting } from "@microsoft/teams-js";
import { UserMeetingRole } from "@microsoft/live-share";
import { Button, Text } from "@fluentui/react-components";
import { Open16Regular } from "@fluentui/react-icons";
import { useContext, useEffect, useState } from "react";
import { TeamsFxContext } from "../Context";
import { getListItems, patchQuestionnaireRootList } from "../../lib/utils";
import SmallPopUp from "../SmallPopUp";
import { useData } from "@microsoft/teamsfx-react";
import config from "../../lib/config";

const exactDateTime = new Date();

const AdminSidePanel = () => {
  const teamsUserCredential = useContext(TeamsFxContext).teamsUserCredential;
  const [userMeetingRole, setUserMeetingRole] = useState("");
  const [needConsent, setNeedConsent] = useState(false);
  const [btnClicked, setBtnClicked] = useState("");

  const fetchData = async () => {
    if (
      !userMeetingRole === UserMeetingRole.organizer ||
      !userMeetingRole === UserMeetingRole.presenter
    )
      return;
    if (!teamsUserCredential) {
      throw new Error("TeamsFx SDK is not initialized.");
    }
    if (needConsent) {
      await teamsUserCredential.login(["Sites.Read.All"]); // "Sites.FullControl.All",
      setNeedConsent(false);
    }
    try {
      // if (!listId) {
      const functionRes = await getListItems(
        teamsUserCredential,
        config.questionnaireRootListId,
        "fields/isInitiated,fields/Created desc"
      );
      // console.log("looking for bool", functionRes);
      if (!functionRes.graphClientMessage) setNeedConsent(true);
      else setNeedConsent(false);
      return functionRes;
      // } // else {}
    } catch (error) {
      console.error("error in useData ques", error);
      if (error.message.includes("Access Denied")) {
        setNeedConsent(true);
      }
    }
  };

  const { loading, data: dataFromRootList, error, reload } = useData(fetchData);

  const setDataToSessionAndAppShare = (listId, quesRowId) => {
    // let modifiedData = [];
    // data.forEach((ques) => {
    //   // console.log(ques);
    //   modifiedData.push({
    //     id: ques.fields.id,
    //     [colNames[0]]: ques.fields[colNames[0]],
    //     [colNames[1]]: ques.fields[colNames[1]],
    //     [colNames[2]]: ques.fields[colNames[2]],
    //     [colNames[3]]: ques.fields[colNames[3]],
    //     [colNames[4]]: ques.fields[colNames[4]],
    //     [colNames[5]]: ques.fields[colNames[5]],
    //   });
    // });
    // // console.log("in Adminsidepanel modified", modifiedData);

    // sessionStorage.setItem("stageQuestionnaire", JSON.stringify(modifiedData));

    // createSearchParams()

    // patchQuestionnaireRootList(teamsUserCredential, quesRowId, exactDateTime);
    meeting.shareAppContentToStage((err, res) => {},
    window.location.origin + `/index.html#/questionnaire?listId=${listId}`);
    // setPageLoading(false);
  };

  /* const handleAppShare = async (e) => {
    setPageLoading(true);
    if (!teamsUserCredential) {
      throw new Error("TeamsFx SDK is not initialized.");
    }
    if (needConsent) {
      await teamsUserCredential.login(["Sites.Read.All"]); // "Sites.FullControl.All",
      setNeedConsent(false);
    }
    // sessionStorage["stageQuestionnaire"] && setDataToSessionAndAppShare([]);
    try {
      const functionRes = await getListItems(teamsUserCredential);


      !needConsent &&
        setDataToSessionAndAppShare(functionRes.graphClientMessage?.value);
      // return functionRes;
    } catch (error) {
      console.error("error in handleShare Adminsidepanel", error);
      if (error.message.includes("Access Denied")) {
        setNeedConsent(true);
      }
    }
  }; */

  useEffect(() => {
    const temp = sessionStorage.getItem("userMeetingRole");
    setUserMeetingRole(temp);
    // eslint-disable-next-line
  }, []);

  // console.log("asdf", dataFromRootList);
  return (
    <>
      {needConsent && (
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
      )}

      <SmallPopUp
        open={!loading && !dataFromRootList?.graphClientMessage.value.length}
        // onOpenChange={(e, data) => setSuccessCreate(data.open)}
        activeActions={false}
        spinner={false}
        modalType="alert"
      >
        <div className="error">
          <Text size={800}>No Questionnaire Created!</Text>
        </div>
      </SmallPopUp>

      <SmallPopUp
        className="loading"
        msg={"Fetching List of Questionnaires..."}
        open={loading}
        spinner={true}
        activeActions={false}
        modalType="alert"
      />

      {dataFromRootList &&
        (userMeetingRole === UserMeetingRole.organizer ||
        userMeetingRole === UserMeetingRole.presenter ? (
          <div className="ag-format-container">
            <div className="ag-courses_box">
              {/* // * card starts */}
              {dataFromRootList.graphClientMessage.value.map((row, ind) => {
                const field = row.fields;
                return (
                  <div className="ag-courses_item" key={ind}>
                    <div className="ag-courses-item_link">
                      <div className="ag-courses-item_bg"></div>

                      <div className="ag-courses-item_title">{field.Title}</div>

                      {/* <div className="ag-courses-item_date-box">
                        Start:
                        <span className="ag-courses-item_date">04.11.2022</span>
                      </div> */}

                      <div className="card-btn">
                        {/* {!field.isInitiated && btnClicked !== ind ? (
                          <Button
                            appearance="primary"
                            icon={<Open16Regular />}
                            onClick={(e) => {
                              setDataToSessionAndAppShare(
                                field.idOfLists,
                                field.id
                              );
                              setBtnClicked(ind);
                            }}
                          >
                            Share to Stage
                          </Button>
                        ) : (
                          <Text className="ag-courses-item_date">
                            Initiated Once
                          </Text>
                        )} */}
                        <Button
                          appearance="primary"
                          icon={<Open16Regular />}
                          onClick={(e) => {
                            setDataToSessionAndAppShare(
                              field.idOfLists,
                              field.id
                            );
                            setBtnClicked(ind);
                          }}
                        >
                          Share to Stage
                        </Button>
                      </div>
                      {/* <div className="card-btn">
                        <Button
                          appearance="primary"
                          icon={<Open16Regular />}
                          // disabled={!field.isInitiated}
                          onClick={(e) => {
                            setDataToSessionAndAppShare(
                              field.idOfLists,
                              field.id
                            );
                          }}
                        >
                          Share to Stage
                        </Button>
                      </div> */}
                    </div>
                  </div>
                );
              })}
              {/* // * card ends */}
            </div>
          </div>
        ) : (
          <h1>you are not eligible</h1>
        ))}
    </>
  );
  //   return (
  //     <>
  //       <h1>I am in AdminSidePanel</h1>
  //       {/* {meeting.getAppContentStageSharingState((err, result) => {
  //     console.log("tab.jsx", result);
  //     console.warn("tab.jsx", err);
  //     if (result.isAppSharing) {
  // // Indicates if app is sharing content on the meeting stage.
  // }
  //   })} */}
  //       {/* {meeting.getAppContentStageSharingCapabilities((err, result) => {
  //     console.log(result.doesAppHaveSharePermission);
  //   })} */}
  //       {/* {console.log("oh bhai", window.location.origin)} */}

  //       <button
  //         onClick={(e) =>
  //           meeting.shareAppContentToStage((err, res) => {},
  //           window.location.origin + "/index.html#/questionnaire")
  //         }
  //       >
  //         Click me
  //       </button>
  //     </>
  //   );
};

export default AdminSidePanel;
