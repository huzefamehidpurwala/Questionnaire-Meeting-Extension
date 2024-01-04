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
      const functionRes = await getListItems(
        teamsUserCredential,
        config.questionnaireRootListId,
        "fields/isInitiated,fields/Created desc"
      );

      if (!functionRes.graphClientMessage) setNeedConsent(true);
      else setNeedConsent(false);
      return functionRes;
    } catch (error) {
      console.error("error in useData ques", error);
      if (error.message.includes("Access Denied")) {
        setNeedConsent(true);
      }
    }
  };

  const { loading, data: dataFromRootList, error, reload } = useData(fetchData);

  const setDataToSessionAndAppShare = async (listId, quesRowId, force) => {
    meeting.shareAppContentToStage((err, res) => {},
    window.location.origin + `/index.html#/questionnaire?listId=${listId}`);

    if (!force) {
      await patchQuestionnaireRootList(
        teamsUserCredential,
        quesRowId,
        exactDateTime
      );

      reload();

      // setBtnClicked("");
    }
  };

  useEffect(() => {
    const temp = sessionStorage.getItem("userMeetingRole");
    setUserMeetingRole(temp);
    // eslint-disable-next-line
  }, []);

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
        open={loading && !dataFromRootList}
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

                      <div className="card-btn">
                        {!field.isInitiated && btnClicked !== field.id ? (
                          <Button
                            appearance="primary"
                            icon={<Open16Regular />}
                            onClick={(e) => {
                              setDataToSessionAndAppShare(
                                field.idOfLists,
                                field.id,
                                false
                              );
                              setBtnClicked(field.id);
                            }}
                          >
                            Share to Stage
                          </Button>
                        ) : (
                          <Text className="ag-courses-item_date">
                            Initiated Once
                          </Text>
                        )}
                      </div>
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
};

export default AdminSidePanel;
