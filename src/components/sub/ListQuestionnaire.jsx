import { useData } from "@microsoft/teamsfx-react";
import { getListItems } from "../../lib/utils";
import config from "../../lib/config";
import React, { useContext, useEffect, useState } from "react";
import { TeamsFxContext } from "../Context";
import SmallPopUp from "../SmallPopUp";
import { Button, Radio, RadioGroup, Text } from "@fluentui/react-components";
import ListOfQuestions from "./ListOfQuestions";

const ListQuestionnaire = () => {
  const teamsUserCredential = useContext(TeamsFxContext).teamsUserCredential;

  const [needConsent, setNeedConsent] = useState(false);
  const [creatorNameArr, setCreatorNameArr] = useState([]);
  const [questionnaireObjArr, setQuestionnaireObjArr] = useState([]);
  const [selectedCreatorName, setSelectedCreatorName] = useState("");
  const [selectedQuestionnaireId, setSelectedQuestionnaireObj] = useState("");

  const { loading, data, error, reload } = useData(async () => {
    if (!teamsUserCredential) {
      throw new Error("TeamsFx SDK is not initialized.");
    }
    if (needConsent) {
      await teamsUserCredential.login(["Sites.Read.All"]); // "Sites.FullControl.All",
      setNeedConsent(false);
    }
    try {
      // *listId of 'questionnaireRootList'
      const questionnaireRootList = await getListItems(
        teamsUserCredential,
        config.questionnaireRootListId
      );

      return questionnaireRootList;
    } catch (error) {
      console.error("error in useData ques api in analytics component", error);
      if (error.message.includes("Access Denied")) {
        setNeedConsent(true);
      }
    }
  });

  const updateCreatorNameArr = () => {
    let tempSet = new Set();
    data.graphClientMessage.value.forEach((row) => {
      tempSet.add(row.createdBy.user.displayName);
    });
    setCreatorNameArr([...tempSet]);
  };

  const updateQuestionnaireObjArr = () => {
    let tempSet = new Set();
    data.graphClientMessage.value.forEach((row) => {
      row.createdBy.user.displayName === selectedCreatorName &&
        tempSet.add({
          questionnaireName: row.fields.Title,
          questionnaireId: row.fields.idOfLists,
        });
    });
    setQuestionnaireObjArr([...tempSet]);
  };

  useEffect(() => {
    data && updateCreatorNameArr();
    return;
    // eslint-disable-next-line
  }, [data]);

  useEffect(() => {
    selectedCreatorName && updateQuestionnaireObjArr();
    return;
    // eslint-disable-next-line
  }, [selectedCreatorName]);

  console.log(data?.graphClientMessage.value);

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
          {/* loading popup */}
          <SmallPopUp
            className="loading"
            msg={"Preparing things..."}
            open={loading}
            spinner={true}
            activeActions={false}
            modalType="alert"
          />

          {!!data?.graphClientMessage.value.length && (
            <>
              <div className="flex flex-row items-center gap-12 h-full p-8">
                <div className="flex-grow">
                  <div>
                    <Button
                      onClick={(e) => {
                        setSelectedCreatorName("");
                        setSelectedQuestionnaireObj("");
                      }}
                    >
                      Clear Selection
                    </Button>
                  </div>
                  <RadioGroup
                    value={selectedCreatorName}
                    onChange={(e, data) => setSelectedCreatorName(data.value)}
                  >
                    {creatorNameArr.map((name) => (
                      <Radio label={name} value={name} key={name} />
                    ))}
                  </RadioGroup>
                </div>

                <div className="flex-grow overflow-auto">
                  {selectedCreatorName && (
                    <RadioGroup
                      value={selectedQuestionnaireId}
                      onChange={(e, data) =>
                        setSelectedQuestionnaireObj(data.value)
                      }
                    >
                      {questionnaireObjArr.map((obj) => (
                        <Radio
                          label={obj.questionnaireName}
                          value={obj.questionnaireId}
                          key={obj.questionnaireId}
                        />
                      ))}
                    </RadioGroup>
                  )}
                </div>

                <div className="flex-grow overflow-auto">
                  {selectedQuestionnaireId && (
                    <ListOfQuestions
                      selectedQuestionnaireId={selectedQuestionnaireId}
                    />
                  )}
                </div>
              </div>
            </>
          )}
        </>
      )}
    </>
  );
};

export default ListQuestionnaire;
