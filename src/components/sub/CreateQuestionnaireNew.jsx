import {
  Button,
  Field,
  Input,
  Radio,
  RadioGroup,
  Text,
  // eslint-disable-next-line
  Textarea,
  Tooltip,
  useToastController,
  Toast,
  ToastTitle,
  Toaster,
  Subtitle1,
  Divider,
} from "@fluentui/react-components";
import {
  Add24Regular,
  ArrowSort28Regular,
  ArrowSortDown24Regular,
  ArrowSortUp24Regular,
  Calendar20Filled,
  Delete20Regular,
} from "@fluentui/react-icons";
import React, { useContext, useEffect, useState } from "react";
import { TeamsFxContext } from "../Context";
import {
  createQuestionnaire,
  propsOfStateObj,
  toTitleCase,
  updateListFields,
} from "../../lib/utils";
import SmallPopUp from "../SmallPopUp";
import { app } from "@microsoft/teams-js";
// eslint-disable-next-line
import { DragDropContext, Draggable, Droppable } from "react-beautiful-dnd";
import { useCreateQuestionnaireCss } from "../../styles";

const numOfOptions = [1, 2, 3, 4];
const minValueOfId = 1000000001;
const maxValueOfId = 9999999999;
const numOfCards = 1;

const CreateQuestionnaireNew = ({ persnolTab }) => {
  window.onbeforeunload = function () {
    return "Your saved questions will be lost!";
  };

  const thisStyles = useCreateQuestionnaireCss();

  const teamsUserCredential = useContext(TeamsFxContext).teamsUserCredential;

  const generateRandomIntegers = (lengthOfArr) => {
    const result = new Set();
    while (result.size < lengthOfArr) {
      const randomInt = Math.floor(
        Math.random() * (maxValueOfId - minValueOfId + 1) + minValueOfId
      );
      result.add(randomInt);
    }
    return Array.from(result);
  };

  const createNewElemInValueArrOfQues = (uniqueId) => {
    let tempObj = {};
    for (const tempElem of propsOfStateObj) {
      tempObj[tempElem] = tempElem === propsOfStateObj[0] ? uniqueId : "";
    }
    return tempObj;
  };

  // *states
  const [valueArrOfQues, setValueArrOfQues] = useState([]);
  const [pageLoading, setPageLoading] = useState(false);
  const [questionnaireExists, setQuestionnaireExists] = useState(false);
  const [successCreate, setSuccessCreate] = useState(false);
  const [nameOfQuestionnaire, setNameOfQuestionnaire] = useState("");
  const [selectedQuesForEdit, setSelectedQuesForEdit] = useState("");

  useEffect(() => {
    let tempArr = generateRandomIntegers(numOfCards);
    tempArr = tempArr.map((idOfQues) =>
      createNewElemInValueArrOfQues(idOfQues)
    );

    setValueArrOfQues(tempArr);

    const index = 0;
    setSelectedQuesForEdit({ index, valObj: tempArr[index] });
  }, []);

  // *toast
  const { dispatchToast } = useToastController("toasterId");
  const notify = (msg, intent) =>
    dispatchToast(
      <Toast>
        <ToastTitle>{msg}</ToastTitle>
      </Toast>,
      { intent }
    );

  // *functions
  const checkObjElemHasValue = (obj) => {
    if (!obj) return;
    let check = false;
    for (const key of Object.keys(obj)) {
      if (key !== propsOfStateObj[0] && obj[key]) {
        check = true;
        break;
      }
    }

    return check;
  };

  const handleAddQuest = (e) => {
    let check = true;
    let checkNum = 0;
    const tempArr = valueArrOfQues.map((valObj) => valObj[propsOfStateObj[0]]);
    while (check && checkNum < valueArrOfQues.length) {
      checkNum++;
      const randomInt = Math.floor(
        Math.random() * (maxValueOfId - minValueOfId + 1) + minValueOfId
      );

      if (!tempArr.includes(randomInt)) {
        check = false;

        setValueArrOfQues((t) => [
          ...t,
          createNewElemInValueArrOfQues(randomInt),
        ]);

        !selectedQuesForEdit &&
          setSelectedQuesForEdit({
            valObj: createNewElemInValueArrOfQues(randomInt),
            index: valueArrOfQues.length,
          });
      }
    }
  };

  const handleDeleteQuest = (valObj) => {
    const valueIndexOfCurrentElem = valueArrOfQues.indexOf(valObj);

    if (valueIndexOfCurrentElem !== -1 && valueArrOfQues.length > 1) {
      setValueArrOfQues((t) => {
        t.splice(valueIndexOfCurrentElem, 1);
        return [...t];
      });
    } else if (
      valueArrOfQues.length === 1 &&
      checkObjElemHasValue(valueArrOfQues[0])
    ) {
      setValueArrOfQues([
        createNewElemInValueArrOfQues(valObj[propsOfStateObj[0]]),
      ]);
    }

    selectedQuesForEdit &&
      valObj[propsOfStateObj[0]] ===
        selectedQuesForEdit.valObj[propsOfStateObj[0]] &&
      setSelectedQuesForEdit("");
  };

  const resetStates = () => {
    const updatedArr = generateRandomIntegers(numOfCards);
    setValueArrOfQues(
      updatedArr.map((idOfQues) => createNewElemInValueArrOfQues(idOfQues))
    );
    setNameOfQuestionnaire("");
  };

  const handleSubmitQues = async (e) => {
    let elemFoundEmpty = false;
    if (nameOfQuestionnaire) {
      for (const valObj of valueArrOfQues) {
        if (!checkObjElemHasValue(valObj)) {
          elemFoundEmpty = true;
          break;
        }
      }

      if (!elemFoundEmpty) {
        setPageLoading(true);
        let columns = [];
        for (const name of propsOfStateObj) {
          if (name === propsOfStateObj[1] || name === propsOfStateObj[0])
            continue;
          columns.push({ name, text: {} });
        }

        const newList = {
          displayName: toTitleCase(nameOfQuestionnaire),
          columns,
          list: {
            template: "genericList",
          },
        };

        try {
          /* const response = */ await createQuestionnaire(
            teamsUserCredential,
            newList,
            updateListFields(valueArrOfQues)
          );
          // console.log("this will be posted to list", updateListFields(valueArrOfQues));
          setSuccessCreate(true);
        } catch (error) {
          console.error("api error in createQuestionnaire", error);
          if (error.message.includes("Name already exists")) {
            setQuestionnaireExists(true);
          }
        }
      } else notify("There is an empty question!", "error");
    } else notify("Give a name for this Questionnaire!");
    setPageLoading(false);
  };

  const handleChangeSelectedQuesForEdit = (e, data) => {
    setSelectedQuesForEdit(({ valObj, index }) => ({
      valObj: { ...valObj, [e.target.name]: data.value.trimStart() },
      index,
    }));
  };

  const getInd = (id) => {
    for (let i = 0; i < valueArrOfQues.length; i++) {
      if (valueArrOfQues[i][propsOfStateObj[0]] === id) {
        return i;
      }
    }
    return -1;
  };

  const handleSaveFormSubmit = (e) => {
    e.preventDefault();

    const ind = getInd(selectedQuesForEdit.valObj[propsOfStateObj[0]]);
    ind !== -1 &&
      setValueArrOfQues((t) => {
        t[ind] = selectedQuesForEdit.valObj;
        return t;
      });

    setSelectedQuesForEdit("");
    notify("Question Saved!", "success");
  };

  const halfOpacity = (e, parentId) => {
    e.preventDefault();

    document.getElementById(parentId.toString()).style.opacity = "0.5";
  };

  const fullOpacity = (e, parentId, onDrop = false) => {
    e.preventDefault();

    // if (e.target.id || onDrop) {
    document.getElementById(parentId.toString()).style.opacity = "1";
    // }
  };

  const onDragEnd = ({ destination, source }) => {
    // if (!destination) return;
    // console.log("finally dropped success", { destination, source });

    setValueArrOfQues((t) => {
      const [reOrderedItem] = t.splice(source, 1); // source.index
      t.splice(destination, 0, reOrderedItem); // destination.index
      return [...t];
    });
  };

  const dummyFunc = (e) => e.preventDefault();

  return (
    <>
      {/* loading */}
      <SmallPopUp
        className="loading"
        msg={"Updating..."}
        open={pageLoading}
        spinner={true}
        activeActions={false}
        modalType="alert"
      />

      {/* success popup */}
      <SmallPopUp
        open={successCreate}
        onOpenChange={(e, data) => {
          setSuccessCreate(data.open);
          resetStates();
        }}
        activeActions={true}
        spinner={false}
        modalType="alert"
      >
        <div className="success-questionnaire">
          <Text size={800}>
            The Questionnaire
            <br />
            <u>
              <i>{toTitleCase(nameOfQuestionnaire)}</i>
            </u>
            <br />
            was successfully created!
          </Text>
        </div>
      </SmallPopUp>

      {/* questionnaire exist popup */}
      <SmallPopUp
        open={questionnaireExists}
        onOpenChange={(e, data) => setQuestionnaireExists(data.open)}
        activeActions={true}
        spinner={false}
        modalType="alert"
      >
        <div className="error">
          <Text size={800}>
            The Questionnaire
            <br />
            <u>
              <i>{toTitleCase(nameOfQuestionnaire)}</i>
            </u>
            <br />
            already exists!
          </Text>
        </div>
      </SmallPopUp>

      {/* schedule meeting btn */}
      {persnolTab && (
        <Button
          className="fixed bottom-8 right-8 z-10"
          appearance="primary"
          onClick={() =>
            app.openLink("https://teams.microsoft.com/_#/scheduling-form/")
          }
          icon={<Calendar20Filled />}
        >
          Schedule Meeting
        </Button>
      )}

      <div className="w-[100vw] h-screen flex overflow-hidden">
        {/* // * SIDE-BAR */}
        <div className="w-2/5">
          <div className="flex flex-col overflow-auto h-full">
            <div className="my-8 mx-auto w-5/6">
              <Field
                // label={<Text weight="semibold">Name of the Questionnaire:</Text>}
                label={"Name of the Questionnaire:"}
                required
                size="large"
                className={thisStyles.questionnaireName}
              >
                <Input
                  required
                  type="text"
                  placeholder="Type here..."
                  value={nameOfQuestionnaire}
                  onChange={(e) =>
                    setNameOfQuestionnaire(e.target.value.trimStart())
                  }
                />
              </Field>
            </div>

            <div className="mb-4 mx-auto w-5/6 flex items-center justify-between">
              <Subtitle1>Questions</Subtitle1>

              <Tooltip content="Add Question" withArrow positioning="before">
                <Button
                  appearance="subtle"
                  icon={<Add24Regular />}
                  onClick={handleAddQuest}
                />
              </Tooltip>
            </div>

            <div className="grow overflow-auto">
              {!!valueArrOfQues.length &&
                valueArrOfQues.map((valObj, index) => (
                  <React.Fragment key={valObj[propsOfStateObj[0]]}>
                    <div
                      id={valObj[propsOfStateObj[0]].toString()}
                      className={`text-white bg-[#5A80BE] hover:bg-blue-500 flex items-center`} // ${!valObj[propsOfStateObj[1]] ? "text-red-500" : ""}
                      // onDragEnter={(e) =>
                      //   halfOpacity(e, valObj[propsOfStateObj[0]])
                      // }
                      onDragLeave={(e) =>
                        fullOpacity(e, valObj[propsOfStateObj[0]])
                      }
                      onDragOver={(e) =>
                        halfOpacity(e, valObj[propsOfStateObj[0]])
                      }
                      droppable="true"
                      onDrop={(e) => {
                        e.preventDefault();
                        onDragEnd({
                          destination: index,
                          source: e.dataTransfer.getData("quesInd"),
                        });
                        fullOpacity(e, valObj[propsOfStateObj[0]], true);
                      }}
                    >
                      <Text
                        block
                        role="button"
                        onClick={(e) =>
                          setSelectedQuesForEdit({ valObj, index })
                        }
                        wrap={false}
                        truncate
                        onDragLeave={dummyFunc}
                        className={`px-3 py-2 grow bg-inherit ${
                          !valObj[propsOfStateObj[1]] && "text-red-600"
                        }`}
                      >
                        {`${index + 1}. `}
                        {valObj[propsOfStateObj[1]] ||
                          "Please complete the question..!"}
                      </Text>

                      <div
                        className="flex bg-white px-1 py-1"
                        onDragLeave={dummyFunc}
                      >
                        {valueArrOfQues.length > 1 && (
                          <Tooltip
                            withArrow
                            content="Drag this Question?"
                            positioning="after"
                          >
                            <Button
                              appearance="subtle"
                              icon={
                                index === 0 ? (
                                  <ArrowSortDown24Regular />
                                ) : index === valueArrOfQues.length - 1 ? (
                                  <ArrowSortUp24Regular />
                                ) : (
                                  <ArrowSort28Regular />
                                )
                              }
                              // disabled={valueArrOfQues.length <= 1}
                              draggable={true}
                              onDragStart={(e) =>
                                e.dataTransfer.setData("quesInd", index)
                              }
                            />
                          </Tooltip>
                        )}

                        <Tooltip
                          withArrow
                          content="Delete this Question?"
                          positioning="after"
                        >
                          <Button
                            appearance="subtle"
                            icon={<Delete20Regular />}
                            onClick={(e) => handleDeleteQuest(valObj)}
                            disabled={
                              valueArrOfQues.length === 1 &&
                              !checkObjElemHasValue(valObj)
                            }
                          />
                        </Tooltip>
                      </div>
                    </div>

                    {/* // * separater */}
                    <>
                      {index !== valueArrOfQues.length - 1 && (
                        <div className="flex-container min-h-[10px]">
                          <Divider inset appearance="strong" />
                        </div>
                      )}
                    </>
                  </React.Fragment>
                ))}
            </div>

            <div className="h-max w-5/6 mx-auto my-4 self-end">
              <Button
                appearance="primary"
                size="large"
                className="w-full"
                onClick={handleSubmitQues}
              >
                Submit Questionnaire
              </Button>
            </div>
          </div>
        </div>

        {/* // * QUESTION INPUT STAGE */}
        <div className="bg-[#5A80BE] w-full h-full flex-container">
          <Toaster toasterId={"toasterId"} position="bottom" />

          {!selectedQuesForEdit ? (
            <div>Lorem ipsum dolor sit amet.</div>
          ) : (
            <div className="w-1/2 h-fit border-2 border-gray-400 overflow-hidden bg-teams-bg-1 p-8 rounded-2xl">
              {/* <div className="bg-teams-bg-1 w-full h-full p-8"> */}
              <form
                action=""
                className="h-full flex flex-col justify-center"
                onSubmit={handleSaveFormSubmit}
              >
                <div className="flex flex-col justify-between gap-4">
                  <Field
                    label={`Question ${selectedQuesForEdit.index + 1}:`}
                    required
                    size="medium"
                  >
                    <Textarea
                      required
                      placeholder="Type here..."
                      name={propsOfStateObj[1]}
                      value={selectedQuesForEdit.valObj[propsOfStateObj[1]]}
                      onChange={handleChangeSelectedQuesForEdit}
                    />
                    {/* <Input
                      required
                      placeholder="Type here..."
                      name={propsOfStateObj[1]}
                      value={selectedQuesForEdit.valObj[propsOfStateObj[1]]}
                      onChange={handleChangeSelectedQuesForEdit}
                    /> */}
                  </Field>

                  <RadioGroup
                    required
                    name={propsOfStateObj[6]}
                    value={selectedQuesForEdit.valObj[propsOfStateObj[6]]}
                    onChange={handleChangeSelectedQuesForEdit}
                  >
                    {numOfOptions.map((elem, key) => (
                      <Field
                        key={key}
                        label={
                          <div className="flex justify-between items-center">
                            <Text>
                              {toTitleCase(propsOfStateObj[elem + 1])}
                              {": "}
                              <span className="text-red-600 text-sm">*</span>
                            </Text>
                            {/* &nbsp;&nbsp;&nbsp;&nbsp; */}
                            <Radio
                              label={<Text size={200}>Correct Answer!</Text>}
                              value={propsOfStateObj[elem + 1]}
                            />
                          </div>
                        }
                        size="medium"
                        className="w-full mb-2"
                      >
                        <Input
                          required
                          className="options-field"
                          type="text"
                          placeholder="Type here..."
                          name={propsOfStateObj[elem + 1]}
                          value={
                            selectedQuesForEdit.valObj[
                              propsOfStateObj[elem + 1]
                            ]
                          }
                          onChange={handleChangeSelectedQuesForEdit}
                        />
                      </Field>
                    ))}
                  </RadioGroup>

                  <div className="text-end">
                    <Button type="submit" appearance="primary">
                      Save
                    </Button>
                  </div>
                </div>
              </form>
              {/* </div> */}
            </div>
          )}
        </div>
      </div>
    </>
  );
};

export default CreateQuestionnaireNew;
