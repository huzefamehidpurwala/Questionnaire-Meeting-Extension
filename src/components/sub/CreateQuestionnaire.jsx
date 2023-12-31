import {
  Button,
  DrawerBody,
  DrawerHeader,
  Field,
  Input,
  InlineDrawer,
  Radio,
  RadioGroup,
  Text,
  Textarea,
  Tooltip,
  useToastController,
  Toast,
  ToastTitle,
  Toaster,
} from "@fluentui/react-components";
import {
  Add24Filled,
  Calendar20Filled,
  Delete20Regular,
} from "@fluentui/react-icons";
import React, { useContext, useEffect, useState } from "react";
import { TeamsFxContext } from "../Context";
import {
  compareObjects,
  createQuestionnaire,
  propsOfStateObj,
  toTitleCase,
  updateListFields,
} from "../../lib/utils";
import SmallPopUp from "../SmallPopUp";
import { executeDeepLink } from "@microsoft/teams-js";
import { DragDropContext, Draggable, Droppable } from "react-beautiful-dnd";

const numOfOptions = [1, 2, 3, 4];
const minValueOfId = 1000000001;
const maxValueOfId = 9999999999;
const numOfCards = 1;

const CreateQuestionnaire = ({ persnolTab }) => {
  window.onbeforeunload = function () {
    return "Your saved questions will be lost!";
  };

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
    setValueArrOfQues(
      generateRandomIntegers(numOfCards).map((idOfQues) =>
        createNewElemInValueArrOfQues(idOfQues)
      )
    );
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
      if (obj[key] && key !== propsOfStateObj[0]) {
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
      }
    }
  };

  const handleDeleteQuest = (valObj) => {
    const valueIndexOfCurrentElem = valueArrOfQues.indexOf(valObj);

    if (valueIndexOfCurrentElem !== -1 && valueArrOfQues.length > 1) {
      setSelectedQuesForEdit("");
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
      compareObjects(valObj, selectedQuesForEdit) &&
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
          console.log(
            "this will be posted to list",
            updateListFields(valueArrOfQues)
          );
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

  const onDragEnd = ({ destination, source }) => {
    if (!destination) return;

    setValueArrOfQues((t) => {
      const [reOrderedItem] = t.splice(source.index, 1);
      t.splice(destination.index, 0, reOrderedItem);
      return [...t];
    });
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

  return (
    <>
      <DragDropContext onDragEnd={onDragEnd}>
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
              executeDeepLink("https://teams.microsoft.com/_#/scheduling-form/")
            }
            icon={<Calendar20Filled />}
          >
            Schedule Meeting
          </Button>
        )}

        <div className="flex flex-col h-screen w-screen">
          <header className="header-bg w-full text-center text-white p-5">
            <span className="font-serif text-4xl underline">
              Create Questionnaire
            </span>
          </header>

          <section className="flex-grow flex w-full h-full">
            {/* SIDE PANEL */}
            <div className="">
              <InlineDrawer
                open
                separator
                style={{ width: "25vw", height: "100%" }}
              >
                <DrawerHeader>
                  <div className="flex flex-col gap-4">
                    <div className="">
                      <Field label="Name of the Questionnaire:" required>
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
                    <div className="flex justify-between">
                      <Text size={500} weight="bold">
                        Questions
                      </Text>

                      <Tooltip
                        content="Add Question"
                        withArrow
                        positioning="before"
                      >
                        <Button
                          appearance="subtle"
                          icon={<Add24Filled />}
                          onClick={handleAddQuest}
                        />
                      </Tooltip>
                    </div>
                  </div>
                </DrawerHeader>
                <DrawerBody>
                  <div className="flex flex-col gap-2 justify-between h-full">
                    <div className="flex-grow overflow-y-auto">
                      <Droppable droppableId="questions-list">
                        {(provided) => (
                          <ol
                            type="1"
                            {...provided.droppableProps}
                            ref={provided.innerRef}
                          >
                            {!!valueArrOfQues.length &&
                              valueArrOfQues.map((valObj, index) => (
                                <React.Fragment
                                  key={valObj[propsOfStateObj[0]]}
                                >
                                  <Draggable
                                    index={index}
                                    draggableId={valObj[
                                      propsOfStateObj[0]
                                    ]?.toString()}
                                  >
                                    {(provided) => (
                                      <div
                                        className={`in-draggable-wrapper list-item p-2 ${
                                          !valObj[propsOfStateObj[1]]
                                            ? "bg-for-in-draggable-wrapper"
                                            : "hover:bg-teams-bg-3"
                                        }`}
                                        {...provided.draggableProps}
                                        {...provided.dragHandleProps}
                                        ref={provided.innerRef}
                                      >
                                        <Tooltip
                                          content="Drag to re-arrange"
                                          positioning="after"
                                        >
                                          <div className="flex items-center justify-between">
                                            <Text
                                              block={true}
                                              role="button"
                                              onClick={(e) =>
                                                setSelectedQuesForEdit({
                                                  valObj,
                                                  index,
                                                })
                                              }
                                              wrap={false}
                                              truncate
                                              className={
                                                valObj[propsOfStateObj[1]]
                                                  ? "grow"
                                                  : "grow opacity-60"
                                              }
                                            >
                                              {index + 1}
                                              {". "}
                                              {valObj[propsOfStateObj[1]] ||
                                                "Please complete the question..."}
                                            </Text>
                                            <div className="question-action-btn hidden">
                                              <Tooltip
                                                withArrow
                                                content="Delete this Question?"
                                                positioning="after"
                                              >
                                                <Button
                                                  appearance="transparent"
                                                  icon={<Delete20Regular />}
                                                  onClick={(e) =>
                                                    handleDeleteQuest(valObj)
                                                  }
                                                  disabled={
                                                    valueArrOfQues.length ===
                                                      1 &&
                                                    !checkObjElemHasValue(
                                                      valObj
                                                    )
                                                  }
                                                />
                                              </Tooltip>
                                            </div>
                                          </div>
                                        </Tooltip>
                                      </div>
                                    )}
                                  </Draggable>
                                </React.Fragment>
                              ))}
                            {provided.placeholder}
                          </ol>
                        )}
                      </Droppable>
                    </div>

                    <div className="h-max">
                      <Button
                        appearance="primary"
                        className="w-full"
                        onClick={handleSubmitQues}
                      >
                        Submit Questionnaire
                      </Button>
                    </div>
                  </div>
                </DrawerBody>
              </InlineDrawer>
            </div>

            {/* // *STAGE VIEW */}
            <div className="flex-grow flex items-center justify-center">
              <Toaster toasterId={"toasterId"} position="bottom-start" />
              {selectedQuesForEdit && (
                <div className="w-4/5 h-4/5 border-2 border-gray-400 overflow-hidden">
                  <div className="bg-teams-bg-1 w-full h-full p-8">
                    <form
                      action=""
                      className="h-full"
                      onSubmit={handleSaveFormSubmit}
                    >
                      <div className="flex flex-col justify-between h-full">
                        <Field
                          label={`Question ${selectedQuesForEdit.index + 1}:`}
                          required
                          size="medium"
                        >
                          <Textarea
                            required
                            placeholder="Type here..."
                            name={propsOfStateObj[1]}
                            value={
                              selectedQuesForEdit.valObj[propsOfStateObj[1]]
                            }
                            onChange={handleChangeSelectedQuesForEdit}
                          />
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
                                <>
                                  <Text>
                                    {toTitleCase(propsOfStateObj[elem + 1])}
                                    {": "}
                                    <span className="text-red-600 text-sm">
                                      *
                                    </span>
                                  </Text>
                                  &nbsp;&nbsp;&nbsp;&nbsp;
                                  <Radio
                                    label={<Text size={200}>Correct Ans!</Text>}
                                    value={propsOfStateObj[elem + 1]}
                                  />
                                </>
                              }
                              size="medium"
                              className="w-full"
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
                  </div>
                </div>
              )}
            </div>
          </section>
        </div>
      </DragDropContext>
    </>
  );
};

export default CreateQuestionnaire;
