import {
  Body1,
  Body2,
  Button,
  Card,
  CardFooter,
  Field,
  Input,
  Radio,
  RadioGroup,
  Text,
  Textarea,
  Tooltip,
} from "@fluentui/react-components";
import {
  Add24Filled,
  Calendar20Filled,
  Delete24Regular,
} from "@fluentui/react-icons";
import { useContext, useState } from "react";
import { TeamsFxContext } from "../Context";
import {
  createQuestionnaire,
  propsOfStateObj,
  toTitleCase,
  updateListFields,
} from "../../lib/utils";
import SmallPopUp from "../SmallPopUp";
import { executeDeepLink } from "@microsoft/teams-js";

const numOfOptions = [1, 2, 3, 4];
const minValueOfId = 1000000001;
const maxValueOfId = 9999999999;
const numOfCards = 3;

const CreateQuestionnaire = ({ persnolTab }) => {
  const teamsUserCredential = useContext(TeamsFxContext).teamsUserCredential;
  // *functions
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

  const getValFromValueArrOfQues = (
    idFromValToRetrieve,
    valToRetrieve = ""
  ) => {
    for (const elem of valueArrOfQues) {
      if (elem[propsOfStateObj[0]] === idFromValToRetrieve) {
        if (valToRetrieve) return elem[valToRetrieve];
        else return elem;
      }
    }
  };

  const setValForValueArrOfQues = (
    idForValToUpdate,
    propToUpdate,
    newValue
  ) => {
    // ? call handleDelete to delete the existing item
    // console.log("outside set", idForValToUpdate);
    setValueArrOfQues((t) => {
      // console.log("keseho t===", t);
      const index = t.findIndex(
        (obj) => obj[propsOfStateObj[0]] === idForValToUpdate
      );
      // console.log("keseho index===", index);

      if (index !== -1) {
        // Replace the existing object with the new object
        const newData = [...t];
        // console.log("in replace", newData);
        newData[index][propToUpdate] = newValue;
        // console.log("in replace", newData);
        return newData;
      } else {
        console.error(
          "not in replace. error in setValForValueArrOfQues function in Questionnaire component"
        );
        // Add the new object to the array
        return [...t];
      }
    });
  };

  const createNewElemInValueArrOfQues = (uniqueId) => {
    let tempObj = {};
    for (const tempElem of propsOfStateObj) {
      tempObj[tempElem] = tempElem === propsOfStateObj[0] ? uniqueId : ""; // add this after unique id to give radio btn a default val tempElem === propsOfStateObj[6] ? propsOfStateObj[2] :
    }
    return tempObj;
  };

  const checkObjElemHasValue = (obj) => {
    if (!obj) return;
    let check = false;
    for (const key of Object.keys(obj)) {
      // console.log("fas", obj[key]);
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
    while (check && checkNum < idArrOfQues.length) {
      checkNum++;
      const randomInt = Math.floor(
        Math.random() * (maxValueOfId - minValueOfId + 1) + minValueOfId
      );

      if (!idArrOfQues.includes(randomInt)) {
        check = false;
        setIdArrOfQues((t) => [...t, randomInt]);
        setValueArrOfQues((t) => [
          ...t,
          createNewElemInValueArrOfQues(randomInt),
        ]);
      }
    }
  };

  const handleDeleteQuest = (e) => {
    !e.target.id &&
      console.error("you clicked this element without id", e.target);

    const currentElem = parseInt(e.target.id);
    const idIndexOfCurrentElem = idArrOfQues.indexOf(currentElem);
    const valueIndexOfCurrentElem = valueArrOfQues.indexOf(
      getValFromValueArrOfQues(currentElem)
    );
    // console.log("delete", currentElem, valueIndexOfCurrentElem);

    if (
      idIndexOfCurrentElem !== -1 &&
      idArrOfQues.length > 1 &&
      valueArrOfQues.length > 1
    ) {
      setIdArrOfQues((t) => {
        t.splice(idIndexOfCurrentElem, 1);
        return [...t];
      });

      setValueArrOfQues((t) => {
        t.splice(valueIndexOfCurrentElem, 1);
        return [...t];
      });
    } else if (
      idArrOfQues.length === 1 &&
      valueArrOfQues.length === 1 &&
      checkObjElemHasValue(valueArrOfQues[0])
    ) {
      setValueArrOfQues([createNewElemInValueArrOfQues(idArrOfQues[0])]);
    }
  };

  const resetStates = () => {
    const updatedArr = generateRandomIntegers(numOfCards);
    setIdArrOfQues([...updatedArr]);
    setValueArrOfQues(
      updatedArr.map((idOfQues) => createNewElemInValueArrOfQues(idOfQues))
    );
    setNameOfQuestionnaire("");
  };

  const handleFormSubmit = async (e) => {
    e.preventDefault();
    setPageLoading(true);

    let columns = [];
    for (const name of propsOfStateObj) {
      if (name === propsOfStateObj[1] || name === propsOfStateObj[0]) continue;
      columns.push({ name, text: {} });
    }

    const newList = {
      displayName: toTitleCase(nameOfQuestionnaire),
      columns,
      list: {
        template: "genericList",
      },
    };

    // console.log("checking again", newList, updateListFields(valueArrOfQues));

    try {
      /* const response = */ await createQuestionnaire(
        teamsUserCredential,
        newList,
        updateListFields(valueArrOfQues)
      );
      setSuccessCreate(true);
      // resetStates();
      // console.log("Hello World", response);
    } catch (error) {
      console.error("api error in createQuestionnaire", error);
      if (error.message.includes("Name already exists")) {
        setQuestionnaireExists(true);
      }
    }
    // console.log("global value", valueArrOfQues);
    setPageLoading(false);
  };

  // *states
  const [idArrOfQues, setIdArrOfQues] = useState(
    generateRandomIntegers(numOfCards)
  );
  const [valueArrOfQues, setValueArrOfQues] = useState(
    idArrOfQues.map((idOfQues) => createNewElemInValueArrOfQues(idOfQues))
  );
  const [pageLoading, setPageLoading] = useState(false);
  const [questionnaireExists, setQuestionnaireExists] = useState(false);
  const [successCreate, setSuccessCreate] = useState(false);
  const [nameOfQuestionnaire, setNameOfQuestionnaire] = useState("");

  // *create site in sharepoint
  // const setIsQuestionnaireSitePresent =
  //   useContext(TeamsFxContext).setIsQuestionnaireSitePresent;
  // useEffect(() => {
  //   setIsQuestionnaireSitePresent(false);
  //   // eslint-disable-next-line
  // }, []);

  // console.log("global id1", valueArrOfQues);
  // console.log("global value", idArrOfQues);
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

      <header className="header-bg w-full text-center text-white mb-5 p-5">
        <span className="font-serif text-4xl underline">
          Create Questionnaire
        </span>
      </header>

      {persnolTab && (
        <div className="fixed bottom-8 left-8 z-10">
          <Button
            appearance="primary"
            // size="large"
            // shape="circular"
            onClick={() =>
              executeDeepLink("https://teams.microsoft.com/_#/scheduling-form/")
            }
            icon={<Calendar20Filled />}
          >
            Schedule Meeting
          </Button>
        </div>
      )}

      {!!idArrOfQues && !!valueArrOfQues && (
        <form action="" onSubmit={handleFormSubmit}>
          <div className="questionnaire-form-flex">
            <div className="w-1/2 max-w-xl">
              <Field label="Name of the Questionnaire:" required size="large">
                <Input
                  required
                  type="text"
                  placeholder="Type here..."
                  value={nameOfQuestionnaire}
                  onChange={(e) => setNameOfQuestionnaire(e.target.value)}
                />
              </Field>
            </div>

            <div className="question-card-flex">
              {idArrOfQues.map((idOfQues, indexOfMap) => (
                <Card
                  key={idOfQues}
                  className=""
                  id={`card-${idOfQues}`} /* style={{width: "35%"}} */
                >
                  <Body1>
                    <Field
                      label={`Question ${indexOfMap + 1}:`}
                      required
                      size="medium"
                    >
                      <Textarea
                        required
                        placeholder="Type here..."
                        value={getValFromValueArrOfQues(
                          idOfQues,
                          propsOfStateObj[1]
                        )}
                        onChange={(e) =>
                          setValForValueArrOfQues(
                            idOfQues,
                            propsOfStateObj[1],
                            e.target.value
                          )
                        }
                      />
                    </Field>

                    <Body2>
                      <RadioGroup
                        required
                        value={getValFromValueArrOfQues(
                          idOfQues,
                          propsOfStateObj[6]
                        )}
                        onChange={(e, data) =>
                          setValForValueArrOfQues(
                            idOfQues,
                            propsOfStateObj[6]?.toString().replace(/\s/gm, ""),
                            data.value
                          )
                        }
                      >
                        <div className="question-card-grid-container">
                          {numOfOptions.map((elem, key) => (
                            <div className="grid-item" key={key}>
                              <div className="options-field">
                                <Field
                                  label={
                                    <>
                                      <Text>
                                        {toTitleCase(propsOfStateObj[elem + 1])}
                                      </Text>
                                      &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                                      <Radio
                                        label={
                                          <Text size={200}>Correct Ans!</Text>
                                        }
                                        value={propsOfStateObj[elem + 1]}
                                      />
                                    </>
                                  }
                                  required
                                  size="medium"
                                  key={key}
                                  className="options-field"
                                >
                                  <Input
                                    required
                                    className="options-field"
                                    type="text"
                                    placeholder="Type here..."
                                    value={getValFromValueArrOfQues(
                                      idOfQues,
                                      propsOfStateObj[elem + 1]
                                    )}
                                    onChange={(e) =>
                                      setValForValueArrOfQues(
                                        idOfQues,
                                        propsOfStateObj[elem + 1],
                                        e.target.value
                                      )
                                    }
                                  />
                                </Field>
                              </div>
                            </div>
                          ))}
                        </div>
                      </RadioGroup>
                    </Body2>
                  </Body1>

                  <CardFooter>
                    <div className="card-footer-flex">
                      <div>
                        <Tooltip
                          withArrow
                          content="Delete this Question?"
                          positioning="below-end"
                        >
                          <Button
                            id={idOfQues}
                            icon={<Delete24Regular id={idOfQues} />}
                            className="justify-self-end"
                            // appearance="transparent"
                            onClick={handleDeleteQuest}
                            // ?need to solve the error that null or undefined is not iterable
                            disabled={
                              idArrOfQues.length === 1 &&
                              valueArrOfQues.length === 1 &&
                              !checkObjElemHasValue(
                                getValFromValueArrOfQues(idOfQues)
                              )
                            }
                          />
                        </Tooltip>
                      </div>
                    </div>
                  </CardFooter>
                </Card>
              ))}

              <div className="flex flex-col justify-center items-center gap-2">
                <Tooltip withArrow content="Add Question">
                  {/* <div> */}
                  <Button
                    icon={<Add24Filled />}
                    size="large"
                    shape="circular"
                    // className="min-w-full h-full"
                    onClick={handleAddQuest}
                    appearance="primary"
                  />
                  {/* </div> */}
                </Tooltip>

                {/* <div className="">
                  <Text>Add Question</Text>
                </div> */}
              </div>
            </div>

            <div className="questionnaire-submit-btn">
              <Button size="medium" appearance="primary" type="submit">
                Submit Questionnaire
              </Button>
            </div>
          </div>
        </form>
      )}
    </>
  );
};

export default CreateQuestionnaire;
