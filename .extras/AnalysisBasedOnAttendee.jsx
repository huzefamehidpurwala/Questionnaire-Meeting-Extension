import HighchartsReact from "highcharts-react-official";
import Highcharts from "highcharts";
import { useData } from "@microsoft/teamsfx-react";
import { TeamsFxContext } from "../Context";
import { useContext, useState, useEffect } from "react";
import { getListItems, toTitleCase } from "../../lib/utils";
import SmallPopUp from "../SmallPopUp";
import { Button, Radio, RadioGroup, Text } from "@fluentui/react-components";
import exporting from "highcharts/modules/exporting";
exporting(Highcharts);

const Analysis = () => {
  const teamsUserCredential = useContext(TeamsFxContext).teamsUserCredential;

  const [needConsent, setNeedConsent] = useState(false);
  const [attendeeNameArr, setAttendeeNameArr] = useState([]);
  const [selectedAttendeeName, setSelectedAttendeeName] = useState("");
  const [questionnaireNameArr, setQuestionnaireNameArr] = useState([]);
  // const [selectedQuestionnaireId, setSelectedQuestionnaireId] = useState("");
  // const [selectedQuestionnaireName, setSelectedQuestionnaireName] = useState("");
  const [initatedquestionnaireIdArr, setInitatedQuestionnaireIdArr] = useState(
    []
  );

  const {
    loading: analyticsOfQuestionnaireLoading,
    data: analyticsOfQuestionnaireData,
    error: analyticsOfQuestionnaireError,
    reload: analyticsOfQuestionnaireReload,
  } = useData(async () => {
    if (!teamsUserCredential) {
      throw new Error("TeamsFx SDK is not initialized.");
    }
    if (needConsent) {
      await teamsUserCredential.login(["Sites.Read.All"]); // "Sites.FullControl.All",
      setNeedConsent(false);
    }
    try {
      // * listId of 'Analytics Of Questionnaire'
      const analyticsOfQuestionnaire = await getListItems(
        teamsUserCredential,
        "d26a4a06-27e1-47cf-9782-155f265f5984"
      );

      return analyticsOfQuestionnaire;
    } catch (error) {
      console.error("error in useData ques", error);
      if (error.message.includes("The application may not be authorized.")) {
        setNeedConsent(true);
      }
    }
  });

  const {
    loading: questionnaireRootListLoading,
    data: questionnaireRootListData,
    error: questionnaireRootListError,
    reload: questionnaireRootListReload,
  } = useData(async () => {
    if (!teamsUserCredential) {
      throw new Error("TeamsFx SDK is not initialized.");
    }
    if (needConsent) {
      await teamsUserCredential.login(["Sites.Read.All"]); // "Sites.FullControl.All",
      setNeedConsent(false);
    }
    try {
      // * listId of 'questionnaireRootList'
      const questionnaireRootList = await getListItems(
        teamsUserCredential,
        "26ca1252-8fdf-467e-a1eb-172aaed95763"
      );

      return questionnaireRootList;
    } catch (error) {
      console.error("error in useData ques", error);
      if (error.message.includes("The application may not be authorized.")) {
        setNeedConsent(true);
      }
    }
  });

  const updateQuestionnaireNameArr = () => {
    let tempSet = new Set();
    questionnaireRootListData.graphClientMessage.value.forEach((row) =>
      initatedquestionnaireIdArr.includes(row.fields.idOfLists)
        ? tempSet.add(row.fields.Title) // tempSet.add({questionnaireName: row.fields.Title, questionnaireId: row.fields.idOfLists, })
        : ""
    );
    // console.log("tempSet", tempSet);
    setQuestionnaireNameArr([...tempSet]);
  };

  const updateAttendeeNameArr = () => {
    let tempSetAtt = new Set();
    let tempSetId = new Set();
    analyticsOfQuestionnaireData.graphClientMessage.value.forEach((row) => {
      tempSetAtt.add(row.fields.attendeeName);
      tempSetId.add(row.fields.questionnaireListId);
    });
    // console.log("bhag bhiya", tempSetAtt, tempSetId);
    setAttendeeNameArr([...tempSetAtt]);
    setInitatedQuestionnaireIdArr([...tempSetId]);
  };

  useEffect(() => {
    analyticsOfQuestionnaireData && updateAttendeeNameArr();
    return;
    // eslint-disable-next-line
  }, [analyticsOfQuestionnaireData]);

  useEffect(() => {
    questionnaireRootListData &&
      !!initatedquestionnaireIdArr.length &&
      updateQuestionnaireNameArr();
    return;
    // eslint-disable-next-line
  }, [questionnaireRootListData, initatedquestionnaireIdArr]);

  const getValForSeries = () => {
    const series = ["Not Answered", "Incorrect Answer", "Correct Answer"].map(
      (name) => ({ name, data: Array(questionnaireNameArr.length).fill(0) })
    );

    return series;
  };

  const options = {
    chart: {
      type: "column",
    },
    title: {
      text: selectedAttendeeName, // `${selectedAttendeeName} - ${selectedQuestionnaireName}`,
      // align: "left",
    },
    // subtitle: {
    //   text: '<div style="color: red">Huzefa</div>',
    //   align: "left",
    // },
    // navigation: {
    //   buttonOptions: {
    //     enabled: true,
    //   },
    // },
    exporting: {
      enabled: true,
    },
    xAxis: {
      categories: questionnaireNameArr, // !Project Title here
      title: {
        text: "Name Of Questionnaires",
        // align: "low",
      },
      gridLineWidth: 0.5, // width of the line that seperates data
      lineWidth: 0.3, // width of the line that is known as axis
    },
    yAxis: {
      min: 0,
      title: {
        text: "Number of Questions",
        // align: "high",
      },
      labels: {
        overflow: "justify",
      },
      gridLineWidth: 0,
      lineWidth: 0.3, // width of the line that is known as axis
    },
    // the popup shown on mouse-hover
    tooltip: {
      valueSuffix: " tasks",
    },
    accessibility: {
      enabled: false, // to supress the console warning
    },
    plotOptions: {
      bar: {
        // borderRadius: "50%", // this is to give round edges
        dataLabels: {
          enabled: true, // to show the value of data after the bar
        },
        groupPadding: 0.1, // this defines the thickness of bar
      },
      // this is to stack up the series
      // series: {
      //   stacking: "normal",
      //   dataLabels: {
      //     enabled: true,
      //   },
      // },
    },
    // this is the card shown that helps to understand which color represents the respective data
    legend: {
      layout: "vertical",
      align: "right",
      verticalAlign: "top",
      x: -40,
      y: 80,
      floating: true,
      borderWidth: 1,
      backgroundColor:
        Highcharts.defaultOptions.legend.backgroundColor || "#FFFFFF",
      shadow: true,
      // reversed: true,
    },
    // this is the domain name shown below the chart when enabled
    credits: {
      enabled: false,
    },
    series: getValForSeries(), // || [{name: "Year 1990",data: [631, 727, 3202, 721],},{name: "Year 2000",data: [814, 841, 3714, 726],},{name: "Year 2018",data: [1276, 1007, 4561, 746],},],
  };

  // console.log("in analytics global==", questionnaireNameArr);
  return (
    <>
      {!!questionnaireRootListError ||
      !!analyticsOfQuestionnaireError ||
      needConsent ? (
        <SmallPopUp
          open={
            !!questionnaireRootListError ||
            !!analyticsOfQuestionnaireError ||
            needConsent
          }
          activeActions={false}
          spinner={false}
          modalType="alert"
        >
          <div className="error">
            <Text size={800}>Fetching Data Failed</Text>
            <br />
            <br />
            <Button
              appearance="primary"
              disabled={questionnaireRootListLoading}
              onClick={questionnaireRootListReload}
            >
              Authorize and call Azure Function
            </Button>{" "}
            <Button
              appearance="primary"
              disabled={analyticsOfQuestionnaireLoading}
              onClick={analyticsOfQuestionnaireReload}
            >
              Authorize and call Azure Function
            </Button>
          </div>
        </SmallPopUp>
      ) : (
        <>
          <SmallPopUp
            className="loading"
            msg={"Preparing Charts..."}
            open={
              analyticsOfQuestionnaireLoading || questionnaireRootListLoading
            }
            spinner={true}
            activeActions={false}
            modalType="alert"
          />

          {!!questionnaireRootListData && !!analyticsOfQuestionnaireData && (
            <>
              <div className="analytics-flex-container">
                <div className="sub-container1">
                  <div>
                    <Button
                      appearance="transparent"
                      onClick={(e) => {
                        setSelectedAttendeeName("");
                        // setSelectedQuestionnaireName("");
                        // setSelectedQuestionnaireId("");
                      }}
                    >
                      Clear Selection
                    </Button>
                  </div>

                  <div>
                    <Text size={500}>Select Attendee Name:</Text>
                    <RadioGroup
                      value={selectedAttendeeName}
                      onChange={(e, data) =>
                        setSelectedAttendeeName(data.value)
                      }
                    >
                      {attendeeNameArr.map((name, ind) => (
                        <Radio
                          label={<Text size={400}>{name}</Text>}
                          value={name}
                          key={ind}
                        />
                      ))}
                    </RadioGroup>
                  </div>

                  {/* <div>
                    <Text size={500}>Select Questionnaire:</Text>
                    <RadioGroup
                      value={selectedQuestionnaireName}
                      onChange={(e, data) => {
                        // console.log("in the chaneg", e.target);
                        setSelectedQuestionnaireName(data.value);
                        setSelectedQuestionnaireId(e.target.id);
                      }}
                    >
                      {questionnaireNameArr.map((name, ind) => (
                        <Radio
                          label={
                            <Text size={400}>
                              {toTitleCase(name.questionnaireName)}
                            </Text>
                          }
                          value={name.questionnaireName}
                          id={name.questionnaireId}
                          key={ind}
                        />
                      ))}
                    </RadioGroup>
                  </div> */}
                </div>
                <div className="sub-container2">
                  {selectedAttendeeName && (
                    /* selectedQuestionnaireId && */ <HighchartsReact
                      // containerProps={{ style: { outerWidth: "100%" } }}
                      highcharts={Highcharts}
                      options={options}
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

export default Analysis;
