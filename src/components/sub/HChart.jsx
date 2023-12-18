import HighchartsReact from "highcharts-react-official";
import Highcharts from "highcharts";
import exporting from "highcharts/modules/exporting";
import { useContext, useState } from "react";
import { TeamsFxContext } from "../Context";
import { useData } from "@microsoft/teamsfx-react";
import { getListItems } from "../../lib/utils";
import SmallPopUp from "../SmallPopUp";
// import NoBgLoading from "../../assets/noBgLoading.webp";
import WhiteBgLoading from "../../assets/loading.gif";
// import Loading from "../../assets/navy_for-light_bg.webp";
import { Button, Image, Text } from "@fluentui/react-components";
exporting(Highcharts);

const HChart = ({
  questionnaireName,
  questionnaireId,
  selectedAttendeeNameArr,
  chartType,
  // selectedDateTime,
  // questionnaireRootListData,
  analyticsOfQuestionnaireData,
}) => {
  const [needConsent, setNeedConsent] = useState(false);
  const teamsUserCredential = useContext(TeamsFxContext).teamsUserCredential;
  const { loading, data, error, reload } = useData(async () => {
    if (!teamsUserCredential) {
      throw new Error("TeamsFx SDK is not initialized.");
    }
    if (needConsent) {
      await teamsUserCredential.login(["Sites.Read.All"]); // "Sites.FullControl.All",
      setNeedConsent(false);
    }
    try {
      // console.log("check list id url query", searchParams.get("listId"));
      const functionRes = await getListItems(
        teamsUserCredential,
        questionnaireId
      );
      return functionRes;
    } catch (error) {
      console.error("error in useData ques", error);
      if (error.message.includes("Access Denied")) {
        setNeedConsent(true);
      }
    }
  });

  const getValForSeries = () => {
    let series = ["Not Answered", "Incorrect Answer", "Correct Answer"].map(
      (name) => ({
        name,
        data: Array(selectedAttendeeNameArr.length).fill(0),
      })
    );

    if (data) {
      analyticsOfQuestionnaireData.graphClientMessage.value.forEach((row) => {
        const currentField = row.fields;
        if (currentField.questionnaireListId === questionnaireId) {
          if (currentField.ansIsCorrect) {
            series[2].data[
              selectedAttendeeNameArr.indexOf(currentField.attendeeName)
            ]++;
          } else {
            series[1].data[
              selectedAttendeeNameArr.indexOf(currentField.attendeeName)
            ]++;
          }
        }
      });

      selectedAttendeeNameArr.forEach((name, index) => {
        // console.log("checking inde", name, index);
        series[0].data[index] =
          data.graphClientMessage.value.length -
          (series[1].data[index] + series[2].data[index]);
      });
      // console.log("checking series", series);

      // series = ["Not Answered", "Incorrect Answer", "Correct Answer"].map(
      //   (name) => ({
      //     name,
      //     data: Array(selectedAttendeeNameArr.length).fill(10),
      //   })
      // );
    }

    return series;
  };
  // console.log("hcharts", data?.graphClientMessage.value.length);

  const options = {
    chart: {
      type: chartType,
    },
    title: {
      text: questionnaireName, // `${selectedAttendeeNameArr} - ${selectedQuestionnaireNameArr}`,
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
      categories: selectedAttendeeNameArr, // !Project Title here
      title: {
        text: "Name Of Students",
        // align: "low",
      },
      gridLineWidth: 0.5, // width of the line that seperates data
      lineWidth: 0.3, // width of the line that is known as axis
    },
    yAxis: {
      min: 0,
      title: {
        text: "No. of Answers",
        // align: "high",
      },
      labels: {
        overflow: "justify",
      },
      gridLineWidth: 0,
      lineWidth: 0.3, // width of the line that is known as axis
    },
    // the popup shown on mouse-hover
    // tooltip: {
    //   valueSuffix: " tasks",
    // },
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
    // legend: {
    //   layout: "vertical",
    //   align: "right",
    //   verticalAlign: "top",
    //   // x: -40,
    //   // y: 80,
    //   floating: true,
    //   borderWidth: 1,
    //   backgroundColor:
    //     Highcharts.defaultOptions.legend.backgroundColor || "#FFFFFF",
    //   shadow: true,
    //   // reversed: true,
    // },
    // this is the domain name shown below the chart when enabled
    credits: {
      enabled: false,
    },
    series: getValForSeries(), // || [{name: "Year 1990",data: [631, 727, 3202, 721],},{name: "Year 2000",data: [814, 841, 3714, 726],},{name: "Year 2018",data: [1276, 1007, 4561, 746],},],
  };

  return !!error || needConsent ? (
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
      <div className="relative flex flex-col justify-center items-center">
        <div className={loading ? "opacity-70" : ""}>
          <HighchartsReact
            // containerProps={{ style: { outerWidth: "100%" } }}
            highcharts={Highcharts}
            options={options}
          />
        </div>
        {loading && (
          // <div className="absolute top-28 right-56 w-28 h-28">
          //   {/* <Spinner size="huge" /> */}
          //   <Image alt="Loading..." src={NoBgLoading} />
          // </div>
            <Image src={WhiteBgLoading} className="absolute"/>
        )}
      </div>
    </>
  );
};

export default HChart;
