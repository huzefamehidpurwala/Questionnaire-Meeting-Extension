import HighchartsReact from "highcharts-react-official";
import Highcharts from "highcharts";
import exporting from "highcharts/modules/exporting";
import { useContext, useState } from "react";
import { TeamsFxContext } from "../Context";
import { useData } from "@microsoft/teamsfx-react";
import { getListItems, handleStringSort } from "../../lib/utils";
import SmallPopUp from "../SmallPopUp";
import WhiteBgLoading from "../../assets/loading.gif";
import { Button, Image, Text } from "@fluentui/react-components";
exporting(Highcharts);

const HChart = ({
  questionnaireName,
  questionnaireId,
  selectedAttendeeNameArr,
  chartType,
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
        series[0].data[index] =
          data.graphClientMessage.value.length -
          (series[1].data[index] + series[2].data[index]);
      });
    }

    return series;
  };

  const options = {
    chart: {
      type: chartType,
    },
    title: {
      text: questionnaireName,
    },
    exporting: {
      enabled: true,
    },
    xAxis: {
      // !Project Title here
      categories: selectedAttendeeNameArr.sort((a, b) =>
        handleStringSort(a, b)
      ), 
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
    },
    credits: {
      enabled: false,
    },
    series: getValForSeries(),
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
          <HighchartsReact highcharts={Highcharts} options={options} />
        </div>
        {loading && <Image src={WhiteBgLoading} className="absolute" />}
      </div>
    </>
  );
};

export default HChart;
