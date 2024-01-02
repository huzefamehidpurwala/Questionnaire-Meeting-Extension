import { Text } from "@fluentui/react-components";
import React, { useEffect, useState } from "react";
import { colNames } from "../../lib/utils";
import { UserMeetingRole } from "@microsoft/live-share";

const optionChars = ["A", "B", "C", "D"];

const QuestionnaireStage = ({
  fields,
  userRole,
  updateArrOfAnsGiven,
  ansArr,
}) => {
  const [selectedOption, setSelectedOption] = useState("");

  const checkAnsArr = () => {
    for (const ansObj of ansArr) {
      if (ansObj.questionOfQuestionnaire === fields[colNames[0]]) {
        return ansObj.ansGivenByAttendee;
      }
    }
    return "";
  };

  useEffect(() => {
    setSelectedOption(ansArr.length ? checkAnsArr() : "");
    // eslint-disable-next-line
  }, [fields]);

  return (
    <div className="relative bg-inherit p-4 min-w-[90vw] min-h-[80vh] flex flex-col justify-evenly items-center select-none">
      <div className="text-center p-5 min-w-fit">
        <Text size={600} weight="bold">
          <strong>{`Q.${fields.id})`}</strong> {fields[colNames[0]]}
        </Text>
      </div>

      <div className="grid-container relative">
        {(selectedOption || userRole === UserMeetingRole.organizer) && (
          <div className="absolute opacity-0 w-full h-full hover:cursor-not-allowed"></div>
        )}

        {optionChars.map((char, ind) => {
          const currOptionIsCorrectAns =
            fields[colNames[5]] === colNames[ind + 1];
          const isCurrOptClicked = selectedOption === colNames[ind + 1];
          return (
            <div
              className={`grid-item hover:bg-teams-bg-1 border border-slate-600 hover:border-blue-700 ${
                selectedOption
                  ? currOptionIsCorrectAns
                    ? isCurrOptClicked
                      ? "bg-green-900"
                      : "bg-green-700"
                    : isCurrOptClicked
                    ? "bg-red-900"
                    : "bg-red-700"
                  : ""
              }`}
              key={ind}
              role="button"
              onClick={(e) => {
                setSelectedOption(colNames[ind + 1]);
                updateArrOfAnsGiven(colNames[ind + 1]);
              }}
            >
              <Text size={400}>
                <strong>{char}.</strong> {fields[colNames[ind + 1]]}
              </Text>
            </div>
          );
        })}
      </div>
    </div>
  );
};

export default QuestionnaireStage;
