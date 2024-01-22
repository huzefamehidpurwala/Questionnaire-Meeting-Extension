import { Text, mergeClasses } from "@fluentui/react-components";
import React, { useEffect, useState } from "react";
import { colNames, isAdmin } from "../../lib/utils";
import { useQuestionnaireCss } from "../../styles";

const optionChars = ["A", "B", "C", "D"];

const NewQuestionnaireStage = ({
  fields,
  userRole,
  updateArrOfAnsGiven,
  ansArr,
}) => {
  const questStyles = useQuestionnaireCss();

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
    <>
      <div className="flex items-start gap-x-6 bg-white text-black w-[90vw] flex-grow rounded-3xl px-10 py-14">
        <div className="flex justify-center items-center bg-[#323350] text-white px-4 py-2 rounded-xl min-w-fit">
          <Text size={500}>Q - {fields.id}</Text>
        </div>

        <div className="flex flex-col gap-y-5">
          <Text size={700} weight="semibold">
            {fields[colNames[0]]}
          </Text>

          <div className="flex flex-col gap-y-3 relative">
            {(selectedOption || isAdmin(userRole)) && <div className="absolute opacity-0 w-full h-full hover:cursor-not-allowed"></div>}

            {optionChars.map((char, ind) => {
              const currOptionIsCorrectAns =
                fields[colNames[5]] === colNames[ind + 1];
              const isCurrOptClicked = selectedOption === colNames[ind + 1];

              return (
                <div
                  key={ind}
                  className={mergeClasses("p-2 border border-slate-500 hover:bg-slate-100", selectedOption ? currOptionIsCorrectAns ? questStyles.correctAns : isCurrOptClicked ? questStyles.wrongAns : "" : "")}
                  role="button"
                  onClick={(e) => {
                    setSelectedOption(colNames[ind + 1]);
                    updateArrOfAnsGiven(colNames[ind + 1]);
                  }}
                >
                  <Text size={500}>
                    <strong>{char}.</strong> {fields[colNames[ind + 1]]}
                  </Text>
                </div>
              );
            })}
          </div>
        </div>
      </div>
    </>
  );

  // return (
  //   <>
  //     <div className="number">Q - {fields.id}</div>
  //     <div className="multi-choice-question">
  //       <div className="question">{fields[colNames[0]]}</div>
  //       <div className="options relative">
  //         {(selectedOption || userRole === UserMeetingRole.organizer) && (
  //           <div className="absolute opacity-0 w-full h-full hover:cursor-not-allowed"></div>
  //         )}

  //         {optionChars.map((char, ind) => {
  //           const currOptionIsCorrectAns =
  //             fields[colNames[5]] === colNames[ind + 1];
  //           const isCurrOptClicked = selectedOption === colNames[ind + 1];
  //           return (
  //             <div
  //               className={
  //                 selectedOption
  //                   ? currOptionIsCorrectAns
  //                     ? "right-option"
  //                     : isCurrOptClicked
  //                       ? "wrong-option"
  //                       : ""
  //                   : ""
  //               }
  //               key={ind}
  //               role="button"
  //               onClick={(e) => {
  //                 setSelectedOption(colNames[ind + 1]);
  //                 updateArrOfAnsGiven(colNames[ind + 1]);
  //               }}
  //             >
  //               <Text size={400}>
  //                 <strong>{char}.</strong> {fields[colNames[ind + 1]]}
  //               </Text>
  //             </div>
  //           );
  //         })}
  //       </div>
  //     </div>
  //   </>
  // );
};

export default NewQuestionnaireStage;
