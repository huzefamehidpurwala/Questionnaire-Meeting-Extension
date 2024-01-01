import React, { useContext, useEffect, useState } from "react";
import { TeamsFxContext } from "../Context";
import { getListItems } from "../../lib/utils";
import SmallPopUp from "../SmallPopUp";

const ListOfQuestions = ({ selectedQuestionnaireId }) => {
  const teamsUserCredential = useContext(TeamsFxContext).teamsUserCredential;

  const [loading, setLoading] = useState(false);
  const [data, setData] = useState([]);

  useEffect(() => {
    (async () => {
      setLoading(true);
      try {
        const questionnaireRootList = await getListItems(
          teamsUserCredential,
          selectedQuestionnaireId
        );

        const tempData = questionnaireRootList.graphClientMessage.value;
        setData(tempData);
      } catch (error) {
        console.error("error ", error);
      }
      setLoading(false);
    })();
    // eslint-disable-next-line
  }, [selectedQuestionnaireId]);

  return (
    <>
      {/* loading popup */}
      <SmallPopUp
        className="loading"
        msg={"Fetching questions..."}
        open={loading}
        spinner={true}
        activeActions={false}
        modalType="alert"
      />

      <div className="relative min-w-max min-h-max">
        <ul>
          {!!data.length &&
            data.map((row) => <li key={row.id}>{row.fields.Title}</li>)}
        </ul>
      </div>
    </>
  );
};

export default ListOfQuestions;
