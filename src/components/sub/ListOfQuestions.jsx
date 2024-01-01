import React, { useContext, useEffect, useState } from "react";
import { TeamsFxContext } from "../Context";
import { getListItems } from "../../lib/utils";
import NoBgLoading from "../../assets/noBgLoading.webp";
import { Image, Spinner } from "@fluentui/react-components";
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

  // console.log("selectedQuestionnaireId == ", NoBgLoading);

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
        {/* loading && (
          <div className="absolute">
            {/* <Image alt="Loading..." className="w-28 h-28" src={NoBgLoading} /> */}
        {/* <Spinner size="large" />
          </div>
        )} */}
        <ul>
          {!!data.length &&
            data.map((row) => <li key={row.id}>{row.fields.Title}</li>)}
        </ul>
      </div>
    </>
  );
};

export default ListOfQuestions;
