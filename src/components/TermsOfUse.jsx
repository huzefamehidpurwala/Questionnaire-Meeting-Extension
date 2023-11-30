import { Button } from "@fluentui/react-components";
import { ArrowLeft24Filled } from "@fluentui/react-icons";
import React from "react";
import { useNavigate } from "react-router-dom";
/**
 * This component is used to display the required
 * terms of use statement which can be found in a
 * link in the about tab.
 */
const TermsOfUse = () => {
  const navigate = useNavigate();

  return (
    <div>
      <Button onClick={(e) => navigate(-1)} icon={<ArrowLeft24Filled />} />
      <h1>Terms of Use</h1>
    </div>
  );
};

export default TermsOfUse;
