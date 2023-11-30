import { createContext } from "react";

export const TeamsFxContext = createContext({
  theme: undefined,
  themeString: "",
  teamsUserCredential: undefined,
  teamsPageType: {},
  setIsQuestionnaireSitePresent: () => {},
  // isQuestionnaireSitePresent: true,
  // setIsQuestionnaireSitePresent: (isQuestionnaireSitePresent)=> {
  //   console.log("i am in context", isQuestionnaireSitePresent)
  //   // const currentVal = this.isQuestionnaireSitePresent;
  //   // this.isQuestionnaireSitePresent = !currentVal;
  // },
});
