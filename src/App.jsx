// https://fluentsite.z22.web.core.windows.net/quick-start
import {
  FluentProvider,
  teamsLightTheme,
  teamsDarkTheme,
  teamsHighContrastTheme,
  Spinner,
  tokens,
} from "@fluentui/react-components";
import { HashRouter as Router, Route, Routes } from "react-router-dom";
import { useTeamsUserCredential } from "@microsoft/teamsfx-react";
import Privacy from "./components/Privacy";
import TermsOfUse from "./components/TermsOfUse";
import Tab from "./components/Tab";
import { TeamsFxContext } from "./components/Context";
import config from "./lib/config";
import TabConfig from "./components/TabConfig";
import Questionnaire from "./components/sub/Questionnaire";
import Analysis from "./components/sub/Analysis";

/**
 * The main app which handles the initialization and routing
 * of the app.
 */
export default function App() {
  const { loading, theme, themeString, teamsUserCredential } =
    useTeamsUserCredential({
      initiateLoginEndpoint: config.initiateLoginEndpoint,
      clientId: config.clientId,
    });

  // *create site in sharepoint
  // const [isQuestionnaireSitePresent, setIsQuestionnaireSitePresent] =
  //   useState(true);
  // console.log("hello app.jsx", isQuestionnaireSitePresent);
  // const teamsPageType = useRef("");
  // useEffect(() => {
  //   // Initialize teams app
  //   app.initialize().then(async () => {
  //     const authConfig = {
  //       initiateLoginEndpoint: config.initiateLoginEndpoint,
  //       clientId: config.clientId,
  //     };
  //     const credential = new TeamsUserCredential(authConfig);
  //     const token = await credential.getToken([
  //       "https://ygr11.sharepoint.com/Sites.ReadWrite.All",
  //     ]);
  //     // console.log("TOKEN------------->", token.token);

  //     *api to create site in sharepoint
  //     const res = axios.post(
  //       "https://ygr11.sharepoint.com/_api/SPSiteManager/create", // * /create /delete  https://learn.microsoft.com/en-us/sharepoint/dev/apis/site-creation-rest
  //       {
  //         request: {
  //           Title: "New Hidden Teams Site by API", // *name of the site
  //           Description: "Site created b API", // *Desc of the site
  //           WebTemplate: "STS#3", // *teams site -> "WebTemplate":"STS#3" (or) communication site -> "WebTemplate":"SITEPAGEPUBLISHING#0"
  //           Url: "https://ygr11.sharepoint.com/sites/hteams1", // *url for that site
  //         },
  //       },
  //       { headers: { Authorization: `Bearer ${token.token}` } }
  //     );

  //     app.notifySuccess();
  //   });
  // }, []);

  return (
    <TeamsFxContext.Provider
      value={{
        theme,
        themeString,
        teamsUserCredential,
        // setIsQuestionnaireSitePresent,
      }}
    >
      <FluentProvider
        theme={
          themeString === "dark"
            ? teamsDarkTheme
            : themeString === "contrast"
            ? teamsHighContrastTheme
            : {
                ...teamsLightTheme,
                colorNeutralBackground3: "#eeeeee",
              }
        }
        style={
          !window.location.href?.includes("config")
            ? {
                background: tokens.colorNeutralBackground3,
                minHeight: "100vh",
              }
            : {}
        }
      >
        <Router>
          {loading ? (
            <Spinner style={{ margin: 100 }} />
          ) : (
            <Routes>
              <Route path="/privacy" element={<Privacy />} />
              <Route path="/termsofuse" element={<TermsOfUse />} />
              <Route path="/tab" element={<Tab />} />
              <Route path="/config" element={<TabConfig />} />
              <Route path="/analytics" element={<Analysis />} />
              <Route path="/questionnaire" element={<Questionnaire />} />
              {/* <Route path="*" element={<Navigate to={"/tab"} />}></Route> */}
            </Routes>
          )}
        </Router>
      </FluentProvider>
    </TeamsFxContext.Provider>
  );
}
