// https://fluentsite.z22.web.core.windows.net/quick-start
import {
  FluentProvider,
  teamsLightTheme,
  teamsDarkTheme,
  teamsHighContrastTheme,
  Spinner,
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
import ListQuestionnaire from "./components/sub/ListQuestionnaire";
import CreateQuestionnaire from "./components/sub/CreateQuestionnaire";

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

    console.log("this is to check")

  return (
    <TeamsFxContext.Provider
      value={{ theme, themeString, teamsUserCredential }}
    >
      <FluentProvider
        theme={
          themeString === "dark"
            ? teamsDarkTheme
            : themeString === "contrast"
            ? teamsHighContrastTheme
            : { ...teamsLightTheme, colorNeutralBackground3: "#eeeeee" }
        }
      >
        <main
          className={
            !window.location.href?.includes("config")
              ? "bg-teams-bg-3 h-screen"
              : ""
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
                <Route
                  path="/listQuestionnaire"
                  element={<ListQuestionnaire />}
                />
                <Route
                  path="/createQuestionnaire"
                  element={<CreateQuestionnaire />}
                />
                <Route path="/questionnaire" element={<Questionnaire />} />
              </Routes>
            )}
          </Router>
        </main>
      </FluentProvider>
    </TeamsFxContext.Provider>
  );
}
