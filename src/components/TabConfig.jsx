import React, { useEffect, useState } from "react";
import { app, pages } from "@microsoft/teams-js";
import { Checkbox, Link } from "@fluentui/react-components";

/**
 * The 'Config' component is used to display your group tabs
 * user configuration options.  Here you will allow the user to
 * make their choices and once they are done you will need to validate
 * their choices and communicate that to Teams to enable the save button.
 */
const TabConfig = () => {
  const [checked, setChecked] = useState(false);
  useEffect(() => {
    // Initialize the Microsoft Teams SDK
    app.initialize().then(() => {
      /**
       * When the user clicks "Save", save the url for your configured tab.
       * This allows for the addition of query string parameters based on
       * the settings selected by the user.
       */
      pages.config.registerOnSaveHandler((saveEvent) => {
        const baseUrl = `https://${window.location.hostname}:${window.location.port}`;
        // console.log("in TabConfig", baseUrl);
        pages.config
          .setConfig({
            entityId: "test",
            contentUrl: baseUrl + "/index.html#/tab",
            websiteUrl: baseUrl + "/index.html#/tab",
            suggestedDisplayName: "Meeting Extensions",
          })
          .then(() => {
            saveEvent.notifySuccess();
          });
        saveEvent.notifySuccess();
      });

      // Hide the loading indicator.
      app.notifySuccess();
    });
  }, []);

  useEffect(() => {
    app.initialize().then(() => {
      /**
       * After verifying that the settings for your tab are correctly
       * filled in by the user you need to set the state of the dialog
       * to be valid.  This will enable the save button in the configuration
       * dialog.
       */
      pages.config.setValidityState(checked);

      // Hide the loading indicator.
      app.notifySuccess();
    });
  }, [checked]);

  return (
    <div>
      <h1>Tab Configuration</h1>
      <div>
        This is where you will add your tab configuration options the user can
        choose when the tab is added to your team/group chat.
      </div>
      <Checkbox
        label={
          <>
            Accept the{" "}
            <Link as="a" href="#/termsofuse">
              terms and conditions
            </Link>
          </>
        }
        required
        checked={checked}
        onChange={(e, data) => setChecked(data.checked)}
      />
    </div>
  );
};

export default TabConfig;
