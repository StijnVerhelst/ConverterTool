/**
 * Task Pane Entry Point
 *
 * Initializes Office.js and renders the React application.
 */

import { createRoot } from "react-dom/client";
import { FluentProvider, webLightTheme } from "@fluentui/react-components";
import { App } from "./App";

// Ensure Office.js is ready before rendering
Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    const container = document.getElementById("app");
    if (container) {
      const root = createRoot(container);
      root.render(
        <FluentProvider theme={webLightTheme}>
          <App />
        </FluentProvider>
      );
    }
  } else {
    // Not running in Excel
    const container = document.getElementById("app");
    if (container) {
      container.innerHTML = `
        <div class="sideload-message">
          <h2>Excel Required</h2>
          <p>This add-in only works in Microsoft Excel.</p>
          <p>
            Please <a href="https://learn.microsoft.com/office/dev/add-ins/testing/test-debug-office-add-ins#sideload-an-office-add-in-for-testing" target="_blank">sideload</a>
            this add-in in Excel to use it.
          </p>
        </div>
      `;
    }
  }
});
