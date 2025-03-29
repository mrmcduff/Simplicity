import React from "react";
import ReactDOM from "react-dom/client";
import { initializeIcons } from "@fluentui/react";
import Taskpane from "./Taskpane";

// Initialize Fluent UI icons
initializeIcons();

// Office initialization
Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    // Render the component after Office is ready
    const root = ReactDOM.createRoot(document.getElementById("container")!);
    root.render(
      <React.StrictMode>
        <Taskpane />
      </React.StrictMode>
    );
  }
});
