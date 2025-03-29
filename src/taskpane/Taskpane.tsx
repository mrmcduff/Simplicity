import React, { useState } from "react";
import {
  PrimaryButton,
  Stack,
  Text,
  MessageBar,
  MessageBarType,
} from "@fluentui/react";
import "./taskpane.css";

export const Taskpane: React.FC = () => {
  const [message, setMessage] = useState<string | null>(null);
  const [error, setError] = useState<string | null>(null);

  const runAddIn = async () => {
    try {
      setMessage(null);
      setError(null);

      await Excel.run(async (context) => {
        const currentWorksheet =
          context.workbook.worksheets.getActiveWorksheet();

        // Create a sample table with data
        const dataRange = currentWorksheet.getRange("A1:D5");
        dataRange.values = [
          ["Quarter", "Revenue", "Expenses", "Profit"],
          ["Q1", 350000, 215000, "=B2-C2"],
          ["Q2", 410000, 245000, "=B3-C3"],
          ["Q3", 495000, 308000, "=B4-C4"],
          ["Q4", 570000, 335000, "=B5-C5"],
        ];

        // Format the table
        dataRange.format.autofitColumns();

        // Create a header row
        const headerRange = currentWorksheet.getRange("A1:D1");
        headerRange.format.fill.color = "#4472C4";
        headerRange.format.font.color = "white";
        headerRange.format.font.bold = true;

        // Add a formula for total
        const totalLabelCell = currentWorksheet.getRange("A7");
        totalLabelCell.values = [["Total"]];
        totalLabelCell.format.font.bold = true;

        // Add formula cells for totals
        const revenueTotalCell = currentWorksheet.getRange("B7");
        revenueTotalCell.formulas = [["=SUM(B2:B5)"]];

        const expensesTotalCell = currentWorksheet.getRange("C7");
        expensesTotalCell.formulas = [["=SUM(C2:C5)"]];

        const profitTotalCell = currentWorksheet.getRange("D7");
        profitTotalCell.formulas = [["=SUM(D2:D5)"]];

        // Create a simple chart
        const chartDataRange = currentWorksheet.getRange("A1:B5");
        const chart = currentWorksheet.charts.add(
          "ColumnClustered",
          chartDataRange,
          "Auto"
        );

        chart.title.text = "Quarterly Revenue";
        chart.setPosition("G1", "L15");

        setMessage("Data and chart created successfully!");

        await context.sync();
      });
    } catch (err: any) {
      setError(`Error: ${err.message}`);
      console.error(err);
    }
  };

  return (
    <div className="ms-welcome">
      <header className="ms-welcome__header ms-bgColor-neutralLighter">
        <img
          width="90"
          height="90"
          src="/logo-filled.png"
          alt="Excel Add-in"
          title="Excel Add-in"
        />
        <h1 className="ms-font-su">Welcome</h1>
      </header>
      <main className="ms-welcome__main">
        <h2 className="ms-font-xl">Excel Add-in with Vite</h2>
        <p className="ms-font-l">
          This is a simple Excel add-in built with Vite and React.
        </p>

        <PrimaryButton
          className="ms-welcome__action"
          iconProps={{ iconName: "ChevronRight" }}
          onClick={runAddIn}
        >
          Run
        </PrimaryButton>

        <Stack
          tokens={{ childrenGap: 15 }}
          style={{ width: "100%", marginTop: 20 }}
        >
          {message && (
            <MessageBar messageBarType={MessageBarType.success}>
              {message}
            </MessageBar>
          )}

          {error && (
            <MessageBar messageBarType={MessageBarType.error}>
              {error}
            </MessageBar>
          )}
        </Stack>
      </main>
    </div>
  );
};

export default Taskpane;
