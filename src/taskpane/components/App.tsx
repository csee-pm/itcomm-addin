/* global Office Excel console */

import React, { useState, useEffect } from "react";
import { makeStyles, Divider } from "@fluentui/react-components";
import { Timeline, type TimelineItem } from "./Timeline";

interface AppProps {
  title: string;
}

interface Issue {
  id: string;
  title: string;
  description: string;
  status?: string;
  assignedTo?: string;
  dueDate?: string;
}

const useStyles = makeStyles({
  root: {
    minHeight: "100vh",
    padding: "10px",
  },
});

function excelDateToJSDate(excelDate: number): Date {
  const excelEpoch = new Date(Date.UTC(1899, 11, 30));
  const jsDate = new Date(excelEpoch.getTime() + excelDate * 24 * 60 * 60 * 1000);
  return jsDate;
}

function formatDate(date: Date): string {
  const day = date.getDate();
  const month = date.toLocaleString("default", { month: "short" });
  const year = date.getFullYear();
  return `${day}-${month}-${year}`;
}

const App: React.FC<AppProps> = () => {
  const [timelineItems, setTimelineItems] = useState<TimelineItem[]>([]);
  const [selectedIssue, setSelectedIssue] = useState<Issue | null>(null);

  const styles = useStyles();

  useEffect(() => {
    Office.onReady((info) => {
      if (info.host === Office.HostType.Excel) {
        Excel.run(async (context) => {
          // get the IT Tracker sheet
          const trackerSheet = context.workbook.worksheets.getItem("IT Tracker");

          // Add the event handler
          trackerSheet.onSelectionChanged.add(handleSelectionChange);
          await context.sync();
        }).catch((error) => {
          console.log(error);
        });
      }
    });
  }, []);

  async function handleSelectionChange(event: Excel.WorksheetSelectionChangedEventArgs) {
    console.log("Selection changed", event.address);
    try {
      await Excel.run(async (context) => {
        const selectedRange = context.workbook.getSelectedRange();
        selectedRange.load(["rowIndex", "values"]);
        await context.sync();

        const trackerSheet = context.workbook.worksheets.getItem("IT Tracker");

        // get value from cell column B and row range.rowIndex
        const issueRow = trackerSheet.getRangeByIndexes(selectedRange.rowIndex, 1, selectedRange.rowIndex, 10);
        issueRow.load("values");
        await context.sync();

        const issueId = issueRow.values[0][0];
        const issueTitle = issueRow.values[0][1];
        const issueDescription = issueRow.values[0][2];
        setSelectedIssue({
          id: issueId,
          title: issueTitle,
          description: issueDescription,
        });

        const activitySheet = context.workbook.worksheets.getItem("Activity");
        const usedRange = activitySheet.getUsedRange();
        usedRange.load("values");
        await context.sync();

        const allRows = usedRange.values;
        const matching: TimelineItem[] = [];

        for (let i = 4; i < allRows.length; i++) {
          const row = allRows[i];
          // check if column E is equal to issueId
          if (row[3] === issueId) {
            matching.push({
              date: formatDate(excelDateToJSDate(row[0])),
              description: row[4],
            });
          }
        }

        console.log("matching", matching);

        setTimelineItems(matching);
      });
    } catch (error) {
      console.log(error);
    }
  }

  return (
    <div className={styles.root}>
      {selectedIssue && (
        <div>
          <h2>{selectedIssue.title}</h2>
          <p>{selectedIssue.description}</p>
          <Divider>activity</Divider>
        </div>
      )}
      <Timeline items={timelineItems} />
    </div>
  );
};

export default App;
