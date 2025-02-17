/* global Office Excel console HTMLTextAreaElement */

import React, { useState, useEffect, ChangeEvent } from "react";
import {
  makeStyles,
  Divider,
  Button,
  Dialog,
  DialogTrigger,
  DialogSurface,
  DialogTitle,
  DialogContent,
  DialogBody,
  DialogActions,
  Field,
  Textarea,
  TextareaOnChangeData,
} from "@fluentui/react-components";
import { DatePicker } from "@fluentui/react-datepicker-compat";
import { Timeline, type TimelineItem } from "./Timeline";
import { useXLService, type Issue, XLService } from "../services/xlservice";

interface AppProps {
  title: string;
}

interface NewActivity {
  date?: Date;
  activity?: string;
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
  return jsDate; // convert to local time zone
}

function jsDateToExcelDate(jsDate: Date): number {
  const excelEpoch = new Date(Date.UTC(1899, 11, 30));
  const excelDate = (jsDate.getTime() - excelEpoch.getTime()) / (24 * 60 * 60 * 1000);
  return Math.round(excelDate); // round to nearest day excelDate;
}

function formatDate(date: Date): string {
  const day = date.getDate();
  const month = date.toLocaleString("default", { month: "short" });
  const year = date.getFullYear();
  return `${day}-${month}-${year}`;
}
const activityInput: NewActivity = {
  date: undefined,
  activity: undefined,
};

let xlService: XLService;

const App: React.FC<AppProps> = () => {
  const [timelineItems, setTimelineItems] = useState<TimelineItem[]>([]);
  const [selectedIssue, setSelectedIssue] = useState<Issue | null>(null);
  const [isCorrectFile, setIsCorrectFile] = useState(false);
  const [activityCanSave, setActivityCanSave] = useState(false);
  const [newActivityOpen, setNewActivityOpen] = useState(false);

  const styles = useStyles();

  useEffect(() => {
    Office.onReady((info) => {
      if (info.host === Office.HostType.Excel) {
        Excel.run(async (context) => {
          xlService = useXLService();
          setIsCorrectFile(await xlService.isCorrectFile(context));

          // get the IT Tracker sheet
          const trackerSheet = context.workbook.worksheets.getItem("IT Tracker");
          const activitySheet = context.workbook.worksheets.getItem("Activity");
          const activeSheet = context.workbook.worksheets.getActiveWorksheet();
          activeSheet.load("name");

          // Add the event handler
          trackerSheet.onSelectionChanged.add(handleSelectionChangeOnTrackerSheet);
          trackerSheet.onActivated.add(handleTrackerSheetActivated);
          trackerSheet.onDeactivated.add(handleTrackerSheetDeactivated);

          activitySheet.onChanged.add(handleActivitySheetChanged);
          await context.sync();
          if (activeSheet.name === "IT Tracker") {
            handleSelectionChangeOnTrackerSheet();
          }
        }).catch((error) => {
          console.log(error);
        });
      }
    });
  }, []);

  function onActivityDateChange(date: Date) {
    activityInput.date = date;
    setActivityCanSave(!!activityInput.date && !!activityInput.activity);
  }

  function onActivityDescriptionChange(_: ChangeEvent<HTMLTextAreaElement>, data: TextareaOnChangeData) {
    activityInput.activity = data.value;
    setActivityCanSave(!!activityInput.date && !!activityInput.activity);
  }

  async function onActivitySave() {
    console.log("Saving activity", activityInput);
    try {
      await xlService.addActivity({
        issueId: selectedIssue!.id,
        date: jsDateToExcelDate(activityInput.date!),
        description: activityInput.activity!,
      });
      await handleActivitySheetChanged();
      await handleSelectionChangeOnTrackerSheet();
    } catch (error) {
      console.log(error);
      return;
    }

    setNewActivityOpen(false);
  }

  function onAddActivity() {
    console.log("Adding activity", activityInput);
    activityInput.activity = undefined;
    activityInput.date = undefined;
    setActivityCanSave(false);
  }

  async function handleTrackerSheetActivated(event: Excel.WorksheetActivatedEventArgs) {
    console.debug("Tracker sheet activated", event.worksheetId);
    return handleSelectionChangeOnTrackerSheet();
  }

  async function handleTrackerSheetDeactivated(event: Excel.WorksheetDeactivatedEventArgs) {
    console.debug("Tracker sheet deactivated", event.worksheetId);
    setSelectedIssue(null);
    setTimelineItems([]);
  }

  async function handleSelectionChangeOnTrackerSheet() {
    setNewActivityOpen(false);
    try {
      const issue = await xlService.getSelectedIssue();
      if (!issue) {
        setSelectedIssue(null);
        setTimelineItems([]);
        return;
      }

      setSelectedIssue(issue);

      const activities = await xlService.getActivityByIssueID(issue.id);
      const timelines = activities.map((activity) => {
        const date = excelDateToJSDate(activity.date);
        const formattedDate = formatDate(date);
        return {
          date: formattedDate,
          description: activity.description,
        };
      });

      setTimelineItems(timelines);
    } catch (error) {
      console.log(error);
    }
  }

  async function handleActivitySheetChanged() {
    Excel.run(async (context) => {
      await xlService.refreshActivityCache(context);
    });
  }

  return (
    <div className={styles.root}>
      {!isCorrectFile && <p>This add-in can only be used with the IT-Commercial Interlock file</p>}
      {isCorrectFile && selectedIssue && (
        <div>
          <h2>{selectedIssue.title}</h2>
          <p>{selectedIssue.description}</p>
          <Divider>activity</Divider>
          <Timeline items={timelineItems} />
          <Dialog open={newActivityOpen} onOpenChange={(_, data) => setNewActivityOpen(data.open)}>
            <DialogTrigger disableButtonEnhancement>
              <Button appearance="primary" style={{ marginTop: "20px" }} onClick={onAddActivity}>
                Add Activity
              </Button>
            </DialogTrigger>
            <DialogSurface aria-describedby={undefined}>
              <DialogBody>
                <DialogTitle>Add activity for {selectedIssue.title}</DialogTitle>
                <DialogContent>
                  <Field label="Date Start" required style={{ marginTop: "10px" }}>
                    <DatePicker placeholder="Select a date" onSelectDate={onActivityDateChange} />
                  </Field>
                  <Field label="Activity" required style={{ marginTop: "10px" }}>
                    <Textarea placeholder="Enter your activity" onChange={onActivityDescriptionChange} />
                  </Field>
                </DialogContent>
                <DialogActions>
                  <Button disabled={!activityCanSave} appearance="primary" onClick={onActivitySave}>
                    Save
                  </Button>
                  <DialogTrigger disableButtonEnhancement>
                    <Button appearance="secondary">Cancel</Button>
                  </DialogTrigger>
                </DialogActions>
              </DialogBody>
            </DialogSurface>
          </Dialog>
        </div>
      )}
    </div>
  );
};

export default App;
