/* global Excel console */

export interface Issue {
  id: number;
  title: string;
  description: string;
  status?: string;
  assignedTo?: string;
  dueDate?: string;
}

export interface Activity {
  issueId: number;
  date: number;
  description: string;
}

const activityColumnIndexes = {
  issueId: 0,
  date: 0,
  description: 0,
};

const trackerColumnIndexes = {
  id: 0,
  title: 0,
  description: 0,
};

const trackerSheetName = "IT Tracker";
const activitySheetName = "Activity";

export class XLService {
  private activityCache: Map<number, Activity[]>;

  private activityHeaderRow: number = 3;
  private trackerHeaderRow: number = 3;

  constructor() {
    Excel.run(async (context) => {
      await this.getActivityHeaderIndexes(context);
      await this.getTrackerHeaderIndexes(context);
      await this.refreshActivityCache(context);
    }).catch((error) => {
      console.log(error);
    });
  }

  async refreshActivityCache(context: Excel.RequestContext) {
    // get all rows from activity sheet
    // for each row, get the issue id and add the activity to the cache
    this.activityCache = new Map<number, Activity[]>();
    const activitySheet = context.workbook.worksheets.getItem(activitySheetName);
    const usedRange = activitySheet.getUsedRange();
    usedRange.load("values");
    await context.sync();

    const rows = usedRange.values;

    for (let i = 4; i < rows.length; i++) {
      const row = rows[i];
      const issueId = row[activityColumnIndexes.issueId];
      if (!issueId) {
        continue;
      }

      const activity = {
        issueId: issueId,
        date: row[activityColumnIndexes.date],
        description: row[activityColumnIndexes.description],
      };

      if (!this.activityCache.has(issueId)) {
        this.activityCache.set(issueId, []);
      }

      this.activityCache.get(issueId)?.push(activity);
    }
  }

  async getActivityHeaderIndexes(context: Excel.RequestContext) {
    const activitySheet = context.workbook.worksheets.getItem(activitySheetName);
    const usedRange = activitySheet.getUsedRange();
    usedRange.load("values");
    await context.sync();

    const headerRow = usedRange.values[this.activityHeaderRow];
    for (let i = 0; i < headerRow.length; i++) {
      switch (headerRow[i].toString().toLowerCase()) {
        case "tracker id":
          activityColumnIndexes.issueId = i;
          break;
        case "date start":
          activityColumnIndexes.date = i;
          break;
        case "activity":
          activityColumnIndexes.description = i;
          break;
      }
    }
  }

  async getTrackerHeaderIndexes(context: Excel.RequestContext) {
    const trackerSheet = context.workbook.worksheets.getItem(trackerSheetName);
    const usedRange = trackerSheet.getUsedRange();
    usedRange.load("values");
    await context.sync();

    const headerRow = usedRange.values[this.trackerHeaderRow];
    for (let i = 0; i < headerRow.length; i++) {
      switch (headerRow[i].toString().toLowerCase()) {
        case "id":
          trackerColumnIndexes.id = i;
          break;
        case "issue":
          trackerColumnIndexes.title = i;
          break;
        case "description":
          trackerColumnIndexes.description = i;
          break;
      }
    }
  }

  async getSelectedIssue(): Promise<Issue | null> {
    return Excel.run(async (context) => {
      const activeSheet = context.workbook.worksheets.getActiveWorksheet();
      activeSheet.load("name");
      await context.sync();

      if (activeSheet.name !== trackerSheetName) {
        return null;
      }

      const selectedRange = context.workbook.getSelectedRange();
      selectedRange.load(["rowIndex", "values"]);
      await context.sync();

      const issueRow = activeSheet.getRangeByIndexes(selectedRange.rowIndex, 1, selectedRange.rowIndex, 10);
      issueRow.load("values");
      await context.sync();

      const issueId = issueRow.values[0][0];
      if (!issueId) {
        return null;
      }

      if (isNaN(issueId)) {
        return null;
      }

      if ((issueId * 10) % 10 === 0) {
        return null;
      }

      const issueTitle = issueRow.values[0][1];
      const issueDescription = issueRow.values[0][2];
      return {
        id: issueId,
        title: issueTitle,
        description: issueDescription,
      };
    }).catch((error) => {
      console.log(error);
      return null;
    });
  }

  async getActivityByIssueID(issueId: number): Promise<Activity[]> {
    return this.activityCache.get(issueId) || [];
  }

  async isCorrectFile(context: Excel.RequestContext): Promise<boolean> {
    const reportProperty = context.workbook.properties.custom.getItem("Report");
    reportProperty.load("value");
    await context.sync();
    return reportProperty.value === "itcomm-tracker";
  }

  async addActivity(activity: Activity) {
    Excel.run(async (context) => {
      // get last row in activity sheet
      // add activity to last row
      // refresh activity cache
      console.log("Adding activity", activity);
      const activitySheet = context.workbook.worksheets.getItem(activitySheetName);
      const usedRange = activitySheet.getUsedRange();
      usedRange.load(["rowCount", "address"]);
      await context.sync();

      const lastRow = usedRange.rowCount;
      const activityRange = activitySheet.getRangeByIndexes(lastRow, 1, 1, 7);
      activityRange.values = [[activity.date, null, null, activity.issueId, activity.description, null, null]];
      await context.sync();
      await this.refreshActivityCache(context);
      activityRange.getCell(0, 0).numberFormat = [["dd-mmm-yyyy"]];
      activityRange.format.borders
        .getItem("EdgeBottom")
        .set({ style: "Continuous", color: "#275317", weight: "Medium" });

      activityRange.format.borders.getItem("EdgeLeft").set({ style: "Continuous", color: "#275317", weight: "Medium" });
      activityRange.format.borders
        .getItem("EdgeRight")
        .set({ style: "Continuous", color: "#275317", weight: "Medium" });
      activityRange.format.borders.getItem("EdgeTop").set({ style: "Continuous", color: "#e0e7db", weight: "Thin" });
      activityRange.format.borders
        .getItem("InsideVertical")
        .set({ style: "Continuous", color: "#275317", weight: "Thin" });
      await context.sync();
    });
  }
}

let xlService: XLService;

export function useXLService() {
  if (!xlService) {
    xlService = new XLService();
  }
  return xlService;
}
