import * as React from "react";
import { makeStyles, shorthands, Text } from "@fluentui/react-components";
import { ArrowCircleRight20Filled } from "@fluentui/react-icons";

const useStyles = makeStyles({
  timelineContainer: {
    display: "flex",
    flexDirection: "column",
    ...shorthands.gap("10px"),
  },
  timelineItem: {
    display: "flex",
    flexDirection: "row",
    // alignItems: "flex-start",
  },
  iconColumn: {
    // we want the icon + line in one vertical stack
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
    marginRight: "8px",
  },
  circleIcon: {
    marginBottom: "4px", // space between icon and line
    color: "#0078d4", // blue color for the icon
  },
  verticalLine: {
    flex: 1,
    width: "2px",
    backgroundColor: "#ccc",
  },
  contentColumn: {
    display: "flex",
    flexDirection: "column",
  },
  dateText: {
    fontWeight: "bold",
  },
});

export interface TimelineItem {
  date: string;
  description: string;
}

interface TimelineProps {
  items: TimelineItem[];
}

export const Timeline: React.FC<TimelineProps> = ({ items }) => {
  const styles = useStyles();

  return (
    <div className={styles.timelineContainer}>
      {items.map((item, index) => {
        const isLast = index === items.length - 1;
        return (
          <div key={index} className={styles.timelineItem}>
            <div className={styles.iconColumn}>
              <ArrowCircleRight20Filled className={styles.circleIcon} />
              {!isLast && <div className={styles.verticalLine} />}
            </div>
            <div className={styles.contentColumn}>
              <Text className={styles.dateText}>{item.date}</Text>
              <Text>{item.description}</Text>
            </div>
          </div>
        );
      })}
    </div>
  );
};
