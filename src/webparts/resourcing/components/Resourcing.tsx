import * as React from "react";
import { useState, useEffect } from "react";
import {
  Pivot,
  PivotItem,
  Stack,
  Text,
  Spinner,
  SpinnerSize,
  MessageBar,
  MessageBarType,
} from "@fluentui/react";
// import { Calendar, DateRangeType } from "@fluentui/react-calendar-compat";
import { TasksList } from "./TasksList";
import { CalendarView } from "./CalendarView";
import { useGraph } from "../hooks/useGraph";
// import { usePnP } from "../hooks/usePnP";
// import styles from "./Resourcing.module.scss";
import type { IResourcingProps } from "./IResourcingProps";
// import { escape } from "@microsoft/sp-lodash-subset";
// import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SPFx } from "@pnp/sp/presets/all";
import { spfi, SPFI } from "@pnp/sp";
import Navigation from "../../common/components/Navigation";
import { makeStyles, tokens } from "@fluentui/react-components";

const useStyles = makeStyles({
  root: {
    overflow: "hidden",
    display: "flex",
    flex: 1,
    minHeight: 0,
    backgroundColor: tokens.colorNeutralBackground1,
    color: tokens.colorNeutralForeground1,
  },
  nav: {
    height: "100vh",
    borderRight: `1px solid ${tokens.colorNeutralStroke1}`,
    padding: "12px",
    backgroundColor: tokens.colorNeutralBackground2,
  },
  content: {
    flex: "1 1 0",
    minWidth: "400px",
    padding: "24px",
    display: "flex",
    flexDirection: "column",
    gap: "16px",
    overflowX: "auto",
    height: "100vh",
  },
});

export default function Resourcing(
  props: IResourcingProps
): React.ReactElement<IResourcingProps> {
  const [selectedView, setSelectedView] = useState<"tasks" | "calendar">(
    props.defaultView || "tasks"
  );
  const [isLoading, setIsLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);
  const styles = useStyles();

  const { graphClient } = useGraph();
  const sp: SPFI = spfi().using(SPFx(props.context));

  useEffect(() => {
    const initializeData = async (): Promise<void> => {
      try {
        setIsLoading(true);
        await Promise.all([
          sp.web.lists.getByTitle(props.tasksListName).items.top(1000)(),
          graphClient?.api("/me/calendar/events").get(),
        ]);
      } catch (err) {
        setError(err.message);
      } finally {
        setIsLoading(false);
      }
    };

    initializeData().catch(console.error);
  }, []);

  if (isLoading) {
    return <Spinner size={SpinnerSize.large} label="Loading..." />;
  }

  if (error) {
    return (
      <MessageBar messageBarType={MessageBarType.error}>{error}</MessageBar>
    );
  }

  return (
    <div className={styles.root}>
      <div
        style={{
          display: "flex",
          justifyContent: "top",
          alignItems: "top",
          flexWrap: "wrap",
          flexDirection: "row",
          width: "100%",
          height: "100vh",
          boxSizing: "border-box",
          padding: "20px",
          overflowY: "auto",
          overflowX: "auto",
        }}
      >
        {/* Left Panel - Navigation Drawer */}
        <nav className={styles.nav}>
          <Navigation context={props.context} />
        </nav>

        {/* Center Panel */}
        <div className={styles.content}>
          <Stack tokens={{ childrenGap: 15 }}>
            <Text variant="xLarge">Resource Management</Text>

            <Pivot
              selectedKey={selectedView}
              onLinkClick={(item) =>
                setSelectedView(item?.props.itemKey as "tasks" | "calendar")
              }
            >
              <PivotItem headerText="Tasks" itemKey="tasks">
                <TasksList
                  listName={props.tasksListName}
                  userDisplayName={props.userDisplayName}
                  context={props.context}
                />
              </PivotItem>

              <PivotItem headerText="Calendar" itemKey="calendar">
                <CalendarView
                  showTeamCalendar={props.showTeamCalendar}
                  groupId={props.groupId}
                  context={props.context}
                />
              </PivotItem>
            </Pivot>
          </Stack>
        </div>
      </div>
    </div>
  );
}
