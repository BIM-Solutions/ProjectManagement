import * as React from "react";
import { useState, useEffect } from "react";
import {
  makeStyles,
  tokens,
  IdPrefixProvider,
  webLightTheme,
  FluentProvider,
  Title1,
  Spinner,
  TabList,
  Tab,
} from "@fluentui/react-components";
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
import { SPProvider } from "../../common/components/SPContext";

const useStyles = makeStyles({
  root: {
    overflow: "hidden",
    display: "flex",
    flex: 1,
    minHeight: 0,
    backgroundColor: tokens.colorNeutralBackground1,
    color: tokens.colorNeutralForeground1,
    width: "100vw",
    height: "100vh",
  },
  nav: {
    height: "100vh",
    borderRight: `1px solid ${tokens.colorNeutralStroke1}`,
    padding: tokens.spacingVerticalL,
    backgroundColor: tokens.colorNeutralBackground2,
    minWidth: "220px",
    boxSizing: "border-box",
  },
  content: {
    flex: "1 1 0",
    minWidth: "400px",
    padding: tokens.spacingVerticalXXL,
    display: "flex",
    flexDirection: "column",
    gap: tokens.spacingVerticalXL,
    overflowX: "auto",
    height: "100vh",
    boxSizing: "border-box",
  },
  header: {
    marginBottom: tokens.spacingVerticalL,
  },
  tabList: {
    marginBottom: tokens.spacingVerticalL,
    width: "100%",
    maxWidth: "600px",
  },
  error: {
    color: tokens.colorPaletteRedForeground1,
    backgroundColor: tokens.colorPaletteRedBackground1,
    padding: tokens.spacingVerticalM,
    borderRadius: tokens.borderRadiusMedium,
    marginBottom: tokens.spacingVerticalL,
    fontWeight: 500,
    maxWidth: "600px",
    width: "100%",
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
    return <Spinner size="large" label="Loading..." />;
  }

  if (error) {
    return <div className={styles.error}>{error}</div>;
  }

  return (
    <IdPrefixProvider value={`month-picker-${props.context.instanceId}-`}>
      <FluentProvider theme={webLightTheme}>
        <SPProvider context={props.context}>
          <div className={styles.root}>
            {/* Left Panel - Navigation Drawer */}
            <nav className={styles.nav}>
              <Navigation context={props.context} currentPage="2" />
            </nav>

            {/* Center Panel */}
            <div className={styles.content}>
              <div className={styles.header}>
                <Title1>Resource Management</Title1>
              </div>
              <TabList
                selectedValue={selectedView}
                onTabSelect={(_, data) =>
                  setSelectedView(data.value as "tasks" | "calendar")
                }
                className={styles.tabList}
              >
                <Tab value="tasks">Tasks</Tab>
                <Tab value="calendar">Calendar</Tab>
              </TabList>
              {selectedView === "tasks" && (
                <TasksList
                  listName={props.tasksListName}
                  userDisplayName={props.userDisplayName}
                  context={props.context}
                />
              )}
              {selectedView === "calendar" && (
                <CalendarView
                  showTeamCalendar={props.showTeamCalendar}
                  groupId={props.groupId}
                  context={props.context}
                  userDisplayName={props.userDisplayName}
                  listName={props.tasksListName}
                />
              )}
            </div>
          </div>
        </SPProvider>
      </FluentProvider>
    </IdPrefixProvider>
  );
}
