// ProgrammeView.tsx
// This component splits the programme view across center and right containers of LandingPage.

import * as React from "react";
import { useEffect, useState, useContext } from "react";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { Project } from "../../services/ProjectSelectionServices";
import { SPContext } from "../../../common/components/SPContext";
import { TaskItem } from "./ProgrammeTab";
import ProgrammeTaskDetails from "./ProgrammeTaskDetails";
import { makeStyles, tokens } from "@fluentui/react-components";

const useStyles = makeStyles({
  container: {
    display: "flex",
    flexDirection: "row",
    width: "100%",
    height: "100%",
  },
  middlePanel: {
    flex: 2,
    height: "100%",
    paddingRight: tokens.spacingHorizontalL,
    boxSizing: "border-box",
  },
  rightPanel: {
    flex: 1,
    height: "100%",
    overflowY: "auto",
    boxSizing: "border-box",
    paddingLeft: tokens.spacingHorizontalL,
    borderLeft: `1px solid ${tokens.colorNeutralStroke1}`,
  },
});

export interface ProgrammeProps {
  context: WebPartContext;
  project: Project;
}

const ProgrammeView: React.FC<ProgrammeProps> = ({ context, project }) => {
  const styles = useStyles();
  const sp = useContext(SPContext);
  const [tasks, setTasks] = useState<TaskItem[]>([]);
  const [selectedTask, setSelectedTask] = useState<TaskItem | undefined>(
    undefined
  );

  const fetchTasks = async (): Promise<void> => {
    const results = await sp.web.lists
      .getByTitle("9719_ProjectTasks")
      .items.filter<TaskItem>(`Title eq '${project.ProjectNumber}'`)
      .top(5000)();
    setTasks(results);
  };

  useEffect(() => {
    fetchTasks().catch(console.error);
  }, [project]);

  return (
    <div className={styles.container}>
      <ProgrammeTaskDetails
        tasks={tasks}
        selectedTask={selectedTask}
        setSelectedTask={setSelectedTask}
        context={context}
        project={project}
        reloadTasks={fetchTasks}
      />
    </div>
  );
};

export default ProgrammeView;
