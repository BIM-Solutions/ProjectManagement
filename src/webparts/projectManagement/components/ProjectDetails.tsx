import * as React from "react";
import { useEffect, useState } from "react";
import { Text, makeStyles } from "@fluentui/react-components";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import {
  Project,
  ProjectSelectionService,
} from "../services/ProjectSelectionServices";
import ProjectForm from "./common/ProjectForm";
import ProjectTabs from "./projectInformation/ProjectTabs";
import { DEBUG } from "./common/DevVariables";
import { eventService } from "../services/EventService";
import { TaskItem } from "./projectCalender/ProgrammeTab";

interface ProjectDetailsProps {
  context: WebPartContext;
  onSave?: () => void;
  onEdit?: () => void;
  onDelete?: () => void;
  onTabChange?: (tab: string) => void;
  tasks: TaskItem[];
  setTasks: React.Dispatch<React.SetStateAction<TaskItem[]>>;
  selectedTask: TaskItem | undefined;
  setSelectedTask: (task: TaskItem | undefined) => void;
  selectedStageId: number | undefined;
  stagesChanged?: () => void;
}

const useStyles = makeStyles({
  root: {
    gap: "20",
    padding: "20",
    marginTop: "20",
    height: "auto",
    minHeight: "50vh",
    overflowY: "auto",
    backgroundColor: "#f1f0ef",
  },

  projectForm: {
    display: "flex",
    flexDirection: "column",
    gap: "20",
    padding: "50",
    margin: "50",
    height: "auto",
    minHeight: "50vh",
    overflowY: "auto",
    backgroundColor: "#f1f0ef",
  },
});

const ProjectDetails: React.FC<ProjectDetailsProps> = ({
  context,
  onEdit,
  onDelete,
  onTabChange,
  tasks,
  setTasks,
  selectedTask,
  setSelectedTask,
  selectedStageId,
  stagesChanged,
}) => {
  const [project, setProject] = useState<Project | undefined>();
  const [isEditing, setIsEditing] = useState(false);
  const styles = useStyles();
  useEffect(() => {
    const service = ProjectSelectionService;
    const listener = (selected: Project | undefined): void => {
      if (!selected) return setProject(undefined);
      if (!DEBUG)
        console.log("ProjectDetails listener - selected project:", selected);
      setProject({ ...selected });
    };
    service.subscribe(listener);
    listener(service.getSelectedProject());
    const unsubscribe = eventService.subscribeToProjectUpdates(() => {
      const latest = service.getSelectedProject();
      if (!DEBUG) console.log("ProjectDetails - updated project:", latest);
      setProject(latest ? { ...latest } : undefined);
    });
    return () => {
      service.unsubscribe(listener);
      unsubscribe();
    };
  }, []);

  if (!project) return <Text>Select a project to view details.</Text>;
  if (!DEBUG) console.log("ProjectDetails - project:", project);

  return (
    <div className={styles.root}>
      {!isEditing ? (
        <ProjectTabs
          context={context}
          project={project}
          onEdit={() => setIsEditing(true)}
          onDelete={onDelete || (() => {})}
          onTabChange={onTabChange}
          tasks={tasks}
          setTasks={setTasks}
          selectedTask={selectedTask}
          setSelectedTask={setSelectedTask}
          selectedStageId={selectedStageId}
          stagesChanged={stagesChanged}
        />
      ) : (
        <div className={styles.projectForm}>
          <ProjectForm
            context={context}
            mode="edit"
            project={project}
            onSuccess={() => {
              setIsEditing(false);
              if (onEdit) {
                onEdit();
              }
            }}
            onCancel={() => setIsEditing(false)}
          />
        </div>
      )}
    </div>
  );
};

export default ProjectDetails;
