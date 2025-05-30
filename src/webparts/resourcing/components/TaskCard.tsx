import * as React from "react";
import {
  Card,
  CardHeader,
  CardFooter,
  // CardPreview,
  makeStyles,
  shorthands,
} from "@fluentui/react-components";
import { ITask } from "../services/TasksService";

export interface TaskCardProps {
  task: ITask;
  onClick: () => void;
}

const useStyles = makeStyles({
  root: {
    marginBottom: "16px",
    borderRadius: "12px",
    boxShadow: "0 2px 8px rgba(0,0,0,0.08)",
    cursor: "pointer",
    ...shorthands.padding("16px"),
    background: "#fff",
    transition: "box-shadow 0.2s",
    ":hover": {
      boxShadow: "0 4px 16px rgba(0,0,0,0.12)",
    },
  },
});

export const TaskCard: React.FC<TaskCardProps> = ({ task, onClick }) => {
  const styles = useStyles();
  return (
    <Card className={styles.root} onClick={onClick}>
      <CardHeader
        header={<span>{task.Title}</span>}
        description={
          <span>Due: {new Date(task.DueDate).toLocaleDateString()}</span>
        }
      />
      <CardFooter>
        <span>Priority: {task.Priority}</span>
        <span style={{ marginLeft: 16 }}>Project: {task.Project}</span>
      </CardFooter>
    </Card>
  );
};

export * from "./TaskCard";
