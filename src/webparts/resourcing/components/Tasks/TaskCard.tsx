import * as React from "react";
import {
  Card,
  CardHeader,
  CardFooter,
  // CardPreview,
  Text,
  Tag,
  Persona,
  tokens,
  makeStyles,
  shorthands,
  Badge,
} from "@fluentui/react-components";
import { ITask } from "../../services/TasksService";
// import {
//   CalendarMonthRegular,
//   // CheckmarkCircleRegular,
//   WarningRegular,
// } from "@fluentui/react-icons";

export interface TaskCardProps {
  task: ITask;
  onClick: () => void;
}

const useStyles = makeStyles({
  root: {
    marginBottom: tokens.spacingVerticalMNudge,
    borderRadius: tokens.borderRadiusXLarge,
    boxShadow: tokens.shadow16,
    cursor: "pointer",
    ...shorthands.padding(tokens.spacingHorizontalM),
    background: tokens.colorNeutralBackground1,
    transition: "box-shadow 0.2s, border 0.2s",
    ":hover": {
      boxShadow: tokens.shadow28,
    },
    display: "flex",
    flexDirection: "column",
    gap: tokens.spacingVerticalMNudge,
    border: `1.5px solid transparent`,
  },
  overdue: {
    border: `2px solid ${tokens.colorPaletteRedBorder2}`,
  },
  pillRow: {
    display: "flex",
    gap: tokens.spacingHorizontalS,
    marginTop: tokens.spacingVerticalMNudge,
    flexWrap: "wrap",
  },
  desc: {
    margin: `${tokens.spacingVerticalMNudge} 0`,
    wordBreak: "break-word",
    color: tokens.colorNeutralForeground2,
  },
});

function getStatusTag(status: string): JSX.Element {
  switch (status) {
    case "Not Started":
      return (
        <Badge appearance="filled" color="informative">
          Not Started
        </Badge>
      );
    case "In Progress":
      return (
        <Badge appearance="filled" color="brand">
          In Progress
        </Badge>
      );
    case "Completed":
      return (
        <Badge appearance="filled" color="success">
          Completed
        </Badge>
      );
    default:
      return (
        <Tag appearance="filled" shape="rounded">
          {status}
        </Tag>
      );
  }
}

// function getPriorityTag(priority: string): JSX.Element {
//   switch (priority) {
//     case "High":
//       return (
//         <Tag appearance="filled" shape="rounded" color="red">
//           High
//         </Tag>
//       );
//     case "Medium":
//       return (
//         <Tag appearance="filled" shape="rounded" color="yellow">
//           Medium
//         </Tag>
//       );
//     case "Low":
//       return (
//         <Tag appearance="filled" shape="rounded" color="grey">
//           Low
//         </Tag>
//       );
//     default:
//       return (
//         <Tag appearance="filled" shape="rounded">
//           {priority}
//         </Tag>
//       );
//   }
// }

export const TaskCard: React.FC<TaskCardProps & { columnKey?: string }> = ({
  task,
  onClick,
}) => {
  const styles = useStyles();
  const dueDate = new Date(task.DueDate);
  const isOverdue = task.Status !== "Completed" && dueDate < new Date();
  return (
    <Card
      className={styles.root}
      onClick={onClick}
      style={{
        border: isOverdue ? "2px solid #d13438" : undefined,
      }}
    >
      <CardHeader
        header={
          <div
            style={{
              display: "flex",
              justifyContent: "space-between",
              width: "100%",
            }}
          >
            {getStatusTag(task.Status)}
            {task.ProjectID && (
              <Text size={300} weight="semibold">
                Project: {task.ProjectID}
              </Text>
            )}
          </div>
        }
        description={
          <Text size={300} weight="semibold">
            Title: {task.Title}
          </Text>
        }
        action={
          task.AssignedTo?.Title && (
            <Persona name={task.AssignedTo.Title} size="extra-small" />
          )
        }
      />
      <Text className={styles.desc}>{task.Description}</Text>
      <div className={styles.pillRow}>
        <Text size={200}>
          Start: {new Date(task.StartDate || "").toLocaleDateString()}
        </Text>
        <Text size={200}>
          Due: {new Date(task.DueDate || "").toLocaleDateString()}
        </Text>
      </div>
      <CardFooter>{/* Optionally add quick edit dropdowns here */}</CardFooter>
    </Card>
  );
};

export * from "./TaskCard";
