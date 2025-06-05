import * as React from "react";
import {
  Dialog,
  DialogSurface,
  DialogTitle,
  DialogBody,
  DialogActions,
  Button,
} from "@fluentui/react-components";
import { ITask } from "../../services/TasksService";

export interface TaskDetailsDialogProps {
  task: ITask;
  open: boolean;
  onClose: () => void;
}

export const TaskDetailsDialog: React.FC<TaskDetailsDialogProps> = ({
  task,
  open,
  onClose,
}) => {
  return (
    <Dialog
      open={open}
      onOpenChange={(_, data) => {
        if (!data.open) onClose();
      }}
    >
      <DialogSurface>
        <DialogTitle>{task.Title}</DialogTitle>
        <DialogBody>
          <div>
            <strong>Description:</strong> {task.Description || "-"}
          </div>
          <div>
            <strong>Due Date:</strong>{" "}
            {new Date(task.DueDate).toLocaleDateString()}
          </div>
          <div>
            <strong>Start Date:</strong>{" "}
            {task.StartDate
              ? new Date(task.StartDate).toLocaleDateString()
              : "-"}
          </div>
          <div>
            <strong>Priority:</strong> {task.Priority}
          </div>
          <div>
            <strong>Status:</strong> {task.Status}
          </div>
          <div>
            <strong>Project:</strong> {task.ProjectID}
          </div>
          <div>
            <strong>Assigned To:</strong> {task.AssignedTo?.Title || "-"}
          </div>
        </DialogBody>
        <DialogActions>
          <Button appearance="secondary" onClick={onClose}>
            Close
          </Button>
        </DialogActions>
      </DialogSurface>
    </Dialog>
  );
};

export * from "./TaskDetailsDialog";
