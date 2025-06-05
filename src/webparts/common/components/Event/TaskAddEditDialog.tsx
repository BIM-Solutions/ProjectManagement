import * as React from "react";
import {
  Dialog,
  DialogSurface,
  DialogTitle,
  DialogBody,
  DialogActions,
  Button,
  Input,
  // Text,
  Textarea,
  Dropdown,
  Option,
  Field,
} from "@fluentui/react-components";
import { TimePicker } from "@fluentui/react-timepicker-compat";
import { ITask } from "../../../resourcing/services/TasksService";
import {
  IProject,
  ProjectService,
} from "../../../projectManagement/services/ProjectService";
import { WebPartContext } from "@microsoft/sp-webpart-base";

interface TaskAddEditDialogProps {
  open: boolean;
  onClose: () => void;
  onSave: (task: Partial<ITask>) => void;
  mode: "add" | "edit" | "view";
  task?: ITask;
}

const priorityOptions = [
  { key: "High", text: "High" },
  { key: "Medium", text: "Medium" },
  { key: "Low", text: "Low" },
];

export const TaskAddEditDialog: React.FC<
  TaskAddEditDialogProps & { context: WebPartContext }
> = ({ open, onClose, onSave, context, mode, task }) => {
  const [title, setTitle] = React.useState("");
  const [startDate, setStartDate] = React.useState("");
  const [startTime, setStartTime] = React.useState("");
  const [dueDate, setDueDate] = React.useState("");
  const [dueTime, setDueTime] = React.useState("");
  const [priority, setPriority] = React.useState("Medium");
  const [projectId, setProjectId] = React.useState("");
  const [description, setDescription] = React.useState("");
  const [projects, setProjects] = React.useState<IProject[]>([]);
  const [projectsLoading, setProjectsLoading] = React.useState(false);
  const [projectSearch, setProjectSearch] = React.useState("");

  React.useEffect(() => {
    if (open) {
      if (mode === "edit" || mode === "view") {
        setTitle(task?.Title || "");
        setStartDate(task?.StartDate?.split("T")[0] || "");
        setStartTime(task?.StartDate?.split("T")[1]?.slice(0, 5) || "");
        setDueDate(task?.DueDate?.split("T")[0] || "");
        setDueTime(task?.DueDate?.split("T")[1]?.slice(0, 5) || "");
        setPriority(task?.Priority || "Medium");
        setProjectId(task?.ProjectID || "");
        setDescription(task?.Description || "");
      } else {
        setTitle("");
        setStartDate("");
        setStartTime("");
        setDueDate("");
        setDueTime("");
        setPriority("Medium");
        setProjectId("");
        setDescription("");
      }
      setProjectsLoading(true);
      const service = new ProjectService(context);
      service
        .getAllProjects()
        .then((data) => {
          setProjects(data);
          setProjectsLoading(false);
        })
        .catch(() => setProjectsLoading(false));
    }
  }, [open, context, mode, task]);

  const handleSave = (): void => {
    onSave({
      Title: title,
      DueDate: dueDate,
      Priority: priority,
      ProjectID: projectId,
      Description: description,
    });
  };

  return (
    <Dialog
      open={open}
      onOpenChange={(_, data) => {
        if (!data.open) onClose();
      }}
    >
      <DialogSurface style={{ padding: 32, maxWidth: 480 }}>
        <DialogTitle>
          {mode === "edit"
            ? "Edit Task"
            : mode === "view"
            ? "Task"
            : "Add New Task"}
        </DialogTitle>
        <DialogBody
          style={{
            display: "flex",
            flexDirection: "column",
            gap: 20,
            // minWidth: 320,
            width: "100%",
            boxSizing: "border-box",
          }}
        >
          <Field label="Title" required={mode !== "view"}>
            <Input
              id="task-title"
              value={title}
              onChange={(_, d) => setTitle(d.value)}
              placeholder="Title"
              required
              size="large"
              style={{
                width: "100%",
                pointerEvents: mode === "view" ? "none" : "auto",
              }}
              readOnly={mode === "view"}
              tabIndex={mode === "view" ? -1 : 0}
            />
          </Field>
          <div style={{ display: "flex", flexDirection: "row", gap: 10 }}>
            <Field label="Start Date" required={mode !== "view"}>
              <Input
                id="task-start-date"
                type="date"
                value={startDate}
                onChange={(_, d) => setStartDate(d.value)}
                size="large"
                style={{ pointerEvents: mode === "view" ? "none" : "auto" }}
                readOnly={mode === "view"}
                tabIndex={mode === "view" ? -1 : 0}
              />
            </Field>
            <Field label="Start Time" required={mode !== "view"}>
              <TimePicker
                value={startTime}
                onTimeChange={(_, data) =>
                  setStartTime(data.selectedTimeText ?? "")
                }
                size="large"
                style={{ pointerEvents: mode === "view" ? "none" : "auto" }}
                readOnly={mode === "view"}
                tabIndex={mode === "view" ? -1 : 0}
                increment={15}
              />
            </Field>
          </div>
          <div style={{ display: "flex", flexDirection: "row", gap: 10 }}>
            <Field label="Due Date" required={mode !== "view"}>
              <Input
                id="task-due-date"
                type="date"
                value={dueDate}
                onChange={(_, d) => setDueDate(d.value)}
                size="large"
                style={{ pointerEvents: mode === "view" ? "none" : "auto" }}
                readOnly={mode === "view"}
                tabIndex={mode === "view" ? -1 : 0}
              />
            </Field>
            <Field label="Due Time" required={mode !== "view"}>
              <TimePicker
                value={dueTime}
                onTimeChange={(_, data) =>
                  setDueTime(data.selectedTimeText ?? "")
                }
                size="large"
                style={{ pointerEvents: mode === "view" ? "none" : "auto" }}
                readOnly={mode === "view"}
                tabIndex={mode === "view" ? -1 : 0}
                increment={15}
              />
            </Field>
          </div>
          <Field label="Priority">
            <Dropdown
              id="task-priority"
              value={priority}
              onOptionSelect={(_, d) => setPriority(d.optionValue as string)}
              placeholder="Priority"
              size="large"
              style={{ width: "100%" }}
            >
              {priorityOptions.map((opt) => (
                <Option key={opt.key} value={opt.key}>
                  {opt.text}
                </Option>
              ))}
            </Dropdown>
          </Field>
          <Field label="Project">
            <Input
              id="project-search"
              value={projectSearch}
              onChange={(_, d) => setProjectSearch(d.value)}
              placeholder="Search projects..."
              size="large"
              disabled={projectsLoading}
              style={{ width: "100%", marginBottom: 8 }}
            />
            <Dropdown
              id="task-project"
              value={projectId}
              onOptionSelect={(_, d) => setProjectId(d.optionValue as string)}
              placeholder={
                projectsLoading ? "Loading projects..." : "Select a project"
              }
              size="large"
              disabled={projectsLoading}
              style={{ width: "100%" }}
            >
              {projects
                .filter((p) =>
                  p.Title.toLowerCase().includes(projectSearch.toLowerCase())
                )
                .map((project) => (
                  <Option key={project.Id} value={project.Id.toString()}>
                    {project.Title}
                  </Option>
                ))}
            </Dropdown>
          </Field>
          <Field label="Description">
            <Textarea
              id="task-description"
              value={description}
              onChange={(_, d) => setDescription(d.value)}
              placeholder="Description"
              size="large"
              rows={3}
              style={{ width: "100%" }}
            />
          </Field>
        </DialogBody>
        <DialogActions style={{ justifyContent: "flex-end" }}>
          <Button appearance="secondary" onClick={onClose}>
            Cancel
          </Button>
          <Button
            appearance="primary"
            onClick={handleSave}
            disabled={!title || !dueDate}
          >
            Save
          </Button>
        </DialogActions>
      </DialogSurface>
    </Dialog>
  );
};
