import * as React from "react";
import { useState, useEffect } from "react";
import {
  Stack,
  Text,
  Dropdown,
  IDropdownOption,
  TextField,
  DatePicker,
  DefaultButton,
  PrimaryButton,
  Dialog,
  DialogType,
  DialogFooter,
  MessageBar,
  MessageBarType,
} from "@fluentui/react";
import styles from "./TasksList.module.scss";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import TasksService, { ITask } from "../../services/TasksService";

interface ITasksListProps {
  listName: string;
  userDisplayName: string;
  context: WebPartContext;
}

export const TasksList: React.FC<ITasksListProps> = (props) => {
  const [tasks, setTasks] = useState<ITask[]>([]);
  const [filteredTasks, setFilteredTasks] = useState<ITask[]>([]);
  const [isLoading, setIsLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);
  const [statusFilter, setStatusFilter] = useState<string>("all");
  const [priorityFilter, setPriorityFilter] = useState<string>("all");
  const [projectFilter, setProjectFilter] = useState<string>("all");
  const [isAddTaskDialogOpen, setIsAddTaskDialogOpen] = useState(false);
  const [newTask, setNewTask] = useState<Partial<ITask>>({});

  const tasksService = TasksService.getInstance(props.context);

  const statusOptions: IDropdownOption[] = [
    { key: "all", text: "All Statuses" },
    { key: "Not Started", text: "Not Started" },
    { key: "In Progress", text: "In Progress" },
    { key: "Completed", text: "Completed" },
  ];

  const priorityOptions: IDropdownOption[] = [
    { key: "all", text: "All Priorities" },
    { key: "High", text: "High" },
    { key: "Medium", text: "Medium" },
    { key: "Low", text: "Low" },
  ];

  const loadTasks = async (): Promise<void> => {
    try {
      setIsLoading(true);
      const items = await tasksService.getTasks(
        props.listName,
        props.userDisplayName
      );
      setTasks(items);
    } catch (err) {
      setError(err.message);
    } finally {
      setIsLoading(false);
    }
  };

  useEffect(() => {
    loadTasks().catch(console.error);
  }, []);

  const filterTasks = async (): Promise<void> => {
    let filtered = [...tasks];

    if (statusFilter !== "all") {
      filtered = filtered.filter((task) => task.Status === statusFilter);
    }

    if (priorityFilter !== "all") {
      filtered = filtered.filter((task) => task.Priority === priorityFilter);
    }

    if (projectFilter !== "all") {
      filtered = filtered.filter((task) => task.ProjectID === projectFilter);
    }

    setFilteredTasks(filtered);
  };

  useEffect(() => {
    filterTasks().catch(console.error);
  }, [tasks, statusFilter, priorityFilter, projectFilter]);

  const handleAddTask = async (): Promise<void> => {
    try {
      await tasksService.addTask(
        props.listName,
        newTask,
        props.userDisplayName
      );
      setIsAddTaskDialogOpen(false);
      setNewTask({});
      loadTasks().catch(console.error);
    } catch (err) {
      setError(err.message);
    }
  };

  if (isLoading) {
    return <Text>Loading tasks...</Text>;
  }

  if (error) {
    return (
      <MessageBar messageBarType={MessageBarType.error}>{error}</MessageBar>
    );
  }

  return (
    <Stack tokens={{ childrenGap: 15 }}>
      <Stack horizontal tokens={{ childrenGap: 10 }}>
        <Dropdown
          label="Status"
          options={statusOptions}
          selectedKey={statusFilter}
          onChange={(_, item) => setStatusFilter(item?.key as string)}
        />
        <Dropdown
          label="Priority"
          options={priorityOptions}
          selectedKey={priorityFilter}
          onChange={(_, item) => setPriorityFilter(item?.key as string)}
        />
        <Dropdown
          label="Project"
          options={[
            { key: "all", text: "All Projects" },
            ...Array.from(new Set(tasks.map((t) => t.ProjectID))).map((p) => ({
              key: p,
              text: p,
            })),
          ]}
          selectedKey={projectFilter}
          onChange={(_, item) => setProjectFilter(item?.key as string)}
        />
        <PrimaryButton
          text="Add Task"
          onClick={() => setIsAddTaskDialogOpen(true)}
        />
      </Stack>

      <Stack tokens={{ childrenGap: 10 }}>
        {filteredTasks.map((task) => (
          <Stack key={task.Id} className={styles.taskCard}>
            <Text variant="large">{task.Title}</Text>
            <Stack horizontal tokens={{ childrenGap: 10 }}>
              <Text>Due: {new Date(task.DueDate).toLocaleDateString()}</Text>
              <Text>Priority: {task.Priority}</Text>
              <Text>Status: {task.Status}</Text>
              <Text>Project: {task.ProjectID}</Text>
            </Stack>
          </Stack>
        ))}
      </Stack>

      <Dialog
        hidden={!isAddTaskDialogOpen}
        onDismiss={() => setIsAddTaskDialogOpen(false)}
        dialogContentProps={{
          type: DialogType.normal,
          title: "Add New Task",
        }}
      >
        <Stack tokens={{ childrenGap: 15 }}>
          <TextField
            label="Title"
            value={newTask.Title || ""}
            onChange={(_, value) => setNewTask({ ...newTask, Title: value })}
          />
          <DatePicker
            label="Due Date"
            value={newTask.DueDate ? new Date(newTask.DueDate) : undefined}
            onSelectDate={(date) =>
              setNewTask({ ...newTask, DueDate: date?.toISOString() })
            }
          />
          <Dropdown
            label="Priority"
            options={priorityOptions.filter((opt) => opt.key !== "all")}
            selectedKey={newTask.Priority}
            onChange={(_, item) =>
              setNewTask({ ...newTask, Priority: item?.key as string })
            }
          />
          <TextField
            label="Project"
            value={newTask.ProjectID || ""}
            onChange={(_, value) =>
              setNewTask({ ...newTask, ProjectID: value })
            }
          />
        </Stack>
        <DialogFooter>
          <DefaultButton
            onClick={() => setIsAddTaskDialogOpen(false)}
            text="Cancel"
          />
          <PrimaryButton onClick={handleAddTask} text="Add" />
        </DialogFooter>
      </Dialog>
    </Stack>
  );
};
