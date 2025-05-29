import * as React from "react";
import { useState, useContext, useEffect } from "react";
import {
  Button,
  Text,
  makeStyles,
  tokens,
  Option,
  Combobox,
  Field,
  Input,
  Textarea,
  Dialog,
  // DialogTrigger,
  DialogSurface,
  DialogBody,
  DialogContent,
  DialogActions,
  Card,
  CardHeader,
  Badge,
  SearchBox,
  Persona,
  DialogTitle,
} from "@fluentui/react-components";
import {
  Edit24Regular,
  Delete24Regular,
  Add24Regular,
  Search24Regular,
} from "@fluentui/react-icons";
import { SPContext } from "../../../common/components/SPContext";
import { DatePicker } from "@fluentui/react-datepicker-compat";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import {
  IPeoplePickerContext,
  PeoplePicker,
} from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { Project } from "../../services/ProjectSelectionServices";
import {
  taskPriorityOptions,
  taskStatusOptions,
  taskTypeOptions,
} from "../../services/ListService";

const listName = "9719_ProjectTasks";

export interface IUserField {
  Id: number;
  Title: string;
  EMail: string;
  JobTitle?: string;
  Department?: string;
  Presence?: string;
}

export interface TaskItem {
  Id: number;
  Title: string;
  TaskName: string;
  Status: string;
  DueDate: string;
  StartDate: string;
  AssignedTo?: IUserField;
  Description?: string;
  Comments?: string;
  Priority?: string;
  TaskType?: string;
  CreatedBy?: string;
  CreatedDate?: string;
  ModifiedBy?: string;
  ModifiedDate?: string;
  TaskID?: string;
  ProjectID?: string;
  Progress?: string;
}

const useStyles = makeStyles({
  container: {
    display: "flex",
    flexDirection: "row",
    gap: tokens.spacingHorizontalXL,
  },
  calendarPanel: { flex: 2, padding: tokens.spacingHorizontalM },
  rightPanel: {
    flex: 1,
    padding: tokens.spacingHorizontalM,
    backgroundColor: tokens.colorNeutralBackground2,
    border: `1px solid ${tokens.colorNeutralStroke1}`,
    borderRadius: "8px",
    display: "flex",
    flexDirection: "column",
    gap: tokens.spacingVerticalM,
  },
  taskList: {
    display: "flex",
    flexDirection: "column",
    gap: "8px",
    maxHeight: "400px",
    overflowY: "auto",
  },
  taskCard: {
    marginBottom: tokens.spacingVerticalS,
  },
  cardHeader: {
    display: "flex",
    justifyContent: "space-between",
    alignItems: "center",
    padding: tokens.spacingHorizontalS,
    width: "100%",
  },
  cardTitle: {
    flex: 1,
    marginRight: tokens.spacingHorizontalM,
  },
  cardActions: {
    display: "flex",
    gap: tokens.spacingHorizontalS,
    marginLeft: "auto",
  },
  detailRow: {
    display: "flex",
    padding: tokens.spacingVerticalXS + " " + tokens.spacingHorizontalS,
    borderBottom: `1px solid ${tokens.colorNeutralStroke1}`,
  },
  detailLabel: {
    flex: "0 0 120px",
    fontWeight: "bold",
  },
  detailValue: {
    flex: 1,
  },
  searchContainer: {
    display: "flex",
    gap: tokens.spacingHorizontalS,
    marginBottom: tokens.spacingVerticalM,
  },
  badge: {
    marginLeft: tokens.spacingHorizontalS,
  },
});

interface ProgrammeTabProps {
  project: Project;
  context: WebPartContext;
  tasks: TaskItem[];
  setTasks: (tasks: TaskItem[]) => void;
  selectedTask?: TaskItem;
  setSelectedTask: (task?: TaskItem) => void;
}

const ProgrammeTab: React.FC<ProgrammeTabProps> = ({
  project,
  context,
  tasks,
  setTasks,
  selectedTask,
  setSelectedTask,
}) => {
  const sp = useContext(SPContext);
  const styles = useStyles();
  const [showCreateModal, setShowCreateModal] = useState(false);
  const [showDeleteDialog, setShowDeleteDialog] = useState(false);
  const [assignedto, setAssignedTo] = useState("");
  const [planId, setPlanId] = useState<string>();
  const [searchQuery, setSearchQuery] = useState("");
  const [showEditModal, setShowEditModal] = useState(false);
  const [editingTask, setEditingTask] = useState<TaskItem | undefined>(
    undefined
  );

  const [newTask, setNewTask] = useState<TaskItem>({
    Id: 0,
    Title: "",
    TaskName: "",
    Description: "",
    StartDate: "",
    DueDate: "",
    AssignedTo: undefined,
    Progress: "",
    CreatedBy: "",
    Status: "",
    TaskType: "",
    Priority: "",
    ProjectID: "",
    TaskID: "",
    ModifiedBy: "",
    ModifiedDate: "",
    Comments: "",
    CreatedDate: "",
  });

  const peoplePickerContext: IPeoplePickerContext = {
    spHttpClient: context.spHttpClient,
    msGraphClientFactory: context.msGraphClientFactory,
    absoluteUrl: context.pageContext.web.absoluteUrl,
  };

  const safeEnsureUserId = async (email: string): Promise<number | null> => {
    if (!email) return null;
    try {
      const user = await sp.web.ensureUser(email);
      return user.Id;
    } catch {
      return null;
    }
  };

  const fetchTasks = async (): Promise<void> => {
    const results = await sp.web.lists
      .getByTitle(listName)
      .items.filter<TaskItem>(`ProjectID eq '${project.ProjectNumber}'`)
      .select(
        "Id",
        "Title",
        "TaskName",
        "Status",
        "DueDate",
        "StartDate",
        "Description",
        "Comments",
        "Priority",
        "TaskType",
        "CreatedBy",
        "CreatedDate",
        "ModifiedBy",
        "ModifiedDate",
        "TaskID",
        "ProjectID",
        "Progress"
      )
      .expand("AssignedTo", "CreatedBy", "ModifiedBy")
      .select(
        "Id",
        "Title",
        "TaskName",
        "Status",
        "DueDate",
        "StartDate",
        "Description",
        "Comments",
        "Priority",
        "TaskType",
        "CreatedDate",
        "ModifiedDate",
        "TaskID",
        "ProjectID",
        "Progress",
        "AssignedTo/Id",
        "AssignedTo/Title",
        "AssignedTo/EMail",
        "AssignedTo/JobTitle",
        "AssignedTo/Department",
        "CreatedBy/Id",
        "CreatedBy/Title",
        "CreatedBy/EMail",
        "CreatedBy/JobTitle",
        "CreatedBy/Department",
        "ModifiedBy/Id",
        "ModifiedBy/Title",
        "ModifiedBy/EMail",
        "ModifiedBy/JobTitle",
        "ModifiedBy/Department"
      )
      .top(5000)();
    setTasks(results);
    console.log("tasks", results);

    setPlanId(project.ProjectNumber);
  };

  // Add useEffect to fetch tasks on mount and project change
  useEffect(() => {
    if (project?.ProjectNumber) {
      fetchTasks().catch(console.error);
    }
  }, [project?.ProjectNumber]); // Re-run when project number changes

  const updateField = (field: string, value: string | number | ""): void => {
    setNewTask((prev) => ({ ...prev, [field]: value }));
  };

  const saveNewTask = async (): Promise<void> => {
    const userId = await safeEnsureUserId(assignedto);
    if (!userId) {
      console.error("Could not resolve user ID");
      return;
    }

    await sp.web.lists.getByTitle(listName).items.add({
      TaskName: newTask.Title,
      Description: newTask.Description,
      StartDate: newTask.StartDate,
      DueDate: newTask.DueDate,
      Title: newTask.Title,
      AssignedToId: userId,
      Status: newTask.Status,
      TaskType: newTask.TaskType,
      Priority: newTask.Priority,
      ProjectID: project.ProjectNumber,
      TaskID: "0",
      ModifiedById: userId,
      ModifiedDate: new Date().toISOString(),
      CreatedById: userId,
      CreatedDate: new Date().toISOString(),
      Comments: newTask.Comments,
      Progress: newTask.Progress,
    });
    setShowCreateModal(false);
    fetchTasks().catch(console.error);
  };

  const deleteTask = async (): Promise<void> => {
    if (!selectedTask?.Id) return;

    try {
      await sp.web.lists
        .getByTitle(listName)
        .items.getById(selectedTask.Id)
        .delete();

      setSelectedTask(undefined);
      if (planId) {
        await fetchTasks();
      }
    } catch (error) {
      console.error("Error deleting task:", error);
    }
  };

  const handleEditClick = (task: TaskItem): void => {
    setEditingTask({ ...task });
    setShowEditModal(true);
    if (task.AssignedTo?.EMail) {
      setAssignedTo(task.AssignedTo.EMail);
    }
  };

  const handleDeleteClick = (): void => {
    setShowDeleteDialog(true);
  };

  const confirmDelete = async (): Promise<void> => {
    await deleteTask();
    setShowDeleteDialog(false);
  };

  const saveEditedTask = async (): Promise<void> => {
    if (!editingTask?.Id) return;

    const userId = await safeEnsureUserId(assignedto);
    if (!userId) {
      console.error("Could not resolve user ID");
      return;
    }

    try {
      await sp.web.lists
        .getByTitle(listName)
        .items.getById(editingTask.Id)
        .update({
          TaskName: editingTask.Title,
          Description: editingTask.Description,
          StartDate: editingTask.StartDate,
          DueDate: editingTask.DueDate,
          Title: editingTask.Title,
          AssignedToId: userId,
          Status: editingTask.Status,
          TaskType: editingTask.TaskType,
          Priority: editingTask.Priority,
          Comments: editingTask.Comments,
          Progress: editingTask.Progress,
        });

      // Fetch the updated task to get the complete data with expanded fields
      const updatedTask = await sp.web.lists
        .getByTitle(listName)
        .items.getById(editingTask.Id)
        .select(
          "Id",
          "Title",
          "TaskName",
          "Status",
          "DueDate",
          "StartDate",
          "Description",
          "Comments",
          "Priority",
          "TaskType",
          "CreatedDate",
          "ModifiedDate",
          "TaskID",
          "ProjectID",
          "Progress",
          "AssignedTo/Id",
          "AssignedTo/Title",
          "AssignedTo/EMail"
        )
        .expand("AssignedTo")();

      setShowEditModal(false);
      setEditingTask(undefined);
      // Update the selected task with the new data
      setSelectedTask(updatedTask);

      // Refresh the task list
      await fetchTasks().catch((error) => {
        console.error("Error fetching tasks after update:", error);
      });
    } catch (error) {
      console.error("Error updating task:", error);
    }
  };

  // Filter tasks based on search query
  const filteredTasks = tasks.filter(
    (task) =>
      (task.Title?.toLowerCase() || "").includes(searchQuery.toLowerCase()) ||
      (task.Description?.toLowerCase() || "").includes(
        searchQuery.toLowerCase()
      ) ||
      (task.Status?.toLowerCase() || "").includes(searchQuery.toLowerCase())
  );

  const formatDate = (dateString: string): string => {
    if (!dateString) return "";
    return new Date(dateString).toLocaleDateString("en-GB", {
      day: "2-digit",
      month: "short",
      year: "numeric",
    });
  };

  return (
    <div className={styles.container}>
      <div className={styles.rightPanel}>
        {selectedTask ? (
          <Card className={styles.taskCard}>
            <CardHeader
              header={
                <div className={styles.cardHeader}>
                  <Text
                    size={500}
                    weight="semibold"
                    className={styles.cardTitle}
                  >
                    {selectedTask.Title}
                  </Text>
                  <div className={styles.cardActions}>
                    <Button
                      icon={<Edit24Regular />}
                      appearance="subtle"
                      onClick={() => handleEditClick(selectedTask)}
                    />
                    <Button
                      icon={<Delete24Regular />}
                      appearance="subtle"
                      onClick={handleDeleteClick}
                    />
                  </div>
                </div>
              }
            />
            <div className={styles.detailRow}>
              <Text className={styles.detailLabel}>Description:</Text>
              <Text className={styles.detailValue}>
                {selectedTask.Description || "No description"}
              </Text>
            </div>
            <div className={styles.detailRow}>
              <Text className={styles.detailLabel}>Status:</Text>
              <Badge
                appearance="filled"
                color={
                  selectedTask.Status === "Completed"
                    ? "success"
                    : "informative"
                }
                className={styles.badge}
              >
                {selectedTask.Status}
              </Badge>
            </div>
            <div className={styles.detailRow}>
              <Text className={styles.detailLabel}>Start Date:</Text>
              <Text className={styles.detailValue}>
                {formatDate(selectedTask.StartDate)}
              </Text>
            </div>
            <div className={styles.detailRow}>
              <Text className={styles.detailLabel}>Due Date:</Text>
              <Text className={styles.detailValue}>
                {formatDate(selectedTask.DueDate)}
              </Text>
            </div>
            <div className={styles.detailRow}>
              <Text className={styles.detailLabel}>Priority:</Text>
              <Badge
                appearance="filled"
                color={
                  selectedTask.Priority === "High"
                    ? "danger"
                    : selectedTask.Priority === "Medium"
                    ? "warning"
                    : "success"
                }
                className={styles.badge}
              >
                {selectedTask.Priority}
              </Badge>
            </div>
            <div className={styles.detailRow}>
              <Text className={styles.detailLabel}>Task Type:</Text>
              <Text className={styles.detailValue}>
                {selectedTask.TaskType}
              </Text>
            </div>
            <div className={styles.detailRow}>
              <Text className={styles.detailLabel}>Assigned To:</Text>
              <Persona
                name={selectedTask.AssignedTo?.Title || ""}
                avatar={{
                  image: {
                    src: `/_layouts/15/userphoto.aspx?accountname=${encodeURIComponent(
                      selectedTask.AssignedTo?.EMail || ""
                    )}`,
                    alt: selectedTask.AssignedTo?.Title,
                  },
                }}
              />
            </div>
          </Card>
        ) : (
          <Text>Select a task to view details</Text>
        )}

        <Text weight="semibold">Task List</Text>
        <div className={styles.searchContainer}>
          <SearchBox
            placeholder="Search tasks..."
            value={searchQuery}
            onChange={(_, data) => setSearchQuery(data.value)}
            contentBefore={<Search24Regular />}
          />
          <Button
            appearance="primary"
            icon={<Add24Regular />}
            onClick={() => setShowCreateModal(true)}
          >
            New Task
          </Button>
        </div>

        <div className={styles.taskList}>
          {filteredTasks.map((task) => (
            <Button
              key={task.Id}
              onClick={() => setSelectedTask(task)}
              appearance={selectedTask?.Id === task.Id ? "primary" : "subtle"}
              style={{
                justifyContent: "space-between",
                width: "100%",
                textAlign: "left",
              }}
            >
              <span>{task.Title}</span>
              <Badge
                appearance="filled"
                color={task.Status === "Completed" ? "success" : "informative"}
                className={styles.badge}
              >
                {formatDate(task.StartDate)}
              </Badge>
            </Button>
          ))}
          {filteredTasks.length === 0 && (
            <Text align="center">No tasks found</Text>
          )}
        </div>

        <Dialog open={showCreateModal}>
          {/* <DialogTrigger disableButtonEnhancement>
            <Button onClick={() => setShowCreateModal(true)}>Create Task</Button>
          </DialogTrigger> */}
          <DialogSurface>
            <DialogBody>
              <DialogContent>
                <Text size={400} weight="semibold">
                  New Task
                </Text>
                <Field label="Title" required>
                  <Input
                    value={newTask.Title}
                    onChange={(_, data) => updateField("Title", data.value)}
                  />
                </Field>
                <Field label="Description">
                  <Textarea
                    value={newTask.Description}
                    onChange={(_, data) =>
                      updateField("Description", data.value)
                    }
                  />
                </Field>
                <Field label="Start Date">
                  <DatePicker
                    value={
                      newTask.StartDate ? new Date(newTask.StartDate) : null
                    }
                    onSelectDate={(d) =>
                      updateField("StartDate", d?.toISOString() || "")
                    }
                  />
                </Field>
                <Field label="End Date">
                  <DatePicker
                    value={newTask.DueDate ? new Date(newTask.DueDate) : null}
                    onSelectDate={(d) =>
                      updateField("DueDate", d?.toISOString() || "")
                    }
                  />
                </Field>
                <Field label="Status" required>
                  <Combobox
                    value={newTask.Status || ""}
                    onOptionSelect={(_, data) =>
                      updateField("Status", data.optionText || "")
                    }
                  >
                    {taskStatusOptions.map((opt) => (
                      <Option key={opt.key} text={opt.value}>
                        {opt.value ?? ""}
                      </Option>
                    ))}
                  </Combobox>
                </Field>
                <Field label="Task Type" required>
                  <Combobox
                    value={newTask.TaskType || ""}
                    onOptionSelect={(_, data) =>
                      updateField("TaskType", data.optionText || "")
                    }
                  >
                    {taskTypeOptions.map((opt) => (
                      <Option key={opt.key} text={opt.value}>
                        {opt.value ?? ""}
                      </Option>
                    ))}
                  </Combobox>
                </Field>
                <Field label="Priority" required>
                  <Combobox
                    value={newTask.Priority || ""}
                    onOptionSelect={(_, data) =>
                      updateField("Priority", data.optionText || "")
                    }
                  >
                    {taskPriorityOptions.map((opt) => (
                      <Option key={opt.key} text={opt.value}>
                        {opt.value ?? ""}
                      </Option>
                    ))}
                  </Combobox>
                </Field>
                <Field label="Assigned To">
                  <PeoplePicker
                    context={peoplePickerContext}
                    personSelectionLimit={1}
                    principalTypes={[PrincipalType.User]}
                    resolveDelay={200}
                    onChange={(items) => {
                      // const selectedUsers = items.map(item => item.secondaryText).filter((text): text is string => text !== undefined);
                      setAssignedTo(items[0]?.secondaryText || "");
                    }}
                  />
                </Field>
              </DialogContent>
              <DialogActions>
                <Button appearance="primary" onClick={saveNewTask}>
                  Save
                </Button>
                <Button
                  appearance="secondary"
                  onClick={() => setShowCreateModal(false)}
                >
                  Cancel
                </Button>
              </DialogActions>
            </DialogBody>
          </DialogSurface>
        </Dialog>

        <Dialog open={showDeleteDialog}>
          <DialogSurface>
            <DialogBody>
              <DialogTitle>
                Confirm Delete: &quot;{selectedTask?.Title}&quot;?
              </DialogTitle>
              <DialogContent>
                This action is permanent. Are you sure you want to continue?
              </DialogContent>
              <DialogActions>
                <Button appearance="primary" onClick={confirmDelete}>
                  Delete
                </Button>
                <Button
                  appearance="secondary"
                  onClick={() => setShowDeleteDialog(false)}
                >
                  Cancel
                </Button>
              </DialogActions>
            </DialogBody>
          </DialogSurface>
        </Dialog>

        <Dialog open={showEditModal}>
          <DialogSurface>
            <DialogBody>
              <DialogContent>
                <Text size={400} weight="semibold">
                  Edit Task
                </Text>
                <Field label="Title" required>
                  <Input
                    value={editingTask?.Title || ""}
                    onChange={(_, data) =>
                      setEditingTask((prev) =>
                        prev ? { ...prev, Title: data.value } : prev
                      )
                    }
                  />
                </Field>
                <Field label="Description">
                  <Textarea
                    value={editingTask?.Description || ""}
                    onChange={(_, data) =>
                      setEditingTask((prev) =>
                        prev ? { ...prev, Description: data.value } : prev
                      )
                    }
                  />
                </Field>
                <Field label="Start Date">
                  <DatePicker
                    value={
                      editingTask?.StartDate
                        ? new Date(editingTask.StartDate)
                        : null
                    }
                    onSelectDate={(d) =>
                      setEditingTask((prev) =>
                        prev
                          ? { ...prev, StartDate: d?.toISOString() || "" }
                          : prev
                      )
                    }
                  />
                </Field>
                <Field label="End Date">
                  <DatePicker
                    value={
                      editingTask?.DueDate
                        ? new Date(editingTask.DueDate)
                        : null
                    }
                    onSelectDate={(d) =>
                      setEditingTask((prev) =>
                        prev
                          ? { ...prev, DueDate: d?.toISOString() || "" }
                          : prev
                      )
                    }
                  />
                </Field>
                <Field label="Status" required>
                  <Combobox
                    value={editingTask?.Status || ""}
                    onOptionSelect={(_, data) =>
                      setEditingTask((prev) =>
                        prev ? { ...prev, Status: data.optionText || "" } : prev
                      )
                    }
                  >
                    {taskStatusOptions.map((opt) => (
                      <Option key={opt.key} text={opt.value}>
                        {opt.value ?? ""}
                      </Option>
                    ))}
                  </Combobox>
                </Field>
                <Field label="Task Type" required>
                  <Combobox
                    value={editingTask?.TaskType || ""}
                    onOptionSelect={(_, data) =>
                      setEditingTask((prev) =>
                        prev
                          ? { ...prev, TaskType: data.optionText || "" }
                          : prev
                      )
                    }
                  >
                    {taskTypeOptions.map((opt) => (
                      <Option key={opt.key} text={opt.value}>
                        {opt.value ?? ""}
                      </Option>
                    ))}
                  </Combobox>
                </Field>
                <Field label="Priority" required>
                  <Combobox
                    value={editingTask?.Priority || ""}
                    onOptionSelect={(_, data) =>
                      setEditingTask((prev) =>
                        prev
                          ? { ...prev, Priority: data.optionText || "" }
                          : prev
                      )
                    }
                  >
                    {taskPriorityOptions.map((opt) => (
                      <Option key={opt.key} text={opt.value}>
                        {opt.value ?? ""}
                      </Option>
                    ))}
                  </Combobox>
                </Field>
                <Field label="Assigned To">
                  <PeoplePicker
                    context={peoplePickerContext}
                    personSelectionLimit={1}
                    principalTypes={[PrincipalType.User]}
                    resolveDelay={200}
                    defaultSelectedUsers={
                      editingTask?.AssignedTo
                        ? [editingTask.AssignedTo.EMail]
                        : []
                    }
                    onChange={(items) => {
                      setAssignedTo(items[0]?.secondaryText || "");
                    }}
                  />
                </Field>
              </DialogContent>
              <DialogActions>
                <Button appearance="primary" onClick={saveEditedTask}>
                  Save Changes
                </Button>
                <Button
                  appearance="secondary"
                  onClick={() => {
                    setShowEditModal(false);
                    setEditingTask(undefined);
                  }}
                >
                  Cancel
                </Button>
              </DialogActions>
            </DialogBody>
          </DialogSurface>
        </Dialog>
      </div>
    </div>
  );
};

export default ProgrammeTab;
