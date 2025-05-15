import * as React from 'react';
import {  useState, useContext} from 'react';
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
  DialogTrigger,
  DialogSurface,
  DialogBody,
  DialogContent,
  DialogActions,
} from '@fluentui/react-components';
import { SPContext } from '../common/SPContext';
import { DatePicker } from '@fluentui/react-datepicker-compat';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { PrincipalType } from '@pnp/spfx-controls-react/lib/PeoplePicker';
import { IPeoplePickerContext, PeoplePicker } from '@pnp/spfx-controls-react/lib/PeoplePicker';
import { Project } from '../../services/ProjectSelectionServices';
import { taskPriorityOptions, taskStatusOptions, taskTypeOptions } from '../../services/ListService';


const listName = '9719_ProjectTasks';

export interface TaskItem {
  Id: number;
  Title: string;
  TaskName: string;
  Status: string;
  DueDate: string;
  StartDate: string;
  AssignToId?: number | undefined;
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
  container: { display: 'flex', flexDirection: 'row', gap: tokens.spacingHorizontalXL },
  calendarPanel: { flex: 2, padding: tokens.spacingHorizontalM },
  rightPanel: {
    flex: 1,
    padding: tokens.spacingHorizontalM,
    backgroundColor: tokens.colorNeutralBackground2,
    border: `1px solid ${tokens.colorNeutralStroke1}`,
    borderRadius: '8px',
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalM,
  },
  taskList: { display: 'flex', flexDirection: 'column', gap: '8px' },
});

interface ProgrammeTabProps {
  project: Project;
  context: WebPartContext;
  tasks: TaskItem[];
  setTasks: (tasks: TaskItem[]) => void;
  selectedTask?: TaskItem;
  setSelectedTask: (task?: TaskItem) => void;
}

const ProgrammeTab: React.FC<ProgrammeTabProps> = ({ project, context, tasks, setTasks, selectedTask, setSelectedTask }) => {
  const sp = useContext(SPContext);
  const styles = useStyles();
  const [status, setStatus] = useState<string | undefined>();
  const [taskType, setTaskType] = useState<string | undefined>();
  const [priority, setPriority] = useState<string | undefined>();
  const [showCreateModal, setShowCreateModal] = useState(false);
  const [assignedto, setAssignedTo] = useState('');
 
  const [planId, setPlanId] = useState<string>();
  
  const [newTask, setNewTask] = useState<TaskItem>({
    Id: 0,
    Title: '',
    TaskName: '',
    Description: '',
    StartDate: '',
    DueDate: '',
    AssignToId: undefined,
    Progress: '',
    CreatedBy: '',
    Status: '',
    TaskType: '',
    Priority: '',
    ProjectID: '',
    TaskID: '',
    ModifiedBy: '',
    ModifiedDate: '',
    Comments: '',
    CreatedDate: '',
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
    const results = await sp.web.lists.getByTitle(listName).items.filter<TaskItem>(`ProjectID eq '${project.ProjectNumber}'`).top(5000)();
    setTasks(results);
    setPlanId(project.ProjectNumber);
  };

  // interface UpdateFieldFunction {
  //   (field: string, value: string | number | Date | null): void;
  // }


  const updateField = (field: string, value: string | number | undefined): void => {
    setNewTask(prev => ({ ...prev, [field]: value }));
  };

  const saveNewTask = async (): Promise<void> => {
    const userId = await safeEnsureUserId(assignedto);
    await sp.web.lists.getByTitle(listName).items.add({
      TaskName: newTask.Title,
      Description: newTask.Description,
      StartDate: newTask.StartDate,
      DueDate: newTask.DueDate,
      Title: newTask.Title,
      AssigntoId: userId,
      Status: newTask.Status,
      TaskType: newTask.TaskType,
      Priority: newTask.Priority,
      ProjectID: project.ProjectNumber,
      TaskID: newTask.Id.toString(),
      ModifiedBy: userId,
      ModifiedDate: new Date().toISOString(),
      CreatedBy: userId,
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
      await sp.web.lists.getByTitle(listName).items.getById(selectedTask.Id).delete();

      setSelectedTask(undefined);
      if (planId) {
        await fetchTasks();
      }
    } catch (error) {
      console.error('Error deleting task:', error);
    }
  };

  return (
    <div className={styles.container}>
      <div className={styles.rightPanel}>
        <Text weight="semibold">Task Details</Text>
        {selectedTask ? (
          <>
            <Text><strong>Title:</strong> {selectedTask.Title}</Text>
            <Text><strong>Description:</strong> {selectedTask.Description}</Text>
            <Text><strong>Start:</strong> {selectedTask.StartDate}</Text>
            <Text><strong>End:</strong> {selectedTask.DueDate}</Text>
            <Text><strong>Assigned To:</strong> {Object.keys(selectedTask.AssignToId|| {}).join(', ')}</Text>
          </>
        ) : <Text>No task selected.</Text>}

        <Text weight="semibold">Task List</Text>
        <div className={styles.taskList}>
          {tasks.map(task => (
            <Button key={task.Id} onClick={() => setSelectedTask(task)} appearance="secondary">
              {task.Title} ({task.StartDate})
            </Button>
          ))}
        </div>

        <Dialog open={showCreateModal}>
          <DialogTrigger disableButtonEnhancement>
            <Button onClick={() => setShowCreateModal(true)}>Create Task</Button>
          </DialogTrigger>
          <DialogSurface>
            <DialogBody>
              <DialogContent>
                <Text size={400} weight="semibold">New Task</Text>
                <Field label="Title" required>
                  <Input value={newTask.Title} onChange={(_, data) => updateField('title', data.value)} />
                </Field>
                <Field label="Description">
                  <Textarea value={newTask.Description} onChange={(_, data) => updateField('description', data.value)} />
                </Field>
                <Field label="Start Date">
                  <DatePicker value={newTask.StartDate ? new Date(newTask.StartDate) : undefined} onSelectDate={(d) => updateField('StartDate', d?.toISOString())} />
                </Field>
                <Field label="End Date">
                  <DatePicker value={newTask.DueDate ? new Date(newTask.DueDate) : undefined} onSelectDate={(d) => updateField('DueDate', d?.toISOString())} />
                </Field>
                <Field label="Status" required>
                  <Combobox value={status} onOptionSelect={(_, data) => setStatus(data.optionText)}>
                    {taskStatusOptions.map(opt => <Option key={opt.key} text={opt.value}>{opt.value ?? ''}</Option>)}
                  </Combobox>
                </Field>
                <Field label="Task Type" required>
                  <Combobox value={taskType} onOptionSelect={(_, data) => setTaskType(data.optionText)}>
                    {taskTypeOptions.map(opt => <Option key={opt.key} text={opt.value}>{opt.value ?? ''}</Option>)}
                  </Combobox>
                </Field>
                <Field label="Priority" required>
                  <Combobox value={priority} onOptionSelect={(_, data) => setPriority(data.optionText)}>
                    {taskPriorityOptions.map(opt => <Option key={opt.key} text={opt.value}>{opt.value ?? ''}</Option>)}
                  </Combobox>
                </Field>
                <PeoplePicker
                  context={peoplePickerContext}
                  personSelectionLimit={1}
                  principalTypes={[PrincipalType.User]}
                  resolveDelay={200}
                  onChange={(items) => {
                    const selectedUsers = items.map(item => item.secondaryText).filter((text): text is string => text !== undefined);
                    setAssignedTo(selectedUsers[0] || '');
                  }}
                />
              </DialogContent>
              <DialogActions>
                <Button appearance="primary" onClick={saveNewTask}>Save</Button>
                <Button appearance="secondary" onClick={() => setShowCreateModal(false)}>Cancel</Button>
              </DialogActions>
            </DialogBody>
          </DialogSurface>
        </Dialog>

        <Button disabled={!selectedTask} onClick={deleteTask} appearance="secondary">Delete Task</Button>
      </div>
    </div>
  );
};

export default ProgrammeTab;
