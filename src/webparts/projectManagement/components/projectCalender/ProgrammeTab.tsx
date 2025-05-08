import * as React from 'react';
import { useEffect, useState, useContext } from 'react';
import {
  Stack,   DefaultButton, Modal, TextField, DatePicker, Dropdown
} from '@fluentui/react';
import {Button, Text,makeStyles, tokens } from '@fluentui/react-components';
import { SPContext } from '../common/SPContext';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { PrincipalType } from '@pnp/spfx-controls-react/lib/PeoplePicker';
import { IPeoplePickerContext, PeoplePicker } from '@pnp/spfx-controls-react/lib/PeoplePicker';
// import TaskCalendar from './TaskCalendar';
import { Project } from '../../services/ProjectSelectionServices';
import { taskPriorityOptions, taskStatusOptions, taskTypeOptions } from '../../services/ListService';


export interface TaskItem {
  Id: number;
  Title: string;
  Description?: string;
  StartDate?: string;
  DueDate?: string;
  AssignToId?: number;
  Progress?: string;
  CreatedBy?: string;
  Status?: string;
}
interface ProgrammeTabProps {
  context: WebPartContext;
  project: Project;
  tasks: TaskItem[];
  setTasks: React.Dispatch<React.SetStateAction<TaskItem[]>>;
  selectedTask: TaskItem | undefined;
  setSelectedTask: (task: TaskItem | undefined) => void;
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

const ProgrammeTab: React.FC<ProgrammeTabProps> = ({ project, context, tasks, setTasks, selectedTask, setSelectedTask }) => {
  const sp = useContext(SPContext);
  const styles = useStyles();

  // const [tasks, setTasks] = useState<TaskItem[]>([]);
  // const [selectedTask, setSelectedTask] = useState<TaskItem | null>(null);
  const [showCreateModal, setShowCreateModal] = useState(false);
  const [assignedto, setAssignedTo] = useState('');
  const [newTask, setNewTask] = useState<TaskItem>({
    Id: 0, Title: '', Description: '', StartDate: '', DueDate: '', AssignToId: undefined, Progress: '', CreatedBy: '', Status: ''
  });

  const peoplePickerContext: IPeoplePickerContext = {
    spHttpClient: context.spHttpClient,
    msGraphClientFactory: context.msGraphClientFactory,
    absoluteUrl: context.pageContext.web.absoluteUrl,
  };

  const fetchTasks = async (): Promise<void> => {
    const result = await sp.web.lists.getByTitle('9719_ProjectTasks').items
      .filter(`Title eq '${project.ProjectNumber}'`).top(5000)();
    setTasks(result);
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

  const updateField = (field: string, value: string | number | undefined): void => {
    setNewTask(prev => ({ ...prev, [field]: value }));
  };

  const saveNewTask = async (): Promise<void> => {
    const userId = await safeEnsureUserId(assignedto);
    await sp.web.lists.getByTitle('9719_ProjectTasks').items.add({
      Title: project.ProjectNumber,
      TaskName: newTask.Title,
      Description: newTask.Description,
      StartDate: newTask.StartDate,
      DueDate: newTask.DueDate,
      AssigntoId: userId,
      Status: newTask.Status,
    });
    setShowCreateModal(false);
    fetchTasks().catch((error) => console.error('Error fetching tasks:', error));
  };

  const deleteTask: () => Promise<void> = async () => {
    if (!selectedTask) return;
    await sp.web.lists.getByTitle('9719_ProjectTasks').items.getById(selectedTask.Id).delete();
    setSelectedTask(undefined);
    fetchTasks().catch((error) => console.error('Error fetching tasks:', error));
  };

  useEffect(() => { 
    fetchTasks().catch((error) => console.error('Error fetching tasks:', error)); 
  }, [project]);

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
            <Text><strong>Assigned To:</strong> {selectedTask.AssignToId}</Text>
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

        <Button onClick={() => setShowCreateModal(true)}>Create Task</Button>
        <DefaultButton text="Delete Task" disabled={!selectedTask} onClick={deleteTask} />
      </div>

      <Modal isOpen={showCreateModal} onDismiss={() => setShowCreateModal(false)} isBlocking={false}>
        <Stack tokens={{ padding: 20 }}>
          <Text size={500}>New Task</Text>
          <TextField label="Title" value={newTask.Title} onChange={(_, v) => updateField('Title', v || '')} />
          <TextField label="Description" value={newTask.Description} multiline rows={3} onChange={(_, v) => updateField('Description', v || '')} />
          <DatePicker label="Start Date" onSelectDate={(d) => updateField('StartDate', d?.toISOString() || '')} />
          <DatePicker label="End Date" onSelectDate={(d) => updateField('DueDate', d?.toISOString() || '')} />
          <Dropdown label="Status" options={taskStatusOptions} onChange={(_, o) => updateField('Status', o?.text || '')} />
          <Dropdown label="Task Type" options={taskTypeOptions} onChange={(_, o) => updateField('TaskType', o?.text || '')} />
          <Dropdown label="Priority" options={taskPriorityOptions} onChange={(_, o) => updateField('Priority', o?.text || '')} />
          <PeoplePicker
            context={peoplePickerContext}
            personSelectionLimit={1}
            principalTypes={[PrincipalType.User]}
            resolveDelay={200}
            onChange={(items) => {
              const selectedUser = items[0];
              if (selectedUser?.secondaryText) setAssignedTo(selectedUser.secondaryText);
            }}
          />
          <Button  onClick={saveNewTask}>Save</Button>
        </Stack>
      </Modal>
    </div>
  );
};

export default ProgrammeTab;
