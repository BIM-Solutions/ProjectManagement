import * as React from 'react';
import { useEffect, useState, useContext } from 'react';
import {
  Stack, Text, DetailsList, DetailsListLayoutMode, SelectionMode, PrimaryButton,
  DefaultButton,  Modal, TextField, DatePicker,
  Dropdown
} from '@fluentui/react';
import { SPContext } from '../common/SPContext';
import { PrincipalType } from '@pnp/spfx-controls-react/lib/PeoplePicker';
import { WebPartContext } from '@microsoft/sp-webpart-base';

import { Project } from '../../services/ProjectSelectionServices';
import  TaskCalendar  from './TaskCalendar';
import './react-big-calendar.css';
import { taskPriorityOptions, taskStatusOptions, taskTypeOptions } from '../../services/ListService';
import { IPeoplePickerContext, PeoplePicker } from '@pnp/spfx-controls-react/lib/PeoplePicker';

interface ProgrammeTabProps {
  project: Project;
  context: WebPartContext;
}

export interface TaskItem {
  Id: number;
  Title: string;
  Description?: string;
  StartDate?: string;
  DueDate?: string;
  AssignToId?: number | undefined;
  Progress?: string;
  CreatedBy?: string;
  Status? : string;
}



const listName = '9719_ProjectTasks';

const ProgrammeTab: React.FC<ProgrammeTabProps> = ({ context, project }) => {
  const sp = useContext(SPContext);

  const [tasks, setTasks] = useState<TaskItem[]>([]);
  const [assignedto, setAssignedTo] = useState('');
  const [selectedTask, setSelectedTask] = useState<TaskItem | null>(null);
  const [isModalOpen, setIsModalOpen] = useState(false);
  const [showCreateModal, setShowCreateModal] = useState(false);

  const safeEnsureUserId = async (email: string): Promise<number | null> => {
    if (!email) return null;
    try {
      const user = await sp.web.ensureUser(email);
      return user.Id;
    } catch {
      return null;
    }
  };

  const [newTask, setNewTask] = useState<TaskItem>({
    Id: 0,
    Title: '',
    Description: '',
    StartDate: '',
    DueDate: '',
    AssignToId: undefined,
    Progress: '',
    CreatedBy: '',
    Status: '',
  });

  const peoplePickerContext: IPeoplePickerContext  = {
      spHttpClient: context.spHttpClient,
      msGraphClientFactory: context.msGraphClientFactory,
      absoluteUrl: context.pageContext.web.absoluteUrl
    };

  const fetchTasks = async (): Promise<void> => {
    const results = await sp.web.lists.getByTitle(listName).items.filter<TaskItem>(`Title eq '${project.ProjectNumber}'`).top(5000)();
    setTasks(results);
  };

  interface UpdateFieldFunction {
    (field: string, value: string | number | Date | null): void;
  }

  const updateField: UpdateFieldFunction = (field, value) =>
    setNewTask((prev) => ({ ...prev, [field]: value }));
  
  const saveNewTask = async (): Promise<void> => {
    const userId = await safeEnsureUserId(assignedto);
    await sp.web.lists.getByTitle(listName).items.add({
      TaskName: newTask.Title,
      Description: newTask.Description,
      StartDate: newTask.StartDate,
      DueDate: newTask.DueDate,
      Title: project.ProjectNumber,
      AssigntoId: userId,
      Status: newTask.Status,
    });
    setShowCreateModal(false);
    fetchTasks().catch(console.error);
  };

  useEffect(() => {
    fetchTasks().catch(console.error);
  }, [project]);

  const openModal = (task: TaskItem): void => {
    setSelectedTask(task);
    setIsModalOpen(true);
  };

  return (
    <Stack wrap tokens={{ childrenGap: 20 }} styles={{ root: { padding: 20 } }}>
      {/* Left: Vertical stack */}
      <Stack  tokens={{ childrenGap: 10 }} styles={{ root: { minWidth: '1000px' } }}>
        <Stack tokens={{ childrenGap: 10 }} styles={{ root: { width: '100%' } }}>
          <Text variant="large">Tasks</Text>
          <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 10 }} styles={{ root: { width: '100%' } }}>
            <PrimaryButton text="Create Task" onClick={() => setShowCreateModal(true)}  style={{ maxWidth: '250px' }}/>
            <DefaultButton text="Edit Task" disabled={!selectedTask} style={{ maxWidth: '250px' }}/>
            <DefaultButton text="Delete Task" disabled={!selectedTask} style={{ maxWidth: '250px' }}/>
          </Stack>

          <DetailsList
            items={tasks}
            columns={[
              { key: 'title', name: 'Title', fieldName: 'Title', minWidth: 100 },
              { key: 'start', name: 'Start Date', fieldName: 'StartDate', minWidth: 80 },
              { key: 'end', name: 'End Date', fieldName: 'DueDate', minWidth: 80 },
            ]}
            layoutMode={DetailsListLayoutMode.fixedColumns}
            selectionMode={SelectionMode.single}
            onItemInvoked={item => setSelectedTask(item)}
          />
        </Stack>

        {/* Center: Task Details */}
        <Stack tokens={{ childrenGap: 10 }} styles={{ root: { width: '100%' } }}>
          <Text variant="large">Task Details</Text>
          {selectedTask ? (
            <>
              <Text><strong>Title:</strong> {selectedTask.Title}</Text>
              <Text><strong>Description:</strong> {selectedTask.Description}</Text>
              <Text><strong>Start:</strong> {selectedTask.StartDate}</Text>
              <Text><strong>End:</strong> {selectedTask.DueDate}</Text>
              <Text><strong>Assigned To:</strong> {selectedTask.AssignToId}</Text>
              <Text><strong>Progress:</strong> {selectedTask.Progress}</Text>
              <Text><strong>Created By:</strong> {selectedTask.CreatedBy}</Text>
            </>
          ) : (
            <Text>No task selected.</Text>
          )}
        </Stack>
      </Stack>

      {/* Right: Calendar with click-to-view */}
      <Stack tokens={{ childrenGap: 10 }} styles={{ root: { flexGrow: 1 } }}>
        <Text variant="large">Calendar</Text>
        {/* Replace this with an actual calendar that supports item clicks */}
        <div className='customCalendar' >
          <TaskCalendar tasks={tasks} onTaskClick={openModal} />
        </div>
        {tasks.map((task) => (
          <Text key={task.Id} onClick={() => openModal(task)} style={{ cursor: 'pointer' }}>
            📅 {task.Title} ({task.StartDate})
          </Text>
        ))}
      </Stack>
      {showCreateModal && (
        <Modal isOpen={showCreateModal} onDismiss={() => setShowCreateModal(false)} isBlocking={false}>
          <Stack tokens={{ padding: 20 }}>
            <Text variant="large">New Task</Text>
            <TextField label="Title" value={newTask.Title} onChange={(_, v) => updateField('Title', v || '')} />
            <TextField label="Description" value={newTask.Description} multiline rows={3} onChange={(_, v) => updateField('Description', v || '')} />
            <DatePicker label="Start Date" onSelectDate={(d) => updateField('StartDate', d ? d.toISOString() : null)} />
            <DatePicker label="End Date" onSelectDate={(d) => updateField('DueDate', d ? d.toISOString() : null)} />
            <Dropdown label="Status" options={taskStatusOptions} onChange={(_, o) => updateField('Status', o?.text ?? '')} />
            <Dropdown label="Task Type" options={taskTypeOptions} onChange={(_, o) => updateField('TaskType', o?.text ?? '')} />
            <Dropdown label="Priority" options={taskPriorityOptions} onChange={(_, o) => updateField('Priority', o?.text ?? '')} />
            <PeoplePicker
              context={peoplePickerContext}
              personSelectionLimit={1}
              principalTypes={[PrincipalType.User]}
              resolveDelay={200}
              onChange={(items) => {
                const selectedUser = items[0];
                if (selectedUser?.secondaryText) {
                  setAssignedTo(selectedUser.secondaryText); // secondaryText = user email
                }
              }}
            />

            <PrimaryButton text="Save" onClick={saveNewTask} />
          </Stack>
        </Modal>
      )}

      {/* Pop-up modal for task */}
      <Modal isOpen={isModalOpen} onDismiss={() => setIsModalOpen(false)} isBlocking={false}>
        <Stack tokens={{ padding: 20 }}>
          <Text variant="xLarge">Task Info</Text>
          {selectedTask && (
            <Stack tokens={{ childrenGap: 10 }} styles={{ root: { width: '100%' } }}>
              <Text><strong>Title:</strong> {selectedTask.Title}</Text>
              <Text><strong>Description:</strong> {selectedTask.Description}</Text>
              <Text><strong>Start:</strong> {selectedTask.StartDate}</Text>
              <Text><strong>End:</strong> {selectedTask.DueDate}</Text>
              <Text><strong>Assigned To:</strong> {selectedTask.AssignToId}</Text>
            </Stack>
          )}
        </Stack>
      </Modal>
    </Stack>
  );
};

export default ProgrammeTab;
