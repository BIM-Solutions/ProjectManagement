import * as React from 'react';
import { useState, useEffect } from 'react';
import {
  makeStyles,
  tokens,
  Button,
  Table,
  TableHeader,
  TableRow,
  TableHeaderCell,
  TableBody,
  TableCell,
  Dialog,
  DialogSurface,
  DialogBody,
  DialogTitle,
  DialogContent,
  DialogActions,
  Input,
  Field,
  Dropdown,
  Option,
  Textarea,
} from '@fluentui/react-components';
import { DatePicker } from '@fluentui/react-datepicker-compat';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { IProject } from '../../services/ProjectService';
import { PeoplePicker, PrincipalType } from '@pnp/spfx-controls-react/lib/PeoplePicker';

export interface ITasksTabProps {
  context: WebPartContext;
  project: IProject | undefined;
}

interface ITask {
  Id: number;
  Title: string;
  Description: string;
  StartDate: string;
  DueDate: string;
  Status: string;
  Priority: string;
  AssignedTo: string;
  Progress: number;
}

const taskStatuses = [
  'Not Started',
  'In Progress',
  'Completed',
  'On Hold',
  'Blocked',
];

const taskPriorities = [
  'Low',
  'Medium',
  'High',
  'Critical',
];

const useStyles = makeStyles({
  container: {
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalM,
  },
  header: {
    display: 'flex',
    justifyContent: 'space-between',
    alignItems: 'center',
  },
  table: {
    width: '100%',
  },
  form: {
    display: 'grid',
    gridTemplateColumns: '1fr 1fr',
    gap: tokens.spacingHorizontalM,
  },
  fullWidth: {
    gridColumn: '1 / -1',
  },
  progress: {
    width: '100%',
  },
});

const TasksTab: React.FC<ITasksTabProps> = ({ context, project }) => {
  const styles = useStyles();
  const [tasks, setTasks] = useState<ITask[]>([]);
  const [showTaskDialog, setShowTaskDialog] = useState(false);
  const [selectedTask, setSelectedTask] = useState<ITask | null>(null);
  const [formData, setFormData] = useState<Partial<ITask>>({});

  const loadTasks = async (): Promise<void> => {
    if (project) {
      try {
        // This would be replaced with actual API call to load tasks
        const mockTasks: ITask[] = [
          {
            Id: 1,
            Title: 'Sample Task',
            Description: 'This is a sample task',
            StartDate: new Date().toISOString(),
            DueDate: new Date(Date.now() + 7 * 24 * 60 * 60 * 1000).toISOString(),
            Status: 'Not Started',
            Priority: 'Medium',
            AssignedTo: 'John Doe',
            Progress: 0,
          },
        ];
        setTasks(mockTasks);
      } catch (error) {
        console.error('Error loading tasks:', error);
      }
    }
  };

  useEffect(() => {
    if (project) {
      loadTasks().catch(console.error);
    }
  }, [project]);

  const handleInputChange = (field: keyof ITask, value: string | number): void => {
    setFormData(prev => ({
      ...prev,
      [field]: value,
    }));
  };

  const handleSave = async (): Promise<void> => {
    try {
      if (selectedTask) {
        // Update existing task
        const updatedTasks = tasks.map(task =>
          task.Id === selectedTask.Id ? { ...task, ...formData } : task
        );
        setTasks(updatedTasks);
      } else {
        // Create new task
        const newTask = {
          ...formData,
          Id: Math.max(...tasks.map(t => t.Id), 0) + 1,
        } as ITask;
        setTasks([...tasks, newTask]);
      }
      setShowTaskDialog(false);
      setSelectedTask(null);
      setFormData({});
    } catch (error) {
      console.error('Error saving task:', error);
    }
  };

  const handleDelete = async (taskId: number): Promise<void> => {
    try {
      setTasks(tasks.filter(task => task.Id !== taskId));
    } catch (error) {
      console.error('Error deleting task:', error);
    }
  };

  const handleEdit = (task: ITask): void => {
    setSelectedTask(task);
    setFormData(task);
    setShowTaskDialog(true);
  };

  return (
    <div className={styles.container}>
      <div className={styles.header}>
        <h2>Project Tasks</h2>
        <Button
          appearance="primary"
          onClick={() => {
            setSelectedTask(null);
            setFormData({});
            setShowTaskDialog(true);
          }}
        >
          New Task
        </Button>
      </div>

      <Table className={styles.table}>
        <TableHeader>
          <TableRow>
            <TableHeaderCell>Title</TableHeaderCell>
            <TableHeaderCell>Status</TableHeaderCell>
            <TableHeaderCell>Priority</TableHeaderCell>
            <TableHeaderCell>Assigned To</TableHeaderCell>
            <TableHeaderCell>Due Date</TableHeaderCell>
            <TableHeaderCell>Progress</TableHeaderCell>
            <TableHeaderCell>Actions</TableHeaderCell>
          </TableRow>
        </TableHeader>
        <TableBody>
          {tasks.map((task) => (
            <TableRow key={task.Id}>
              <TableCell>{task.Title}</TableCell>
              <TableCell>{task.Status}</TableCell>
              <TableCell>{task.Priority}</TableCell>
              <TableCell>{task.AssignedTo}</TableCell>
              <TableCell>{new Date(task.DueDate).toLocaleDateString()}</TableCell>
              <TableCell>{task.Progress}%</TableCell>
              <TableCell>
                <Button appearance="subtle" onClick={() => handleEdit(task)}>
                  Edit
                </Button>
                <Button appearance="subtle" onClick={() => handleDelete(task.Id)}>
                  Delete
                </Button>
              </TableCell>
            </TableRow>
          ))}
        </TableBody>
      </Table>

      <Dialog open={showTaskDialog}>
        <DialogSurface>
          <DialogBody>
            <DialogTitle>
              {selectedTask ? 'Edit Task' : 'New Task'}
            </DialogTitle>
            <DialogContent>
              <div className={styles.form}>
                <Field label="Title" required>
                  <Input
                    value={formData.Title || ''}
                    onChange={(_, data) => handleInputChange('Title', data.value)}
                  />
                </Field>

                <Field label="Status">
                  <Dropdown
                    value={formData.Status || ''}
                    onOptionSelect={(_, data) => handleInputChange('Status', data.optionValue || '')}
                  >
                    {taskStatuses.map((status) => (
                      <Option key={status} value={status}>
                        {status}
                      </Option>
                    ))}
                  </Dropdown>
                </Field>

                <Field label="Priority">
                  <Dropdown
                    value={formData.Priority || ''}
                    onOptionSelect={(_, data) => handleInputChange('Priority', data.optionValue || '')}
                  >
                    {taskPriorities.map((priority) => (
                      <Option key={priority} value={priority}>
                        {priority}
                      </Option>
                    ))}
                  </Dropdown>
                </Field>

                <Field label="Progress">
                  <Input
                    type="number"
                    min={0}
                    max={100}
                    value={formData.Progress?.toString() || '0'}
                    onChange={(_, data) => handleInputChange('Progress', Number(data.value))}
                    className={styles.progress}
                  />
                </Field>

                <Field label="Start Date">
                  <DatePicker
                    value={formData.StartDate ? new Date(formData.StartDate) : undefined}
                    onSelectDate={(date: Date | null | undefined) =>
                      handleInputChange('StartDate', date?.toISOString() || '')
                    }
                  />
                </Field>

                <Field label="Due Date">
                  <DatePicker
                    value={formData.DueDate ? new Date(formData.DueDate) : undefined}
                    onSelectDate={(date: Date | null | undefined) =>
                      handleInputChange('DueDate', date?.toISOString() || '')
                    }
                  />
                </Field>

                <Field label="Assigned To" className={styles.fullWidth}>
                  <PeoplePicker
                    context={{
                      spHttpClient: context.spHttpClient,
                      msGraphClientFactory: context.msGraphClientFactory,
                      absoluteUrl: context.pageContext.web.absoluteUrl,
                    }}
                    personSelectionLimit={1}
                    showHiddenInUI={false}
                    principalTypes={[PrincipalType.User]}
                    resolveDelay={1000}
                    onChange={(items) => {
                      if (items.length > 0) {
                        handleInputChange('AssignedTo', items[0].text || '');
                      }
                    }}
                  />
                </Field>

                <Field label="Description" className={styles.fullWidth}>
                  <Textarea
                    value={formData.Description || ''}
                    onChange={(_, data) => handleInputChange('Description', data.value)}
                  />
                </Field>
              </div>
            </DialogContent>
            <DialogActions>
              <Button appearance="secondary" onClick={() => setShowTaskDialog(false)}>
                Cancel
              </Button>
              <Button appearance="primary" onClick={handleSave}>
                Save
              </Button>
            </DialogActions>
          </DialogBody>
        </DialogSurface>
      </Dialog>
    </div>
  );
};

export default TasksTab; 