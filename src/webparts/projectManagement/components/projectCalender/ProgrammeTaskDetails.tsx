// ProgrammeTaskDetails.tsx - Right container content using Fluent UI v2
import * as React from 'react';
// import { useState } from 'react';
import {
  Button,
  makeStyles,
  Text,
} from '@fluentui/react-components';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { TaskItem } from './ProgrammeTab';
import { Project } from '../../services/ProjectSelectionServices';

const useStyles = makeStyles({
  container: {
    display: 'flex',
    flexDirection: 'column',
    gap: '16px',
  },
});

interface Props {
  tasks: TaskItem[];
  selectedTask: TaskItem | undefined;
  setSelectedTask: (task: TaskItem | undefined ) => void;
  context: WebPartContext;
  project: Project;
  reloadTasks: () => void;
}

const ProgrammeTaskDetails: React.FC<Props> = ({
  tasks,
  selectedTask,
  setSelectedTask,
  context,
  project,
  reloadTasks,
}) => {
  const styles = useStyles();

  return (
    <div className={styles.container}>
      <Text weight="semibold" size={400}>Task List</Text>
      {tasks.map((task) => (
        <Button key={task.Id} appearance="secondary" onClick={() => setSelectedTask(task)}>
          {task.Title} ({task.StartDate})
        </Button>
      ))}

      <Text weight="semibold" size={400}>Task Details</Text>
      {selectedTask ? (
        <div>
          <Text><strong>Title:</strong> {selectedTask.Title}</Text>
          <Text><strong>Description:</strong> {selectedTask.Description}</Text>
          <Text><strong>Start:</strong> {selectedTask.StartDate}</Text>
          <Text><strong>End:</strong> {selectedTask.DueDate}</Text>
          <Text><strong>Assigned To:</strong> {selectedTask.AssignToId}</Text>
          <Text><strong>Progress:</strong> {selectedTask.Progress}</Text>
          <Text><strong>Created By:</strong> {selectedTask.CreatedBy}</Text>
        </div>
      ) : (
        <Text>No task selected.</Text>
      )}
    </div>
  );
};

export default ProgrammeTaskDetails;
