import * as React from 'react';
import { useState, useEffect } from 'react';
import {
  makeStyles,
  tokens,
  Input,
  Button,
  Field,
  Textarea,
} from '@fluentui/react-components';
import { DatePicker } from '@fluentui/react-datepicker-compat';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { IProject, ProjectService } from '../../services/ProjectService';

export interface IProjectInfoTabProps {
  context: WebPartContext;
  project: IProject | undefined;
  projectService: ProjectService;
  onProjectChange: (project: IProject) => void;
}

const useStyles = makeStyles({
  container: {
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalM,
  },
  form: {
    display: 'grid',
    gridTemplateColumns: '1fr 1fr',
    gap: tokens.spacingHorizontalL,
  },
  fullWidth: {
    gridColumn: '1 / -1',
  },
  actions: {
    display: 'flex',
    justifyContent: 'flex-end',
    gap: tokens.spacingHorizontalM,
  },
});

const ProjectInfoTab: React.FC<IProjectInfoTabProps> = ({
  context,
  project,
  projectService,
  onProjectChange,
}) => {
  const styles = useStyles();
  const [formData, setFormData] = useState<Partial<IProject>>({});
  const [isEditing, setIsEditing] = useState(false);

  useEffect(() => {
    if (project) {
      setFormData(project);
    }
  }, [project]);

  const handleInputChange = (field: keyof IProject, value: string | number): void => {
    setFormData(prev => ({
      ...prev,
      [field]: value,
    }));
  };

  const handleSave = async (): Promise<void> => {
    try {
      if (project?.Id) {
        await projectService.updateProject(project.Id, formData);
        const updatedProject = await projectService.getProject(project.Id);
        onProjectChange(updatedProject);
      } else {
        const newProject = await projectService.createProject(formData);
        onProjectChange(newProject);
      }
      setIsEditing(false);
    } catch (error) {
      console.error('Error saving project:', error);
    }
  };

  if (!project && !isEditing) {
    return (
      <div className={styles.container}>
        <Button onClick={() => setIsEditing(true)}>Create New Project</Button>
      </div>
    );
  }

  return (
    <div className={styles.container}>
      <div className={styles.form}>
        <Field label="Project Title" required>
          <Input
            value={formData.Title || ''}
            onChange={(_, data) => handleInputChange('Title', data.value)}
            disabled={!isEditing}
          />
        </Field>

        <Field label="Project Number" required>
          <Input
            value={formData.ProjectNumber || ''}
            onChange={(_, data) => handleInputChange('ProjectNumber', data.value)}
            disabled={!isEditing}
          />
        </Field>

        <Field label="Project Manager">
          <Input
            value={formData.ProjectManager || ''}
            onChange={(_, data) => handleInputChange('ProjectManager', data.value)}
            disabled={!isEditing}
          />
        </Field>

        <Field label="Client Name">
          <Input
            value={formData.ClientName || ''}
            onChange={(_, data) => handleInputChange('ClientName', data.value)}
            disabled={!isEditing}
          />
        </Field>

        <Field label="Start Date">
          <DatePicker
            value={formData.StartDate ? new Date(formData.StartDate) : undefined}
            onSelectDate={(date: Date | null | undefined) =>
              handleInputChange('StartDate', date?.toISOString() || '')
            }
            disabled={!isEditing}
          />
        </Field>

        <Field label="End Date">
          <DatePicker
            value={formData.EndDate ? new Date(formData.EndDate) : undefined}
            onSelectDate={(date: Date | null | undefined) =>
              handleInputChange('EndDate', date?.toISOString() || '')
            }
            disabled={!isEditing}
          />
        </Field>

        <Field label="Budget">
          <Input
            type="number"
            value={formData.Budget?.toString() || ''}
            onChange={(_, data) => handleInputChange('Budget', Number(data.value))}
            disabled={!isEditing}
          />
        </Field>

        <Field label="Status">
          <Input
            value={formData.Status || ''}
            onChange={(_, data) => handleInputChange('Status', data.value)}
            disabled={!isEditing}
          />
        </Field>

        <Field label="Description" className={styles.fullWidth}>
          <Textarea
            value={formData.Description || ''}
            onChange={(_, data) => handleInputChange('Description', data.value)}
            disabled={!isEditing}
          />
        </Field>
      </div>

      <div className={styles.actions}>
        {isEditing ? (
          <>
            <Button appearance="secondary" onClick={() => setIsEditing(false)}>
              Cancel
            </Button>
            <Button appearance="primary" onClick={handleSave}>
              Save
            </Button>
          </>
        ) : (
          <Button appearance="primary" onClick={() => setIsEditing(true)}>
            Edit
          </Button>
        )}
      </div>
    </div>
  );
};

export default ProjectInfoTab; 