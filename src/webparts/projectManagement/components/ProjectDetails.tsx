import * as React from 'react';
import { Stack, Text, PrimaryButton, DefaultButton, Pivot, PivotItem } from '@fluentui/react';
import { IProjectInfo } from './ProjectList';

interface ProjectDetailsProps {
  project: IProjectInfo | undefined;
  onEdit?: () => void;
  onDelete?: () => void;
}

const ProjectDetails: React.FC<ProjectDetailsProps> = ({ project, onEdit, onDelete }) => {
  if (project === undefined) {
    return <Text>Select a project to view details.</Text>;
  }

  return (
    <Stack tokens={{ childrenGap: 10 }} styles={{ root: { padding: 10 } }}>
      <Text variant="xLarge">{project.Title}</Text>
      <Text><strong>Project Number:</strong> {project.ProjectNumber}</Text>
      <Text><strong>Sector:</strong> {project.Sector}</Text>
      <Text><strong>Key Personnel:</strong> {project.KeyPersonnel}</Text>

      <Stack horizontal tokens={{ childrenGap: 10 }}>
        <PrimaryButton text="Edit" iconProps={{ iconName: 'Edit' }} onClick={onEdit} />
        <DefaultButton text="Delete" iconProps={{ iconName: 'Delete' }} onClick={onDelete} />
      </Stack>

      <Pivot>
        <PivotItem headerText="Programme">
          <Text>Programme tab coming soon...</Text>
        </PivotItem>
        <PivotItem headerText="Stages">
          <Text>Stages tab coming soon...</Text>
        </PivotItem>
        <PivotItem headerText="Documents">
          <Text>Documents tab coming soon...</Text>
        </PivotItem>
        <PivotItem headerText="Fees">
          <Text>Fees tab coming soon...</Text>
        </PivotItem>
      </Pivot>
    </Stack>
  );
};

export default ProjectDetails;
