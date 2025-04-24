import * as React from 'react';
import { Stack, Text, PrimaryButton, DefaultButton, Pivot, PivotItem } from '@fluentui/react';
import { IProjectInfo } from './ProjectList';

interface ProjectDetailsProps {
  project: IProjectInfo | undefined;
  onEdit?: () => void;
  onDelete?: () => void;
}

/**
 * A component that displays the details of a project.
 * If the project is undefined, it displays a message to select a project.
 * Otherwise, it displays the project number, name, sector, and information manager, and provides
 * buttons to edit or delete the project.
 * The component also provides a pivot to navigate to other components related to the project.
 * @param project The project to display, or undefined if no project is selected.
 * @param onEdit A callback to call when the 'Edit' button is clicked.
 * @param onDelete A callback to call when the 'Delete' button is clicked.
 */
const ProjectDetails: React.FC<ProjectDetailsProps> = ({ project, onEdit, onDelete }) => {
  if (project === undefined) {
    return <Text>Select a project to view details.</Text>;
  }

  return (
    <Stack tokens={{ childrenGap: 10 }} styles={{ root: { padding: 10 } }}>
      <Text variant="xLarge">{project.ProjectNumber}</Text>
      <Text><strong>Project Name:</strong> {project.ProjectName}</Text>
      <Text><strong>Sector:</strong> {project.Sector}</Text>
      <Text><strong>Information Manager:</strong> {project.Manager}</Text>

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
