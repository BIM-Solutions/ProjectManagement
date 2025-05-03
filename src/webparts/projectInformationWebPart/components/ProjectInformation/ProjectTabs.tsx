import * as React from 'react';
import { Stack, Text, Pivot, PivotItem, PrimaryButton, DefaultButton, Separator } from '@fluentui/react';
import { Project } from '../../../common/services/ProjectSelectionServices';
import ProjectOverview from './ProjectsOverview';
import ProjectTeam from './ProjectTeam';

interface ProjectTabsProps {
  project: Project;
  onEdit: () => void;
  onDelete: () => void;
}

const ProjectTabs: React.FC<ProjectTabsProps> = ({ project, onEdit, onDelete }) => (
  <Pivot>
    <PivotItem headerText="Overview">
      <Stack tokens={{ childrenGap: 30 }} styles={{ root: { paddingTop: 20, justifyContent: 'space-between' } }} horizontal>
        <Stack tokens={{ childrenGap: 16 }} styles={{ root: { flex: 1 } }}>
          <Text variant="xxLargePlus">{project.ProjectNumber} - {project.ProjectName}</Text>
          <Separator />
          <ProjectOverview project={project} />
        </Stack>
        <Separator vertical />
        <ProjectTeam project={project} />
      </Stack>
      <div style={{ marginTop: 20, display: 'flex', justifyContent: 'center', gap: 20 }}>
        <PrimaryButton text="Edit" iconProps={{ iconName: 'Edit' }} onClick={onEdit} />
        <DefaultButton text="Delete" iconProps={{ iconName: 'Delete' }} onClick={onDelete} />
      </div>
    </PivotItem>
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
);

export default ProjectTabs;