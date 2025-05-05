import * as React from 'react';
import { Stack, Text, Pivot, PivotItem, PrimaryButton, DefaultButton, Separator } from '@fluentui/react';
import { Project } from '../../../common/services/ProjectSelectionServices';
import ProjectOverview from './ProjectsOverview';
import ProjectTeam from './ProjectTeam';
import ProgrammeTab from '../projectCalender/ProgrammeTab';
import DocumentsTab from '../projectDocuments/DocumentsTab';
import StagesTab from '../projectStages/StagesTab';
import FeesTab from '../projectFees/FeesTab';
import { WebPartContext } from '@microsoft/sp-webpart-base';

interface ProjectTabsProps {
  project: Project;
  context: WebPartContext;
  onEdit: () => void;
  onDelete: () => void;
}

const ProjectTabs: React.FC<ProjectTabsProps> = ({ context, project, onEdit, onDelete }) => (
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
      <ProgrammeTab project={project} context={context} />
    </PivotItem>
    <PivotItem headerText="Stages">
      <StagesTab project={project} context={context} />
    </PivotItem>
    <PivotItem headerText="Documents">
      <DocumentsTab project={project} context={context} />
    </PivotItem>
    <PivotItem headerText="Fees">
      <FeesTab project={project} context={context} />
    </PivotItem>
  </Pivot>
);

export default ProjectTabs;