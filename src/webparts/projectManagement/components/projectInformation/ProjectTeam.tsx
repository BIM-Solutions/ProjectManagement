import * as React from 'react';
import { Text, Stack} from '@fluentui/react';
import { Project } from '../../services/ProjectSelectionServices';
import UserPanel from './UserPanel';
interface ProjectTeamProps {
  project: Project;
}

const ProjectTeam: React.FC<ProjectTeamProps> = ({ project }) => (
  <Stack tokens={{ childrenGap: 20 }} styles={{ root: { width: "35%" } }}>
    <Text variant="xLargePlus">Project Team</Text>
    {project.PM && (
      <>
      <UserPanel title="Project Manager" user={project.PM} />
      </>
    )}
    {project.Manager && (
      <>
      <UserPanel title="Information Manager" user={project.Manager} />
      </>
    )}
    {project.Checker && (
      <>
      <UserPanel title="Project Checker" user={project.Checker} />
      </>
    )}
    {project.Approver && (
      <>
      <UserPanel title="Project Approver" user={project.Approver} />
      </>
    )} 
  </Stack>
);

export default ProjectTeam;