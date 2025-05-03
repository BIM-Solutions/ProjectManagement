import * as React from 'react';
import { Text, Stack, Separator, Persona, PersonaSize } from '@fluentui/react';
import { Project } from '../../../common/services/ProjectSelectionServices';

interface ProjectTeamProps {
  project: Project;
}

const ProjectTeam: React.FC<ProjectTeamProps> = ({ project }) => (
  <Stack tokens={{ childrenGap: 20 }} styles={{ root: { width: "35%" } }}>
    <Text variant="xLargePlus">Project Team</Text>
    <Separator />
    {project.PM && (
      <>
        <Text variant="mediumPlus"><strong>Project Manager</strong></Text>
        <Persona
          text={project.PM.Title}
          secondaryText={project.PM.JobTitle}
          tertiaryText={project.PM.Department}
          imageUrl={`/_layouts/15/userphoto.aspx?accountname=${encodeURIComponent(project.PM.Email)}`}
          size={PersonaSize.size48}
        />
        <Separator />
      </>
    )}
    {project.Manager && (
      <>
        <Text variant="mediumPlus"><strong>Information Manager</strong></Text>
        <Persona
          text={project.Manager.Title}
          secondaryText={project.Manager.JobTitle}
          tertiaryText={project.Manager.Department}
          imageUrl={`/_layouts/15/userphoto.aspx?accountname=${encodeURIComponent(project.Manager.Email)}`}
          size={PersonaSize.size48}
        />
        <Separator />
      </>
    )}
    {project.Checker && (
      <>
        <Text variant="mediumPlus"><strong>Project Checker</strong></Text>
        <Persona
          text={project.Checker.Title}
          secondaryText={project.Checker.JobTitle}
          tertiaryText={project.Checker.Department}
          imageUrl={`/_layouts/15/userphoto.aspx?accountname=${encodeURIComponent(project.Checker.Email)}`}
          size={PersonaSize.size48}
        />
        <Separator />
      </>
    )}
    {project.Approver && (
      <>
        <Text variant="mediumPlus"><strong>Project Approver</strong></Text>
        <Persona
          text={project.Approver.Title}
          secondaryText={project.Approver.JobTitle}
          tertiaryText={project.Approver.Department}
          imageUrl={`/_layouts/15/userphoto.aspx?accountname=${encodeURIComponent(project.Approver.Email)}`}
          size={PersonaSize.size48}
        />
        <Separator />
      </>
    )}
  </Stack>
);

export default ProjectTeam;