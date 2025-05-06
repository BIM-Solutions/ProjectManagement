import * as React from 'react';
import { Text, Stack, Separator } from '@fluentui/react';
import { Image } from '@fluentui/react/lib/Image';
import { Project } from '../../services/ProjectSelectionServices';


interface ProjectOverviewProps {
  project: Project;
}

const ProjectOverview: React.FC<ProjectOverviewProps> = ({ project }) => {
  let projectImageElement = null;
  if (project.ProjectImage) {
    try {
      const image = JSON.parse(project.ProjectImage);
      projectImageElement = (
        <Image
          src={image.serverRelativeUrl}
          alt={image.fileName}
          style={{ height: 200, width: 'auto', marginBottom: 20 }}
        />
      );
    } catch (e) {
      console.warn('Invalid ProjectImage JSON:', e);
    }
  }

  return (
    <Stack tokens={{ childrenGap: 8 }}>
      {projectImageElement}
      <Text variant="large"><strong>Status:</strong> {project.Status}</Text>
      <Text variant="large"><strong>Project Name:</strong> {project.ProjectName}</Text>
      <Text variant="large"><strong>Project Number:</strong> {project.ProjectNumber}</Text>
      <Text variant="large"><strong>Project Description:</strong> {project.ProjectDescription}</Text>
      <Text variant="large"><strong>Sector:</strong> {project.Sector}</Text>
      <Text variant="large"><strong>Deltek Code:</strong> {project.DeltekSubCodes}</Text>
      <Text variant="large"><strong>Sub Codes:</strong> {project.SubCodes}</Text>
      <Text variant="large"><strong>Client:</strong> {project.Client}</Text>
      <Text variant="large"><strong>Client Contact:</strong> {project.ClientContact}</Text>
      <Separator />
    </Stack>
  );
};

export default ProjectOverview;