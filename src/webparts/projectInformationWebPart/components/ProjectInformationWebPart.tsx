import * as React from 'react';
import { useEffect, useState } from 'react';
import { Stack, Text, PrimaryButton, DefaultButton, Pivot, PivotItem, Persona, PersonaSize } from '@fluentui/react';
import { Project, ProjectSelectionService } from '../../common/services/ProjectSelectionServices';

interface ProjectDetailsProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  onSave?: () => void;
  onEdit?: () => void;
  onDelete?: () => void;
}

const ProjectDetails: React.FC<ProjectDetailsProps> = ({ onEdit, onDelete }) => {
  const [project, setProject] = useState<Project | undefined>(undefined);

  useEffect(() => {
    const service = ProjectSelectionService;

    const listener = (selected: Project | undefined): void => {
      if (!selected) {
        setProject(undefined);
        return;
      }

      setProject({
        id: selected.id,
        name: selected.name,
        number: selected.number,
        status: selected.status,
        client: selected.client || '',
        sector: selected.sector || '',
        pm: selected.pm,
        manager: selected.manager
      });
    };

    service.subscribe(listener);
    listener(service.getSelectedProject());

    return () => {
      service.unsubscribe(listener);
    };
  }, []);

  if (!project) {
    return <Text>Select a project to view details.</Text>;
  }
  console.log('ProjectDetails', project.pm);
  return (
    <Stack tokens={{ childrenGap: 20 }} styles={{ root: { padding: 20 } }}>
      <Stack horizontal tokens={{ childrenGap: 30 }} styles={{ root: { justifyContent: 'space-between' } }}>
        {/* Left column - Project info */}
        <Stack tokens={{ childrenGap: 10 }} styles={{ root: { flex: 1 } }}>
          <Text variant="xLarge">{project.number}</Text>
          <Text><strong>Name:</strong> {project.name}</Text>
          <Text><strong>Status:</strong> {project.status}</Text>
          <Text><strong>Sector:</strong> {project.sector}</Text>
          <Text><strong>Client:</strong> {project.client}</Text>
          <Text><strong>Email:</strong> {project.pm?.Email}</Text>
        </Stack>

        {/* Right column - People */}
        <Stack tokens={{ childrenGap: 20 }} styles={{ root: { width: 250 } }}>
          <Text variant="mediumPlus"><strong>Project Team</strong></Text>
          <Text variant="mediumPlus"><strong>Project Manager</strong></Text>
          <Persona
            text={project.pm?.Title}
            secondaryText={project.pm?.JobTitle || ''}
            tertiaryText={project.pm?.Department || ''}
            imageUrl={`/_layouts/15/userphoto.aspx?accountname=${encodeURIComponent(project.pm?.Email|| '')}`}
            size={PersonaSize.size72}
          />
          <Text variant="mediumPlus"><strong>Information Manager</strong></Text>
          <Persona
            text={project.manager?.Title}
            secondaryText= {project.manager?.JobTitle || ''}
            tertiaryText={project.manager?.Department || ''}
            imageUrl={`/_layouts/15/userphoto.aspx?accountname=${encodeURIComponent(project.manager?.Email || '')}`}
            size={PersonaSize.size72}
          />
        </Stack>
      </Stack>

      {/* Actions */}
      <Stack horizontal tokens={{ childrenGap: 10 }}>
        <PrimaryButton text="Edit" iconProps={{ iconName: 'Edit' }} onClick={onEdit} />
        <DefaultButton text="Delete" iconProps={{ iconName: 'Delete' }} onClick={onDelete} />
      </Stack>

      {/* Tabs */}
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
