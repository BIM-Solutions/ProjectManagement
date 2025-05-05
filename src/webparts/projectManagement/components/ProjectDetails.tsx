import * as React from 'react';
import { useEffect, useState } from 'react';
import { Text, Stack } from '@fluentui/react';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { Project, ProjectSelectionService } from '../../common/services/ProjectSelectionServices';
import ProjectForm from '../../common/components/ProjectForm';
import ProjectTabs from './projectInformation/ProjectTabs';
import { DEBUG } from '../../common/DevVariables';
import { EventService } from '../../common/services/EventService';


interface ProjectDetailsProps {
  context: WebPartContext;
  onSave?: () => void;
  onEdit?: () => void;
  onDelete?: () => void;
}



const ProjectDetails: React.FC<ProjectDetailsProps> = ({ context, onEdit, onDelete }) => {
  const [project, setProject] = useState<Project | undefined>();
  const [isEditing, setIsEditing] = useState(false);

  useEffect(() => {
    const service = ProjectSelectionService;
    const listener = (selected: Project | undefined): void => {
      if (!selected) return setProject(undefined);
      if (!DEBUG) console.log('ProjectDetails listener - selected project:', selected);
      setProject({ ...selected });
    };
    service.subscribe(listener);
    listener(service.getSelectedProject());
    const unsubscribe = EventService.subscribeToProjectUpdates(() => {
      
      const latest = service.getSelectedProject();
      if (!DEBUG) console.log('ProjectDetails - updated project:', latest);
      setProject(latest? { ...latest } : undefined);
    });
    return () => {
      service.unsubscribe(listener);
      unsubscribe();
    };
  }, []);

  if (!project) return <Text>Select a project to view details.</Text>;
  if (!DEBUG) console.log('ProjectDetails - project:', project);

  return (
    <Stack tokens={{ childrenGap: 20 }} styles={{ root: { padding: 20, marginTop: 20, height: 'auto', minHeight: '50vh', overflowY: 'auto' } }}>
      {!isEditing ? (
        <ProjectTabs
          context={context}
          project={project}
          onEdit={() => setIsEditing(true)}
          onDelete={onDelete || (() => {})}
        />
      ) : (
        <ProjectForm
          context={context}
          mode="edit"
          project={project}
          onSuccess={() => {
            setIsEditing(false);
            if (onEdit) {
              onEdit();
            }

          }}
          onCancel={() => setIsEditing(false)}
        />
      )}
    </Stack>
  );
};

export default ProjectDetails;