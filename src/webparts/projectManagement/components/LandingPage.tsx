import * as React from 'react';
import { useState, useEffect } from 'react';
import {  Stack, Dialog, DialogType, PrimaryButton, DialogFooter, DefaultButton } from '@fluentui/react';
// import CreateProjectForm from './CreateProjectForm';
import ProjectList from './ProjectList';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { useLoading } from '../services/LoadingContext';
import ProjectForm from '../../common/components/ProjectForm';
import { ProjectSelectionService } from '../../common/services/ProjectSelectionServices';
import { Project } from '../../common/services/ProjectSelectionServices';
import ProjectDetails from './ProjectDetails';


interface ILandingPageProps {
  context: WebPartContext;
}

interface CustomWindow {
  __setListLoading?: (isLoading: boolean) => void;
}


/**
 * A component that displays a pivot control with two items: "View Projects" and "Create Project".
 * The first item displays the ProjectList component, and the second item displays the CreateProjectForm
 * component. The component is connected to the WebPartContext, and it's used to load the list of projects
 * from the "9719_ProjectInformationDatabase" list.
 * @param props The properties of the component, including the WebPartContext.
 */
const LandingPage: React.FC<ILandingPageProps> = ({ context }) => {
  // const [selectedKey, setSelectedKey] = useState<string>('view');
  const [isDialogOpen, setIsDialogOpen] = useState(false);
  const { setIsLoading } = useLoading();
  const [selectedProject, setSelectedProject] = useState<Project | undefined>();

  useEffect(() => {
    const service = ProjectSelectionService;
    const listener = (selected: Project | undefined): void => {
      setSelectedProject(selected);
    };
    service.subscribe(listener);
    return () => {
      service.unsubscribe(listener);
    };
  }, []);
  useEffect(() => {
    (window as CustomWindow).__setListLoading = setIsLoading;
    return () => { delete (window as CustomWindow).__setListLoading; };
  }, [setIsLoading]);
  return (
    <div>
      {selectedProject && (
        <ProjectDetails
          
          context={context}
          
        />
      )}

      <Stack tokens={{ padding: 20 }} styles={{ root: { width: '100%' , height: 'auto' } }}>
      
      <PrimaryButton text="Add Project" iconProps={{ iconName: 'Add' }} onClick={() => setIsDialogOpen(true)} style={{ marginBottom: 10, width : '250px'}} />

      <ProjectList />
     

      <Dialog
        minWidth={740}
        maxWidth={1000}
        hidden={!isDialogOpen}
        onDismiss={() => setIsDialogOpen(false)}
        dialogContentProps={{
          type: DialogType.largeHeader,
          title: 'Create New Project',
        }}
        modalProps={{
          isBlocking: false,
          styles: {
            main: {
              selectors: {
                ['@media (min-width: 480px)']: {
                  width: 650,
                  minWidth: 650,
                  maxWidth: '1000px'
                }
              }
            }
          }
        }}
      >
        <ProjectForm onSuccess={() => setIsDialogOpen(false)} context={context} mode="create"/>
        {/* <CreateProjectForm onSuccess={() => setIsDialogOpen(false)} context={context} /> */}
        <DialogFooter>
          <DefaultButton text="Cancel" onClick={() => setIsDialogOpen(false)} />
        </DialogFooter>
      </Dialog>
    </Stack>
    </div>
  );
};

export default LandingPage;
