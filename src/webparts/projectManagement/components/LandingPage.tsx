import * as React from 'react';
import { useState } from 'react';
import {  Stack, Dialog, DialogType, PrimaryButton, DialogFooter, DefaultButton } from '@fluentui/react';
import CreateProjectForm from './CreateProjectForm';
import ProjectList from './ProjectList';
import { WebPartContext } from '@microsoft/sp-webpart-base';


interface ILandingPageProps {
  context: WebPartContext;
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

  return (
    <Stack tokens={{ padding: 20 }} styles={{ root: { width: '100%' } }}>
      
      <PrimaryButton text="➕ Add Project" onClick={() => setIsDialogOpen(true)} style={{ marginBottom: 10 }} />
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
        <CreateProjectForm onSuccess={() => setIsDialogOpen(false)} context={context} />
        <DialogFooter>
          <DefaultButton text="Cancel" onClick={() => setIsDialogOpen(false)} />
        </DialogFooter>
      </Dialog>
    </Stack>
  );
};

export default LandingPage;
