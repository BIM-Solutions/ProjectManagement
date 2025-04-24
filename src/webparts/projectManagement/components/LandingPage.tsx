import * as React from 'react';
import { useState } from 'react';
import { Pivot, PivotItem, Stack } from '@fluentui/react';
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
  const [selectedKey, setSelectedKey] = useState<string>('view');

  return (
    <Stack tokens={{ padding: 20 }} styles={{ root: { width: '100%' } }}>
      <Pivot selectedKey={selectedKey} onLinkClick={(item) => setSelectedKey(item?.props.itemKey || 'view')}>
        <PivotItem headerText="📋 View Projects" itemKey="view">
          <ProjectList />
        </PivotItem>
        <PivotItem headerText="➕ Create Project" itemKey="create">
          <CreateProjectForm onSuccess={() => setSelectedKey('view')} context={context} />
        </PivotItem>
      </Pivot>
    </Stack>
  );
};

export default LandingPage;
