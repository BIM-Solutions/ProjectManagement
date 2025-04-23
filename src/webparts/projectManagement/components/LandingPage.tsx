import * as React from 'react';
import { useState } from 'react';
import { Pivot, PivotItem, Stack } from '@fluentui/react';
import CreateProjectForm from './CreateProjectForm';
import ProjectList from './ProjectList';


const LandingPage: React.FC = () => {
  const [selectedKey, setSelectedKey] = useState<string>('view');

  return (
    <Stack tokens={{ padding: 20 }} styles={{ root: { width: '100%' } }}>
      <Pivot selectedKey={selectedKey} onLinkClick={(item) => setSelectedKey(item?.props.itemKey || 'view')}>
        <PivotItem headerText="📋 View Projects" itemKey="view">
          <ProjectList />
        </PivotItem>
        <PivotItem headerText="➕ Create Project" itemKey="create">
            <CreateProjectForm onSuccess={() => setSelectedKey('view')} />
        </PivotItem>
      </Pivot>
    </Stack>
  );
};

export default LandingPage;
