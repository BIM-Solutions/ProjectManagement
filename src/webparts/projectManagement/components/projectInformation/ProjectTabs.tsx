import * as React from 'react';
import {
  makeStyles,
  Tab,
  TabList,
  TabValue,
  Text,
  Button,
  Divider,
} from '@fluentui/react-components';
import {
  Edit24Regular,
  Delete24Regular,
  Document24Regular,
  CalendarLtr24Regular,
  Money24Regular,
  Info24Regular,
} from '@fluentui/react-icons';

import { Project } from '../../services/ProjectSelectionServices';
import ProjectOverview from './ProjectsOverview';
import ProjectTeam from './ProjectTeam';
import ProgrammeTab from '../projectCalender/ProgrammeTab';

import DocumentsTab from '../projectDocuments/DocumentsTab';
import StagesTab from '../projectStages/StagesTab';
import FeesTab from '../projectFees/FeesTab';
import { WebPartContext } from '@microsoft/sp-webpart-base';

const useStyles = makeStyles({
  root: {
    display: 'flex',
    flexDirection: 'column',
    rowGap: '24px',
  },
  layout: {
    display: 'flex',
    flexDirection: 'row',
    columnGap: '24px',
    alignItems: 'flex-start',
  },
  leftColumn: {
    flexGrow: 1,
  },
  buttonRow: {
    display: 'flex',
    justifyContent: 'center',
    gap: '16px',
    marginTop: '24px',
  },
});

interface ProjectTabsProps {
  project: Project;
  context: WebPartContext;
  onEdit: () => void;
  onDelete: () => void;
  onTabChange?: (tab: TabValue) => void;
}

const ProjectTabs: React.FC<ProjectTabsProps> = ({ context, project, onEdit, onDelete, onTabChange }) => {
  const styles = useStyles();
  const [selectedValue, setSelectedValue] = React.useState<TabValue>('overview');

  const renderTabContent = (): JSX.Element | null => {
    switch (selectedValue) {
      case 'overview':
        return (
          <>
            <div className={styles.layout}>
              <div className={styles.leftColumn}>
                <Text size={700} weight="bold">
                  {project.ProjectNumber} - {project.ProjectName}
                </Text>
                <Divider />
                <ProjectOverview project={project} />
              </div>
              <Divider vertical />
              <ProjectTeam project={project} />
            </div>
            <div className={styles.buttonRow}>
              <Button icon={<Edit24Regular />} onClick={onEdit} appearance="primary">
                Edit
              </Button>
              <Button icon={<Delete24Regular />} onClick={onDelete} appearance="secondary">
                Delete
              </Button>
            </div>
          </>
        );
      case 'programme':
        return <ProgrammeTab project={project} context={context} />;
      case 'stages':
        return <StagesTab project={project} context={context} />;
      case 'documents':
        return <DocumentsTab project={project} context={context} />;
      case 'fees':
        return <FeesTab project={project} context={context} />;
      default:
        return null;
    }
  };

  return (
    <div className={styles.root}>
      <TabList
        selectedValue={selectedValue}
        onTabSelect={(e, data) => {
          setSelectedValue(data.value);
          onTabChange?.(data.value); // <-- notify parent
        }}

        size="large"
      >
        <Tab icon={<Info24Regular />} value="overview">
          Overview
        </Tab>
        <Tab icon={<CalendarLtr24Regular />} value="programme">
          Programme
        </Tab>
        <Tab icon={<Edit24Regular />} value="stages">
          Stages
        </Tab>
        <Tab icon={<Document24Regular />} value="documents">
          Documents
        </Tab>
        <Tab icon={<Money24Regular />} value="fees">
          Fees
        </Tab>
      </TabList>
      {renderTabContent()}
    </div>
  );
};

export default ProjectTabs;
