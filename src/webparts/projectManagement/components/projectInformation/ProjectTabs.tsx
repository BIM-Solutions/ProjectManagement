import * as React from 'react';
import {
  makeStyles,
  Tab,
  TabList,
  TabValue,
  Text,
  Button,
  Divider,
  tokens
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
import { TaskItem } from '../projectCalender/ProgrammeTab';

const useStyles = makeStyles({
  root: {
    display: 'flex',
    flexDirection: 'column',
    rowGap: '24px',
    margin: '24px'
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
  rightPanel: {
    width: '40%',
    padding: '16px',
    backgroundColor: tokens.colorNeutralBackground2,
    border: `1px solid ${tokens.colorNeutralStroke1}`,
    borderRadius: '8px',
  },
  verticalDivider: {
    height: 'auto',
    alignSelf: 'stretch',
    borderLeft: `1px solid ${tokens.colorNeutralStroke2}`,
    margin: '0 12px',
  },
});

interface ProjectTabsProps {
  project: Project;
  context: WebPartContext;
  onEdit: () => void;
  onDelete: () => void;
  onTabChange?: (tab: TabValue) => void;
  tasks: TaskItem[];
  setTasks: React.Dispatch<React.SetStateAction<TaskItem[]>>;
  selectedTask: TaskItem | undefined;
  setSelectedTask: (task: TaskItem | undefined) => void;
}

const ProjectTabs: React.FC<ProjectTabsProps> = ({ context, project, onEdit, onDelete, onTabChange, tasks, setTasks, selectedTask, setSelectedTask }) => {
  const styles = useStyles();
  const [selectedValue, setSelectedValue] = React.useState<TabValue>('overview');

  const renderTabContent = (): JSX.Element | null => {
    switch (selectedValue) {
      case 'overview':
        return (
          <>
            <div className={styles.layout}>
              <div className={styles.leftColumn}>
                <Text size={600} weight="bold" style={{margin: '12px'}}>
                  {project.ProjectNumber} - {project.ProjectName}
                </Text>
                <Divider />
                <ProjectOverview project={project} />
              </div>
              <div className={styles.verticalDivider} />
              <div className={styles.rightPanel}>
                <ProjectTeam project={project} context={context} />
              </div>

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
        return <ProgrammeTab project={project} context={context}  tasks={tasks} setSelectedTask={setSelectedTask} setTasks={setTasks} selectedTask={selectedTask}/>;
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
