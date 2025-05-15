import * as React from 'react';
import { useState, useEffect } from 'react';
import {
  Tab,
  TabValue,
  SelectTabEvent,
  SelectTabData,
  makeStyles,
  tokens,
  Text,
  TabList,
} from '@fluentui/react-components';
// import {
//   TabGroup,
// } from '@fluentui/react-tabs';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { ProjectService, IProject } from '../services/ProjectService';
import { TemplateService } from '../services/TemplateService';
import { DocumentService } from '../services/DocumentService';
import { IntegrationService } from '../services/IntegrationService';

// Import your tab components here
import ProjectInfoTab from './projectInfo/ProjectInfoTab';
import DocumentsTab from './documents/DocumentsTab';
import TasksTab from './tasks/TasksTab';
import TemplatesTab from './templates/TemplatesTab';
import CalendarTab from './calendar/CalendarTab';
import DashboardTab from './dashboard/DashboardTab';

export interface IProjectManagementProps {
  context: WebPartContext;
}

const useStyles = makeStyles({
  container: {
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalM,
    padding: tokens.spacingHorizontalM,
  },
  header: {
    display: 'flex',
    justifyContent: 'space-between',
    alignItems: 'center',
    padding: tokens.spacingVerticalS,
  },
  content: {
    backgroundColor: tokens.colorNeutralBackground1,
    borderRadius: tokens.borderRadiusMedium,
    padding: tokens.spacingHorizontalM,
  },
});

export const ProjectManagement: React.FC<IProjectManagementProps> = ({ context }) => {
  const styles = useStyles();
  const [selectedProject, setSelectedProject] = useState<IProject | undefined>(undefined);
  const [selectedTab, setSelectedTab] = useState<TabValue>(0);

  // Initialize services
  const projectService = new ProjectService(context);
  const templateService = new TemplateService(context);
  const documentService = new DocumentService(context);
  const integrationService = new IntegrationService(context);

  useEffect(() => {
    // Load initial project data if needed
    const loadInitialProject = async (): Promise<void> => {
      try {
        const projects = await projectService.getAllProjects();
        if (projects.length > 0) {
          setSelectedProject(projects[0]);
        }
      } catch (error) {
        console.error('Error loading initial project:', error);
      }
    };

 
    loadInitialProject().catch(console.error);
    
  }, []);

  const handleProjectChange = async (project: IProject): Promise<void> => {
    setSelectedProject(project);
  };

  const handleTabSelect = (_: SelectTabEvent, data: SelectTabData): void => {
    setSelectedTab(data.value);
  };

  return (
    <div className={styles.container}>
      <div className={styles.header}>
        <Text size={800} weight="semibold">
          Project Management
        </Text>
      </div>

      <div className={styles.content}>
        <TabList selectedValue={selectedTab} onTabSelect={handleTabSelect}>
          <Tab value={0}>Project Info</Tab>
          <Tab value={1}>Documents</Tab>
          <Tab value={2}>Tasks</Tab>
          <Tab value={3}>Templates</Tab>
          <Tab value={4}>Calendar</Tab>
          <Tab value={5}>Dashboard</Tab>
        </TabList>

        {selectedTab === 0 && (
          <ProjectInfoTab
            context={context}
            project={selectedProject}
            onProjectChange={handleProjectChange}
            projectService={projectService}
          />
        )}
        {selectedTab === 1 && (
          <DocumentsTab
            context={context}
            project={selectedProject}
            documentService={documentService}
            templateService={templateService}
          />
        )}
        {selectedTab === 2 && (
          <TasksTab
            context={context}
            project={selectedProject}
          />
        )}
        {selectedTab === 3 && (
          <TemplatesTab
            context={context}
            project={selectedProject}
            templateService={templateService}
          />
        )}
        {selectedTab === 4 && (
          <CalendarTab
            context={context}
            project={selectedProject}
          />
        )}
        {selectedTab === 5 && (
          <DashboardTab
            context={context}
            project={selectedProject}
            integrationService={integrationService}
          />
        )}
      </div>
    </div>
  );
}; 