import * as React from 'react';
import { useState, useEffect } from 'react';
import {
  Dialog,
  Button,
  DialogTrigger,
  DialogSurface,
  DialogBody,
  DialogContent,
  DialogActions,
  makeStyles,
  tokens,
  Tooltip,
} from '@fluentui/react-components';
import {
  AddRegular,
  PersonCircle32Regular,
  Board20Filled,
  Board20Regular,
  bundleIcon,
} from '@fluentui/react-icons';
import {
  AppItem,
  Hamburger,
  NavDrawer,
  NavDrawerBody,
  NavDrawerHeader,
  NavItem,
} from '@fluentui/react-nav-preview';
import ProjectList from './ProjectList';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { useLoading } from '../services/LoadingContext';
import ProjectForm from './common/ProjectForm';
import { ProjectSelectionService } from '../services/ProjectSelectionServices';
import { Project } from '../services/ProjectSelectionServices';
import ProjectDetails from './ProjectDetails';
import ProgrammeView from './projectCalender/ProgrammeView';


const useStyles = makeStyles({
  root: {
    overflow: 'hidden',
    display: 'flex',
    flex: 1,
    minHeight: 0, 
  },
  nav: {
    minWidth: '200px',
  },
  content: {
    flex: '1',
    padding: '16px',
    display: 'grid',
    justifyContent: 'flex-start',
    alignItems: 'flex-start',
  },
  field: {
    display: 'flex',
    marginTop: '4px',
    marginLeft: '8px',
    flexDirection: 'column',
    gridRowGap: tokens.spacingVerticalS,
  },
});

interface ILandingPageProps {
  context: WebPartContext;
}

interface CustomWindow {
  __setListLoading?: (isLoading: boolean) => void;
}

const LandingPage: React.FC<ILandingPageProps> = ({ context }) => {
  const { setIsLoading } = useLoading();
  const [selectedProject, setSelectedProject] = useState<Project | undefined>();
  const [isOpen, setIsOpen] = React.useState(true);
  const Dashboard = bundleIcon(Board20Filled, Board20Regular);
  const styles = useStyles();
  const [tab, setTab] = useState<string>('overview');

  const linkDestination = 'https://contoso.sharepoint.com/sites/ContosoHR/Pages/Home.aspx';

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
    return () => {
      delete (window as CustomWindow).__setListLoading;
    };
  }, [setIsLoading]);

  return (
    <div className={styles.root}>
      <div
        style={{
          display: 'flex',
          justifyContent: 'top',
          alignItems: 'top',
          flexWrap: 'wrap',
          flexDirection: 'row',
          width: '100%',
          height: '100vh',
          boxSizing: 'border-box',
          padding: '20px',
          overflowY: 'auto',
          overflowX: 'auto',
        }}
      >
        {/* Left Panel - Navigation Drawer */}
        {isOpen && (
          <NavDrawer
            defaultSelectedValue="1"
            defaultSelectedCategoryValue=""
            open={isOpen}
            type="inline"
            multiple={true}
            className={styles.nav}
            style={{ height: '100%', minWidth: '200px' }}
          >
            <NavDrawerHeader>
              <Tooltip content="Close Navigation" relationship="label">
                <Hamburger onClick={() => setIsOpen(!isOpen)} />
              </Tooltip>
            </NavDrawerHeader>

            <NavDrawerBody>
              <AppItem icon={<PersonCircle32Regular />} as="a" href={linkDestination}>
                BIM Team
              </AppItem>
              <NavItem href={linkDestination} icon={<Dashboard />} value="1">
                Projects
              </NavItem>
              <NavItem href={linkDestination} icon={<Dashboard />} value="2">
                Resourcing
              </NavItem>
              <NavItem href={linkDestination} icon={<Dashboard />} value="3">
                CoFW
              </NavItem>
            </NavDrawerBody>
          </NavDrawer>
        )}
        {!isOpen && (
          <Tooltip content="Open Navigation" relationship="label">
            <Hamburger style={{ height: '32px' }} onClick={() => setIsOpen(!isOpen)} />
          </Tooltip>
        )}

        {/* Center Panel */}
        <div
          style={{
            flex: '1 1 400px',
            height: '100%',
            overflowY: 'auto',
            minWidth: '300px',
            maxWidth: '100%',
            boxSizing: 'border-box',
            padding: '16px',
          }}
        >
          <Dialog modalType="modal">
            <DialogTrigger disableButtonEnhancement>
              <Button
                appearance="primary"
                icon={<AddRegular />}
                iconPosition="before"
                data-trigger="AddRegular"
                style={{ marginBottom: 10, width: '250px' }}
              >
                Add Project
              </Button>
            </DialogTrigger>
            <DialogSurface style={{ width: '100%' }} aria-hidden="true">
              <DialogBody>
                <DialogContent>
                  <ProjectForm
                    onSuccess={() => {
                      const dialogTrigger: HTMLElement | null = document.querySelector('[data-trigger]');
                      if (dialogTrigger) dialogTrigger.click();
                    }}
                    context={context}
                    mode="create"
                  />
                </DialogContent>
                <DialogActions>
                  <DialogTrigger disableButtonEnhancement>
                    <Button appearance="secondary">Close</Button>
                  </DialogTrigger>
                </DialogActions>
              </DialogBody>
            </DialogSurface>
          </Dialog>

          {tab === 'programme' ? <ProgrammeView context={context} project={selectedProject!} /> : <ProjectList />}
        </div>

        {/* Right Panel - Project Details */}
        <div
          style={{
            flex: '0 0 30%',
            height: '100%',
            overflowY: 'auto',
            minWidth: '250px',
            boxSizing: 'border-box',
            padding: '16px',
          }}
        >
          {selectedProject && (
            <ProjectDetails
              context={context}
              onTabChange={(newTab) => setTab(newTab)}
            />
          )}
        </div>
      </div>
    </div>
  );
};

export default LandingPage;
