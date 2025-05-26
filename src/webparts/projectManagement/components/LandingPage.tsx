import * as React from "react";
import { spfi, SPFx } from "@pnp/sp/presets/all";
import { useState, useEffect } from "react";
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
} from "@fluentui/react-components";
import { AddRegular } from "@fluentui/react-icons";

import ProjectList from "./ProjectList";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { useLoading } from "../services/LoadingContext";
import ProjectForm from "./common/ProjectForm";
import { ProjectSelectionService } from "../services/ProjectSelectionServices";
import { Project } from "../services/ProjectSelectionServices";
import ProjectDetails from "./ProjectDetails";
import TaskCalendar from "./projectCalender/TaskCalendar";
import Navigation from "../../common/components/Navigation";
import { TaskItem } from "./projectCalender/ProgrammeTab";
// import StagesTab from './projectStages/StagesTab';
import DocumentsOverview from "./projectDocuments/DocumentsOverview";
// import FeesTab from './projectFees/FeesTab';
import FeeOverview from "./projectFees/FeeOverview";
import StageOverview from "./projectStages/StageOverview";

import { DocumentService } from "../services/DocumentService";
import { TemplateService } from "../services/TemplateService";

const useStyles = makeStyles({
  root: {
    overflow: "hidden",
    display: "flex",
    flex: 1,
    minHeight: 0,
    backgroundColor: tokens.colorNeutralBackground1,
    color: tokens.colorNeutralForeground1,
  },
  nav: {
    height: "100vh",
    borderRight: `1px solid ${tokens.colorNeutralStroke1}`,
    padding: "12px",
    backgroundColor: tokens.colorNeutralBackground2,
  },
  content: {
    flex: "1 1 0",
    minWidth: "400px", // Prevents center collapse
    padding: "24px",
    display: "flex",
    flexDirection: "column",
    gap: "16px",
    overflowX: "auto",
    height: "100vh",
  },
  rightPanel: {
    flex: "0 0 30%",
    minWidth: "280px",
    maxWidth: "30%",
    height: "100vh",
    overflowY: "auto",
    padding: "20px",
    borderLeft: `1px solid ${tokens.colorNeutralStroke1}`,
    backgroundColor: tokens.colorNeutralBackground2,
    //remove this for production
    "@media (max-width: 2600px)": {
      maxWidth: "100%",
    },
  },
  field: {
    display: "flex",
    marginTop: "4px",
    marginLeft: "8px",
    flexDirection: "column",
    gridRowGap: tokens.spacingVerticalS,
  },
});

interface ILandingPageProps {
  context: WebPartContext;
  project: Project;
}

interface CustomWindow {
  __setListLoading?: (isLoading: boolean) => void;
}

const LandingPage: React.FC<ILandingPageProps> = ({ context, project }) => {
  const { setIsLoading } = useLoading();
  const [selectedProject, setSelectedProject] = useState<Project | undefined>();
  const [documentService] = useState(() => new DocumentService(context));
  const [templateService] = useState(() => new TemplateService(context));
  const styles = useStyles();
  const [tab, setTab] = useState<string>("overview");
  const [tasks, setTasks] = useState<TaskItem[]>([]);
  const [selectedTask, setSelectedTask] = useState<TaskItem | undefined>();
  const [selectedStageId, setSelectedStageId] = useState<number>();
  const [stagesChanged, setStagesChanged] = useState(0);

  const sp = spfi().using(SPFx(context));

  const renderTabContent = (): JSX.Element | null => {
    switch (tab) {
      case "overview":
        return <ProjectList />;
      case "programme":
        return <TaskCalendar tasks={tasks} onTaskClick={setSelectedTask} />;
      case "stages":
        return (
          <StageOverview
            project={selectedProject}
            context={context}
            onStageSelect={setSelectedStageId}
            stagesChanged={stagesChanged}
          />
        );
      case "documents":
        return (
          <DocumentsOverview
            sp={sp}
            project={selectedProject}
            context={context}
            documentService={documentService}
            templateService={templateService}
          />
        );
      case "fees":
        return <FeeOverview />;
      default:
        return null;
    }
  };

  const renderRightPanel = (): JSX.Element | null => {
    if (selectedProject) {
      return (
        <>
          <ProjectDetails
            context={context}
            onTabChange={(newTab) => setTab(newTab)}
            tasks={tasks}
            setTasks={setTasks}
            selectedTask={selectedTask}
            setSelectedTask={setSelectedTask}
            selectedStageId={selectedStageId}
            stagesChanged={() => setStagesChanged(stagesChanged + 1)}
          />
        </>
      );
    }
    return null;
  };

  useEffect(() => {
    const service = ProjectSelectionService;
    const listener = (selected: Project | undefined): void => {
      if (selected) {
        setSelectedProject(selected);
      }
    };
    service.subscribe(listener);
    // Initialize with current selection
    const currentProject = service.getSelectedProject();
    if (currentProject) {
      setSelectedProject(currentProject);
    }
    return () => {
      service.unsubscribe(listener);
    };
  }, []);

  // Add effect to maintain project selection when tab changes
  useEffect(() => {
    const currentProject = ProjectSelectionService.getSelectedProject();
    if (currentProject && !selectedProject) {
      setSelectedProject(currentProject);
    }
  }, [tab, selectedProject]);

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
          display: "flex",
          justifyContent: "top",
          alignItems: "top",
          flexWrap: "wrap",
          flexDirection: "row",
          width: "100%",
          height: "100vh",
          boxSizing: "border-box",
          padding: "20px",
          overflowY: "auto",
          overflowX: "auto",
        }}
      >
        {/* Left Panel - Navigation Drawer */}
        <nav className={styles.nav}>
          <Navigation context={context} />
        </nav>

        {/* Center Panel */}
        <div className={styles.content}>
          <Dialog modalType="modal">
            <DialogTrigger disableButtonEnhancement>
              <div style={{ marginBottom: "24px" }}>
                {tab === "overview" && (
                  <Button
                    appearance="primary"
                    icon={<AddRegular />}
                    iconPosition="before"
                    data-trigger="AddRegular"
                    style={{ width: "250px", fontWeight: "600" }}
                  >
                    Add Project
                  </Button>
                )}
              </div>
            </DialogTrigger>
            <DialogSurface style={{ width: "100%" }} aria-hidden="true">
              <DialogBody>
                <DialogContent>
                  <ProjectForm
                    onSuccess={() => {
                      const dialogTrigger: HTMLElement | null =
                        document.querySelector("[data-trigger]");
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
          {renderTabContent()}
        </div>

        {/* Right Panel - Project Details */}
        <div className={styles.rightPanel}>{renderRightPanel()}</div>
      </div>
    </div>
  );
};

export default LandingPage;
