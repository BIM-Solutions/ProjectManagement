

export interface IUserField {
  Id: number;
  Title: string;
  Email: string;
  JobTitle?: string;
  Department?: string;
  Presence?: string;
}
export interface Project {
  id: number;
  ProjectNumber: string;
  ProjectName: string;
  Status: string;
  Client: string;
  ClientContact?: string; 
  ProjectDescription?: string;
  DeltekSubCodes?: string;
  SubCodes?: string;
  Sector: string;
  PM?: IUserField;
  Manager?: IUserField;
  Checker?: IUserField;
  Approver?: IUserField;
  ProjectImage?: string;
}


  
  type Listener = (project: Project | undefined) => void;
  
  declare global {
    interface Window {
      __sharedProjectSelection: {
        selectedProject?: Project;
        listeners: Listener[];
      };
    }
  }
  
  if (!window.__sharedProjectSelection) {
    window.__sharedProjectSelection = {
      selectedProject: undefined,
      listeners: [],
    };
  }
  
  export class ProjectSelectionService {
    public static setSelectedProject(project: Project | undefined): void {
      // console.log("Setting selected project (global):", project);
      window.__sharedProjectSelection.selectedProject = project;
      window.__sharedProjectSelection.listeners.forEach(l => l(project));
    }
  
    public static getSelectedProject(): Project | undefined {
      return window.__sharedProjectSelection.selectedProject;
    }
  
    public static subscribe(listener: Listener): void {
      window.__sharedProjectSelection.listeners.push(listener);
    }
  
    public static unsubscribe(listener: Listener): void {
      window.__sharedProjectSelection.listeners =
        window.__sharedProjectSelection.listeners.filter(l => l !== listener);
    }
  }
  