import { spfi, SPFx } from "@pnp/sp/presets/all";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SPFI } from "@pnp/sp";

export interface IProject {
  Id: number;
  Title: string;
  ProjectNumber: string;
  ProjectManager: string;
  StartDate: string;
  EndDate: string;
  Status: string;
  Description: string;
  ClientName: string;
  Budget: number;
  TemplateVersion: string;
}

export class ProjectService {
  private sp: SPFI;
  private readonly projectListName = "Projects";

  constructor(context: WebPartContext) {
    this.sp = spfi().using(SPFx(context));
  }

  public async createProject(project: Partial<IProject>): Promise<IProject> {
    const result = await this.sp.web.lists
      .getByTitle(this.projectListName)
      .items.add(project);
    return result.data as IProject;
  }

  public async getProject(projectId: number): Promise<IProject> {
    return await this.sp.web.lists
      .getByTitle(this.projectListName)
      .items.getById(projectId)();
  }

  public async updateProject(projectId: number, updates: Partial<IProject>): Promise<void> {
    await this.sp.web.lists
      .getByTitle(this.projectListName)
      .items.getById(projectId)
      .update(updates);
  }

  public async getAllProjects(): Promise<IProject[]> {
    return await this.sp.web.lists
      .getByTitle(this.projectListName)
      .items.select(
        "Id",
        "Title",
        "ProjectNumber",
        "ProjectManager",
        "StartDate",
        "EndDate",
        "Status",
        "Description",
        "ClientName",
        "Budget",
        "TemplateVersion"
      )();
  }
} 