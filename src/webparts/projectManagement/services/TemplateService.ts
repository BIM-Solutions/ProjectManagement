import { spfi, SPFx } from "@pnp/sp/presets/all";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SPFI } from "@pnp/sp";

export interface ITemplate {
  Id: number;
  Title: string;
  Version: string;
  Description: string;
  TemplateUrl: string;
  Category: string;
  IsActive: boolean;
  CreatedDate: string;
  ModifiedDate: string;
}

export class TemplateService {
  private sp: SPFI;
  private readonly templateListName = "ProjectTemplates";
//   private readonly standardsLibraryName = "StandardsLibrary";

  constructor(context: WebPartContext) {
    this.sp = spfi().using(SPFx(context));
  }

  public async getTemplates(): Promise<ITemplate[]> {
    return await this.sp.web.lists
      .getByTitle(this.templateListName)
      .items.select(
        "Id",
        "Title",
        "Version",
        "Description",
        "TemplateUrl",
        "Category",
        "IsActive",
        "Created",
        "Modified"
      )();
  }

  public async getTemplateById(id: number): Promise<ITemplate> {
    return await this.sp.web.lists
      .getByTitle(this.templateListName)
      .items.getById(id)();
  }

  public async createTemplate(template: Partial<ITemplate>): Promise<ITemplate> {
    const result = await this.sp.web.lists
      .getByTitle(this.templateListName)
      .items.add(template);
    return result.data as ITemplate;
  }

  public async copyTemplateToProject(templateId: number, projectFolder: string): Promise<void> {
    const template = await this.getTemplateById(templateId);
    const sourceFile = template.TemplateUrl;
    const destinationUrl = `${projectFolder}/Standards/${template.Title}`;
    
    await this.sp.web.getFileByUrl(sourceFile)
      .copyTo(destinationUrl, true);
  }

  public async getActiveStandardsVersion(): Promise<string> {
    const versions = await this.sp.web.lists
      .getByTitle(this.templateListName)
      .items.filter("IsActive eq true")
      .orderBy("Version", false)
      .top(1)();
    
    return versions.length > 0 ? versions[0].Version : "1.0";
  }
} 