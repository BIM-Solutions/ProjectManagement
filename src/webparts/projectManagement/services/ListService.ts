import {  SPFI } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/fields";
import { IList } from "@pnp/sp/lists";
// import { ChoiceFieldFormatType } from "@pnp/sp/fields/types";
import {
  AddTextProps,
  AddChoiceProps,
  AddNumberProps,
  AddDateTimeProps,
  
} from '@pnp/sp/fields/types';

export class ListService {
  private sp: SPFI;

  constructor(sp: SPFI) {
    this.sp = sp;
  }

  public async ensureListSchema(): Promise<void> {
    await this.ensureProjectInformation();
    await this.ensureProjectTasks();
    await this.ensureProjectFees();
    await this.ensureProjectDocuments();
    await this.ensureChangeControl();
    await this.ensureProjectStages();
  }

  private async ensureTextField(list: IList, title: string, settings: AddTextProps = {}): Promise<void> {
    try {
      await list.fields.getByTitle(title)();
    } catch {
      await list.fields.addText(title, settings);
      await list.defaultView.fields.add(title);
    }
  }
  
  private async ensureChoiceField(list: IList, title: string, settings: AddChoiceProps): Promise<void> {
    try {
      await list.fields.getByTitle(title)();
    } catch {
      const { Choices, ...rest } = settings;
      await list.fields.addChoice(title, {
        Choices: Choices ?? [], // Ensure it's always a string[]
        EditFormat: 0, // Dropdown
        FillInChoice: false,
        ...rest,
      });
      await list.defaultView.fields.add(title);
    }
  }
  
  private async ensureNumberField(list: IList, title: string, settings: AddNumberProps = {}): Promise<void> {
    try {
      await list.fields.getByTitle(title)();
    } catch {
      await list.fields.addNumber(title, settings);
      await list.defaultView.fields.add(title);
    }
  }
  
  private async ensureDateTimeField(list: IList, title: string, settings: AddDateTimeProps = {}): Promise<void> {
    try {
      await list.fields.getByTitle(title)();
    } catch {
      await list.fields.addDateTime(title, settings);
      await list.defaultView.fields.add(title);
    }
  }
  
  

  private async ensureProjectInformation(): Promise<void> {
    const { list } = await this.sp.web.lists.ensure("ProjectInformation", "Stores project metadata");
    await this.ensureTextField(list, "ProjectNumber");
    await this.ensureChoiceField(list, "Sector",  { Choices: ["Commercial", "Residential", "Infrastructure"] });
    await this.ensureTextField(list, "KeyPersonnel");
  }

  private async ensureProjectTasks(): Promise<void> {
    const { list } = await this.sp.web.lists.ensure("ProjectTasks", "Stores project tasks");
    await this.ensureTextField(list, "TaskName");
    await this.ensureChoiceField(list, "Status",  { Choices: ["Not Started", "In Progress", "Complete"]});
    await this.ensureTextField(list, "DueDate");
  }

  private async ensureProjectFees(): Promise<void> {
    const { list } = await this.sp.web.lists.ensure("ProjectFees", "Stores budget and fees");
    await this.ensureTextField(list, "FeeName");
    await this.ensureNumberField(list, "Amount");
    await this.ensureChoiceField(list, "Stage",  { Choices: ["Concept", "Design", "Construction"] });
  }

  private async ensureProjectDocuments(): Promise<void> {
    const { list } = await this.sp.web.lists.ensure("ProjectDocuments", "Stores document metadata");
    await this.ensureTextField(list, "DocumentTitle");
    await this.ensureChoiceField(list, "Category",  { Choices: ["Drawing", "Specification", "Report"]});
  }

  private async ensureChangeControl(): Promise<void> {
    const { list } = await this.sp.web.lists.ensure("ChangeControl", "Tracks changes");
    await this.ensureTextField(list, "ChangeTitle");
    await this.ensureChoiceField(list, "ChangeType",  { Choices: ["Scope", "Budget", "Timeline"] });
    await this.ensureChoiceField(list, "Approved",  { Choices: ["Yes", "No", "Pending"] });
  }

  private async ensureProjectStages(): Promise<void> {
    const { list } = await this.sp.web.lists.ensure("ProjectStages", "Stages of project");
    await this.ensureTextField(list, "StageName");
    await this.ensureDateTimeField(list, "StartDate");
    await this.ensureDateTimeField(list, "EndDate");
  }
}
