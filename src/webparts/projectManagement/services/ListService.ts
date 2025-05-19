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
  AddMultilineTextProps,
  AddUserProps,
  AddCurrencyProps
} from '@pnp/sp/fields/types';
import { DropdownProps } from "@fluentui/react-components";
export interface ComboboxOption {
  key: string;
  value: string;
}

export const taskStatusOptions: DropdownProps[] = [
  { key: 'Not Started', value: 'Not Started' },
  { key: 'In Progress', value: 'In Progress' },
  { key: 'Complete', value: 'Complete' },
  { key: 'Cancelled', value: 'Cancelled' },
  { key: 'On Hold', value: 'On Hold' },
  { key: 'Postponed', value: 'Postponed' }
];

export const taskPriorityOptions: DropdownProps[] = [
  { key: 'High', value: 'High' },
  { key: 'Medium', value: 'Medium' },
  { key: 'Low', value: 'Low' },
  { key: 'Critical', value: 'Critical' },
  { key: 'Urgent', value: 'Urgent' }
];

export const taskTypeOptions: DropdownProps[] = [
  { key: 'Task', value: 'Task' },
  { key: 'Meeting', value: 'Meeting' },
  { key: 'DocumentWorkflow', value: 'Document Workflow' },
  { key: 'Reporting', value: 'Reporting' },
]

export const changeTypeOptions: DropdownProps[] = [
  { key: 'Scope', value: 'Scope' },
  { key: 'Schedule', value: 'Schedule' },
  { key: 'Cost', value: 'Cost' },
  { key: 'Risk', value: 'Risk' },
]

export const changeStatusOptions: DropdownProps[] = [
  { key: 'Approved', value: 'Approved' },
  { key: 'Pending', value: 'Pending' },
  { key: 'Rejected', value: 'Rejected' },
]


export const sectorOptions: ComboboxOption[] = [
  { key: 'Public', value: 'Public' },
  { key: 'Defence', value: 'Defence' },
  { key: 'Power', value: 'Power' }
];

export const projectStatusOptions: ComboboxOption[] = [
  { key: 'Enquiry', value: 'Enquiry' },
  { key: 'Active', value: 'Active' },
  { key: 'Complete', value: 'Complete' },
  { key: 'Lost', value: 'Lost' },
  { key: 'Cancelled', value: 'Cancelled' },
  { key: 'Inactive', value: 'Inactive' }
];

export const projectStageStatusOptions: ComboboxOption[] = [
  { key: 'Active', value: 'Active' },
  { key: 'Completed', value: 'Completed' },
  { key: 'Cancelled', value: 'Cancelled' }
];


export class ListService {
  private sp: SPFI;

  /**
   * Creates a new instance of the ListService class.
   * @param sp The PnPJS SPFI instance to use for interacting with the SharePoint REST API.
   */
  constructor(sp: SPFI) {
    this.sp = sp;
  }

  /**
   * Ensures that all the required lists and their fields are provisioned on the site.
   * This method is called when the web part is initialized to ensure that the necessary lists are available.
   * It calls individual methods to ensure each list and its fields are created or updated as needed.
   * The lists include:
   * - ProjectInformationDatabase
   * - ProjectTasks
   * - ProjectFees
   * - ProjectDocuments
   * - ChangeControl
   * - ProjectStages
   */
  public async ensureFirstListSchema(): Promise<boolean> {
    const listCreated = await this.ensureProjectInformationExists();

    return listCreated;
  }

  /**
   * Ensures that all the required lists and their fields are provisioned on the site.
   * @returns A promise that resolves when all the lists have been provisioned.
   */
  public async ensureListSchema(): Promise<void> {
    await this.ensureProjectInformation();
    await this.ensureProjectTasks();
    await this.ensureProjectFees();
    await this.ensureProjectDocuments();
    await this.ensureChangeControl();
    await this.ensureProjectStages();
  }

  /**
   * Ensures that a text field exists on the given list and is configured according to the given settings.
   * @param list The list to check for the field.
   * @param title The title of the field to check for.
   * @param settings The settings to use when creating the field if it does not already exist.
   * @returns A promise that resolves when the field has been ensured.
   */
  private async ensureTextField(list: IList, title: string, settings: AddTextProps = {}): Promise<void> {
    try {
      await list.fields.getByTitle(title)();
    } catch {
      await list.fields.addText(title, settings);
      await list.defaultView.fields.add(title);
    }
  }
  
  /**
   * Ensures that a choice field exists on the given list and is configured according to the given settings.
   * @param list The list to check for the field.
   * @param title The title of the field to check for.
   * @param settings The settings to use when creating the field if it does not already exist.
   * @returns A promise that resolves when the field has been ensured.
   */
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
  
  /**
   * Ensures that a number field exists on the given list and is configured according to the given settings.
   * @param list The list to check for the field.
   * @param title The title of the field to check for.
   * @param settings The settings to use when creating the field if it does not already exist.
   * @returns A promise that resolves when the field has been ensured.
   */
  private async ensureNumberField(list: IList, title: string, settings: AddNumberProps = {}): Promise<void> {
    try {
      await list.fields.getByTitle(title)();
    } catch {
      await list.fields.addNumber(title, settings);
      await list.defaultView.fields.add(title);
    }
  }
  
  /**
   * Ensures that a currency field exists on the given list and is configured according to the given settings.
   * @param list The list to check for the field.
   * @param title The title of the field to check for.
   * @param settings The settings to use when creating the field if it does not already exist.
   * @returns A promise that resolves when the field has been ensured.
   */
  private async ensureCurrencyField(list: IList, title: string, settings: AddCurrencyProps = {}): Promise<void> {
    try {
      await list.fields.getByTitle(title)();
    } catch {
      await list.fields.addCurrency(title, settings);
      await list.defaultView.fields.add(title);
    }
  }

  /**
   * Ensures that a multi-line text field exists on the given list and is configured according to the given settings.
   * @param list The list to check for the field.
   * @param title The title of the field to check for.
   * @param settings The settings to use when creating the field if it does not already exist.
   * @returns A promise that resolves when the field has been ensured.
   */
  private async ensureMultiLineField(list: IList, title: string, settings: AddMultilineTextProps = {}): Promise<void> {
    try {
      await list.fields.getByTitle(title)();
    } catch {
      await list.fields.addMultilineText(title, settings);
      await list.defaultView.fields.add(title);
    }
  }
  /**
   * Ensures that a datetime field exists on the given list and is configured according to the given settings.
   * @param list The list to check for the field.
   * @param title The title of the field to check for.
   * @param settings The settings to use when creating the field if it does not already exist.
   * @returns A promise that resolves when the field has been ensured.
   */
  private async ensureDateTimeField(list: IList, title: string, settings: AddDateTimeProps = {}): Promise<void> {
    try {
      await list.fields.getByTitle(title)();
    } catch {
      await list.fields.addDateTime(title, settings);
      await list.defaultView.fields.add(title);
    }
  }

  /**
   * Ensures that a user field exists on the given list and is configured according to the given settings.
   * @param list The list to check for the field.
   * @param title The title of the field to check for.
   * @param settings The settings to use when creating the field if it does not already exist.
   * @returns A promise that resolves when the field has been ensured.
   */
  private async ensureUserField(list: IList, title: string, settings: AddUserProps = {}): Promise<void> {
    try {
      await list.fields.getByTitle(title)();
    } catch {
      await list.fields.addUser(title, settings);
      await list.defaultView.fields.add(title);
    }
  }

  private async ensureImageField(list: IList, title: string): Promise<void> {
    try {
      await list.fields.getByTitle(title)();
    } catch {
      await list.fields.addImageField(title);
      await list.defaultView.fields.add(title);
    }
  }

  /**
   * Checks if the "ProjectInformationDatabase" list exists on the site.
   * If the list does not exist, it will not be created.
   * @returns A promise that resolves with a boolean indicating whether the list exists.
   */
  private async ensureProjectInformationExists(): Promise<boolean> {
    let listCreated = false;
    try {
      await this.sp.web.lists.getByTitle("9719_ProjectInformationDatabase")();
      listCreated = true;
    } catch (error) {
      console.error("List does not exist:", error);
      listCreated = false;
    }
    return listCreated;
  }

/**
 * Ensures that the "ProjectInformationDatabase" list exists and is configured with the required fields.
 * Fields include text, multi-line, choice, and user fields to store project metadata.
 * It configures:
 * - ProjectName: Text field for the project name.
 * - ProjectDescription: Multi-line text field for the project description.
 * - DeltekSubCodes: Multi-line text field for sub-codes related to Deltek.
 * - Client: Text field for the client name.
 * - Manager, Checker, Approver: User fields for respective personnel.
 * - Status: Choice field with options for project status.
 * - Sector: Choice field with predefined sectors.
 * - ClientContact: Multi-line text field for client contact details.
 * 
 * @returns A promise that resolves when the list and fields have been ensured.
 */
  private async ensureProjectInformation(): Promise<void> {
    const { list } = await this.sp.web.lists.ensure("9719_ProjectInformationDatabase", "Stores project metadata");
    await this.ensureTextField(list, "ProjectName");
    await this.ensureMultiLineField(list, "ProjectDescription");
    await this.ensureMultiLineField(list, "DeltekSubCodes");
    await this.ensureTextField(list, "Client");
    await this.ensureUserField(list, "PM");
    await this.ensureUserField(list, "Manager");
    await this.ensureUserField(list, "Checker");
    await this.ensureUserField(list, "Approver");
    await this.ensureChoiceField(list, "Status",  { Choices: projectStatusOptions.map(option => option.value).filter((value): value is string => value !== undefined) });
    await this.ensureChoiceField(list, "Sector",  { Choices: sectorOptions.map(option => option.value).filter((value): value is string => value !== undefined) });
    await this.ensureMultiLineField(list, "ClientContact");
    await this.ensureImageField(list, "ProjectImage");
  }

  /**
   * Ensures that the "ProjectTasks" list exists and is configured with the required fields.
   * Fields include text, choice, user, and datetime fields to store project task metadata.
   * It configures:
   * - TaskName: Text field for the task name.
   * - Status: Choice field with options for task status.
   * - DueDate, StartDate: Datetime fields for task start and end dates.
   * - AssignedTo: User field for the person assigned to the task.
   * - Description, Comments: Multi-line text fields for task description and comments.
   * - Priority: Choice field with predefined priority levels.
   * - TaskType: Choice field with predefined task types.
   * - CreatedBy, CreatedDate, ModifiedBy, ModifiedDate: User and datetime fields for tracking changes.
   * - TaskID, ParentTaskID: Text fields for task id and parent task id.
   * 
   * @returns A promise that resolves when the list and fields have been ensured.
   */
  private async ensureProjectTasks(): Promise<void> {
    const { list } = await this.sp.web.lists.ensure("9719_ProjectTasks", "Stores project tasks");
    await this.ensureTextField(list, "TaskName");
    await this.ensureChoiceField(list, "Status",  { Choices: taskStatusOptions.map(option => option.value).filter((value): value is string => value !== undefined) });
    await this.ensureDateTimeField(list, "DueDate");
    await this.ensureDateTimeField(list, "StartDate");
    await this.ensureUserField(list, "AssignedTo");
    await this.ensureMultiLineField(list, "Description");
    await this.ensureMultiLineField(list, "Comments");
    await this.ensureChoiceField(list, "Priority",  { Choices: taskPriorityOptions.map(option => option.value).filter((value): value is string => value !== undefined) });
    await this.ensureChoiceField(list, "TaskType",  { Choices: taskTypeOptions.map(option => option.value).filter((value): value is string => value !== undefined) });
    await this.ensureUserField(list, "CreatedBy");
    await this.ensureDateTimeField(list, "CreatedDate");
    await this.ensureUserField(list, "ModifiedBy");
    await this.ensureDateTimeField(list, "ModifiedDate");
    await this.ensureTextField(list, "TaskID");
    await this.ensureTextField(list, "ProjectID");
    await this.ensureTextField(list, "Progress");
  }

  /**
   * Ensures that the "9719_ProjectFees" list is configured with the required fields.
   * Fields include text, currency, and number fields to store project fee metadata.
   * It configures:
   * - FeeName: Text field for the fee name.
   * - FeeAmount, BudgetFee, SpendToDate: Currency fields for fee amount, budget, and spend to date.
   * - ForecastHours: Number field for forecasted hours.
   * - ActualHours: Number field for actual hours.
   * - RemianingBudgetOverspend: Currency field for remaining budget overspend.
   * 
   * @returns A promise that resolves when the list and fields have been ensured.
   */
  private async ensureProjectFees(): Promise<void> {
    const { list } = await this.sp.web.lists.ensure("9719_ProjectFees", "Stores budget and fees");
    await this.ensureTextField(list, "FeeName");
    await this.ensureCurrencyField(list, "FeeAmount", { CurrencyLocaleId: 2057 }); // 2057 is for UK English
    await this.ensureCurrencyField(list, "BudgetFee", { CurrencyLocaleId: 2057 });
    await this.ensureCurrencyField(list, "SpendToDate", { CurrencyLocaleId: 2057 });
    await this.ensureNumberField(list, "ForecastHours", { MinimumValue: 0 });
    await this.ensureNumberField(list, "ActualHours");
    await this.ensureCurrencyField(list, "RemianingBudgetOverspend", { CurrencyLocaleId: 2057 });
  }

  /**
   * Ensures that the "9719_ProjectDocuments" list is configured with the required fields.
   * Fields include text fields to store document metadata for various project documents.
   * It configures fields for:
   * - EIR, AIR, PIR, SMP, EIRAppraisal, PreContractBEP, BEP, MPDT, AIDP, LAP_EIR, DRM, MIDP, TIDPs, ResponsibilityAssignmentMatrix, IMRiskRegister, MobilisationPlan, FederatedModel, QAR, DataReports, HelthCheck_WarningReport
   * - ProjectExecutionPlan, ProjectManagementPlan, ProjectQualityPlan, ProjectControlPlan, ProjectRiskPlan, ProjectChangePlan, ProjectCostPlan, ProjectSchedulePlan, ProjectResourcePlan, ProjectCommunicationsPlan, ProjectProcurementPlan, ProjectStakeholderPlan
   * 
   * @returns A promise that resolves when the list and fields have been ensured.
   */
  private async ensureProjectDocuments(): Promise<void> {
    const { list } = await this.sp.web.lists.ensure("9719_ProjectDocuments", "Stores document metadata");
    await this.ensureTextField(list, "EIR");
    await this.ensureTextField(list, "AIR");
    await this.ensureTextField(list, "PIR");
    await this.ensureTextField(list, "SMP");
    await this.ensureTextField(list, "EIRAppraisal");
    await this.ensureTextField(list, "PreContractBEP");
    await this.ensureTextField(list, "BEP");
    await this.ensureTextField(list, "MPDT");
    await this.ensureTextField(list, "AIDP"); 
    await this.ensureTextField(list, "LAP_EIR");
    await this.ensureTextField(list, "DRM");
    await this.ensureTextField(list, "MIDP");
    await this.ensureTextField(list, "TIDPs");
    await this.ensureTextField(list, "ResponsibilityAssignmentMatrix");
    await this.ensureTextField(list, "IMRiskRegister");
    await this.ensureTextField(list, "MobilisationPlan");
    await this.ensureTextField(list, "FederatedModel");
    await this.ensureTextField(list, "QAR");
    await this.ensureTextField(list, "DataReports");
    await this.ensureTextField(list, "HelthCheck_WarningReport");
    await this.ensureTextField(list, "ProjectExecutionPlan");
    await this.ensureTextField(list, "ProjectManagementPlan");
    await this.ensureTextField(list, "ProjectQualityPlan");
    await this.ensureTextField(list, "ProjectControlPlan");
    await this.ensureTextField(list, "ProjectRiskPlan");
    await this.ensureTextField(list, "ProjectChangePlan");
    await this.ensureTextField(list, "ProjectCostPlan");
    await this.ensureTextField(list, "ProjectSchedulePlan");
    await this.ensureTextField(list, "ProjectResourcePlan");
    await this.ensureTextField(list, "ProjectCommunicationsPlan");
    await this.ensureTextField(list, "ProjectProcurementPlan");
    await this.ensureTextField(list, "ProjectStakeholderPlan");
  }

/**
 * Ensures that the "9719_ChangeControl" list exists and is configured with the required fields.
 * Fields include text, choice, and currency fields to track changes within a project.
 * It configures:
 * - ProjectNumber: Text field for the project number.
 * - ChangeDescription: Text field for the description of the change.
 * - ChangeRequestor: Text field for the name of the person requesting the change.
 * - ChangeRequestDate: Text field for the date when the change was requested.
 * - DCCNumber: Text field for the Document Control Center number.
 * - DDCLocation: Text field for the Document Distribution Center location.
 * - ChangeType: Choice field with options for the type of change (Scope, Budget, Timeline).
 * - Approved: Choice field to indicate if the change is approved (Yes, No, Pending).
 * - Fee: Currency field for the fee associated with the change.
 * 
 * @returns A promise that resolves when the list and fields have been ensured.
 */

  private async ensureChangeControl(): Promise<void> {
    const { list } = await this.sp.web.lists.ensure("9719_ChangeControl", "Tracks changes");
    await this.ensureTextField(list, "ProjectNumeber");
    await this.ensureTextField(list, "ChangeDescription");
    await this.ensureTextField(list, "ChangeRequestor");
    await this.ensureTextField(list, "ChangeRequestDate");
    await this.ensureTextField(list, "DCCNumber");
    await this.ensureTextField(list, "DDCLocation");
    await this.ensureChoiceField(list, "ChangeType",  { Choices: ["Scope", "Budget", "Timeline"] });
    await this.ensureChoiceField(list, "Approved",  { Choices: ["Yes", "No", "Pending"] });
    await this.ensureCurrencyField(list, "Fee", { CurrencyLocaleId: 2057 }); // 2057 is for UK English
  }

  /**
   * Ensures that the "9719_ProjectStages" list exists and is configured with the required fields.
   * Fields include text and date fields to track project stages.
   * It configures:
   * - ProjectNumber: Text field for the project number.
   * - StageName: Text field for the name of the project stage.
   * - StageDescription: Text field for the description of the project stage.
   * - StartDate, EndDate: Date fields for the start and end dates of the project stage.
   * 
   * @returns A promise that resolves when the list and fields have been ensured.
   */
  private async ensureProjectStages(): Promise<void> {
    const { list } = await this.sp.web.lists.ensure("9719_ProjectStages", "Stages of project");
    await this.ensureTextField(list, "ProjectNumber");
    await this.ensureTextField(list, "StageName");
    await this.ensureTextField(list, "StageDescription");
    await this.ensureDateTimeField(list, "StartDate");
    await this.ensureDateTimeField(list, "EndDate");
    await this.ensureTextField(list, "StageColor");
    await this.ensureChoiceField(list, "Status",  { Choices: projectStageStatusOptions.map(option => option.value).filter((value): value is string => value !== undefined) });
  }
}
