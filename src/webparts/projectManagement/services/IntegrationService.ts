import { WebPartContext } from "@microsoft/sp-webpart-base";
import { spfi, SPFx } from "@pnp/sp/presets/all";
import { SPFI } from "@pnp/sp";

export interface IViewPointData {
  projectId: string;
  status: string;
  lastSync: string;
  data: string;
}

export interface IDeltekData {
  projectId: string;
  timeEntries: string;
  expenses: string;
  lastSync: string;
}

export interface ISyncData {
  lastSync: string;
  data: IViewPointData | IDeltekData;
}

export class IntegrationService {
  private sp: SPFI;
  private readonly integrationListName = "SystemIntegrations";
  private readonly viewPointApiUrl: string;
  private readonly deltekApiUrl: string;

  constructor(context: WebPartContext) {
    this.sp = spfi().using(SPFx(context));
    // These would come from configuration
    this.viewPointApiUrl = "https://api.viewpoint.com";
    this.deltekApiUrl = "https://api.deltek.com";
  }

  public async syncWithViewPoint(projectId: string): Promise<IViewPointData> {
    // This is a placeholder for the actual API integration
    // You would need to implement the actual API calls to ViewPoint
    const response = await fetch(`${this.viewPointApiUrl}/projects/${projectId}`);
    const data = await response.json();

    // Store the sync data in SharePoint
    await this.sp.web.lists.getByTitle(this.integrationListName).items.add({
      ProjectId: projectId,
      SystemName: "ViewPoint",
      LastSync: new Date().toISOString(),
      SyncData: JSON.stringify(data)
    });

    return {
      projectId,
      status: "synced",
      lastSync: new Date().toISOString(),
      data
    };
  }

  public async syncWithDeltek(projectId: string): Promise<IDeltekData> {
    // This is a placeholder for the actual API integration
    // You would need to implement the actual API calls to Deltek
    const timeResponse = await fetch(`${this.deltekApiUrl}/projects/${projectId}/time`);
    const expenseResponse = await fetch(`${this.deltekApiUrl}/projects/${projectId}/expenses`);
    
    const timeEntries = await timeResponse.json();
    const expenses = await expenseResponse.json();

    // Store the sync data in SharePoint
    await this.sp.web.lists.getByTitle(this.integrationListName).items.add({
      ProjectId: projectId,
      SystemName: "Deltek",
      LastSync: new Date().toISOString(),
      SyncData: JSON.stringify({ timeEntries, expenses })
    });

    return {
      projectId,
      timeEntries,
      expenses,
      lastSync: new Date().toISOString()
    };
  }

  public async getLastSyncData(projectId: string, system: "ViewPoint" | "Deltek"): Promise<ISyncData | undefined> {
    const items = await this.sp.web.lists
      .getByTitle(this.integrationListName)
      .items.filter(`ProjectId eq '${projectId}' and SystemName eq '${system}'`)
      .orderBy("LastSync", false)
      .top(1)();

    if (items.length > 0) {
      return {
        lastSync: items[0].LastSync,
        data: JSON.parse(items[0].SyncData)
      };
    }

    return undefined;
  }
} 