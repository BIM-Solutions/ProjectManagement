import { spfi, SPFI } from "@pnp/sp";
import { SPFx } from "@pnp/sp/presets/all";
import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface ITask {
  Id: number;
  Title: string;
  DueDate: string;
  StartDate?: string;
  Priority: string;
  Status: string;
  Project: string;
  AssignedTo?: {
    Title: string;
    EMail: string;
  };
  Description?: string;
}

class TasksService {
  private static instance: TasksService;
  private sp: SPFI;
  private tasksCache: Map<string, ITask[]> = new Map();
  private lastFetchTime: Map<string, number> = new Map();
  private readonly CACHE_DURATION = 5 * 60 * 1000; // 5 minutes in milliseconds

  private constructor(context: WebPartContext) {
    this.sp = spfi().using(SPFx(context));
  }

  public static getInstance(context: WebPartContext): TasksService {
    if (!TasksService.instance) {
      TasksService.instance = new TasksService(context);
    }
    return TasksService.instance;
  }

  private getCacheKey(listName: string, userDisplayName: string): string {
    return `${listName}-${userDisplayName}`;
  }

  private isCacheValid(cacheKey: string): boolean {
    const lastFetch = this.lastFetchTime.get(cacheKey);
    if (!lastFetch) return false;
    return Date.now() - lastFetch < this.CACHE_DURATION;
  }

  public async getTasks(
    listName: string,
    userDisplayName: string
  ): Promise<ITask[]> {
    const cacheKey = this.getCacheKey(listName, userDisplayName);

    // Return cached tasks if they exist and are still valid
    if (this.tasksCache.has(cacheKey) && this.isCacheValid(cacheKey)) {
      return this.tasksCache.get(cacheKey)!;
    }

    try {
      const items = await this.sp.web.lists
        .getByTitle(listName)
        .items.filter(`AssignedTo/Title eq '${userDisplayName}'`)();

      // Update cache
      this.tasksCache.set(cacheKey, items);
      this.lastFetchTime.set(cacheKey, Date.now());

      return items;
    } catch (error) {
      console.error("Error fetching tasks:", error);
      throw error;
    }
  }

  public async addTask(
    listName: string,
    task: Partial<ITask>,
    userDisplayName: string
  ): Promise<void> {
    try {
      await this.sp.web.lists.getByTitle(listName).items.add({
        Title: task.Title,
        DueDate: task.DueDate,
        Priority: task.Priority,
        Status: "Not Started",
        Project: task.Project,
        AssignedTo: { Title: userDisplayName },
      });

      // Invalidate cache for this user
      const cacheKey = this.getCacheKey(listName, userDisplayName);
      this.tasksCache.delete(cacheKey);
      this.lastFetchTime.delete(cacheKey);
    } catch (error) {
      console.error("Error adding task:", error);
      throw error;
    }
  }

  public async updateTask(listName: string, task: ITask): Promise<void> {
    try {
      // console.log("new jtash", task);
      await this.sp.web.lists
        .getByTitle(listName)
        .items.getById(task.Id)
        .update({
          Title: task.Title,
          DueDate: task.DueDate,
          StartDate: task.StartDate,
          Priority: task.Priority,
          Status: task.Status,
          ProjectID: task.Project,
        });

      // Invalidate cache for this user
      const cacheKey = this.getCacheKey(listName, task.AssignedTo?.Title || "");
      this.tasksCache.delete(cacheKey);
      this.lastFetchTime.delete(cacheKey);
    } catch (error) {
      console.error("Error updating task:", error);
      throw error;
    }
  }

  public async updateTaskStatus(
    listName: string,
    taskId: number,
    newStatus: string,
    assignedTo?: string
  ): Promise<void> {
    try {
      await this.sp.web.lists
        .getByTitle(listName)
        .items.getById(taskId)
        .update({
          Status: newStatus,
        });
      // Invalidate cache for this user
      const cacheKey = this.getCacheKey(listName, assignedTo || "");
      this.tasksCache.delete(cacheKey);
      this.lastFetchTime.delete(cacheKey);
    } catch (error) {
      console.error("Error updating task status:", error);
      throw error;
    }
  }

  public clearCache(listName: string, userDisplayName: string): void {
    const cacheKey = this.getCacheKey(listName, userDisplayName);
    this.tasksCache.delete(cacheKey);
    this.lastFetchTime.delete(cacheKey);
  }
}

export default TasksService;
