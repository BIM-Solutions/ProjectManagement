import { spfi, SPFx } from "@pnp/sp/presets/all";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SPFI } from "@pnp/sp";
import { IFileInfo, IFileUploadProgressData } from "@pnp/sp/files";
import "@pnp/sp/files";
import "@pnp/sp/folders";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/webs";

export interface IDocument {
  Id: number;
  Title: string;
  // Title: string;
  FileLeafRef: string;
  FileRef: string;
  ProjectId: string;
  DocumentType: string;
  Status: string;
  Version?: string;
  ModifiedBy: string;
  Created: string;
  Modified: string;
  IsFolder?: boolean;
  ServerRelativeUrl?: string;
  Children?: IDocument[];
  UniclassCode1?: string;
  UniclassTitle1?: string;
  UniclassCode2?: string;
  UniclassTitle2?: string;
  UniclassCode3?: string;
  UniclassTitle3?: string;
}

export interface IFolderInfo {
  Name: string;
  ServerRelativeUrl: string;
  ItemCount: number;
  TimeCreated: string;
  TimeLastModified: string;
}

export interface IUploadProgressCallback {
  (progress: number, file: File): void;
}

export class DocumentService {
  private sp: SPFI;
  private readonly documentsLibrary = "Projects";
  private readonly documentsLibraryPath = "Projects";
  private libraryChecked: boolean = false;
  private checkedFolders: Set<string> = new Set();
  private documentCache: Map<string, { documents: IDocument[], timestamp: number }> = new Map();
  private readonly CACHE_DURATION = 30000; // 30 seconds cache

  constructor(context: WebPartContext) {
    this.sp = spfi().using(SPFx(context));
  }

  private async ensureDocumentLibrary(): Promise<void> {
    if (this.libraryChecked) return;
    
    const maxRetries = 3;
    let retryCount = 0;

    while (retryCount < maxRetries) {
      try {
        // First check if the list exists by trying to get its properties
        let listExists = false;
        try {
          await this.sp.web.lists.getByTitle(this.documentsLibrary).select('Id')();
          listExists = true;
        } catch {
          listExists = false;
        }
        
        if (!listExists) {
          // Only try to create if it doesn't exist
          await this.sp.web.lists.ensure(this.documentsLibrary, 'Documents Library', 101, true);
        }
        
        const list = this.sp.web.lists.getByTitle(this.documentsLibrary);
        
        // Get existing fields
        const fields = await list.fields();
        const existingFields = new Set(fields.map(f => f.InternalName));

        // Define required fields
        const requiredFields = [
          { name: "ProjectId", required: true },
          { name: "DocumentType", required: true },
          { name: "Status", required: false },
          { name: "UniclassCode1", required: false },
          { name: "UniclassTitle1", required: false },
          { name: "UniclassCode2", required: false },
          { name: "UniclassTitle2", required: false },
          { name: "UniclassCode3", required: false },
          { name: "UniclassTitle3", required: false }
        ];

        // Add only missing fields
        for (const field of requiredFields) {
          if (!existingFields.has(field.name)) {
            try {
              await list.fields.addText(field.name, { Required: field.required });
            } catch (error) {
              console.warn(`Error adding field ${field.name}:`, error);
            }
          }
        }
        
        this.libraryChecked = true;
        return;
      } catch (error) {
        retryCount++;
        if (retryCount === maxRetries) {
          console.error('Error accessing document library after retries:', error);
          throw new Error('Failed to access document library: ' + error.message);
        }
        // Wait before retrying (exponential backoff)
        await new Promise(resolve => setTimeout(resolve, Math.pow(2, retryCount) * 1000));
      }
    }
  }

  private async ensureProjectFolder(projectId: string): Promise<void> {
    if (this.checkedFolders.has(projectId)) return;

    const maxRetries = 3;
    let retryCount = 0;

    while (retryCount < maxRetries) {
      try {
        const web = this.sp.web;
        const webData = await web.select('ServerRelativeUrl')();
        const folderPath = `${webData.ServerRelativeUrl}/${this.documentsLibraryPath}/${projectId}`;
        
        const result = await web.getFolderByServerRelativePath(folderPath).select('Exists')();
        
        if (!result.Exists) {
          const libraryPath = `${webData.ServerRelativeUrl}/${this.documentsLibraryPath}`;
          await web.getFolderByServerRelativePath(libraryPath).addSubFolderUsingPath(projectId);
        }
        
        this.checkedFolders.add(projectId);
        return;
      } catch (error) {
        retryCount++;
        if (retryCount === maxRetries) {
          console.error('Error ensuring project folder after retries:', error);
          throw new Error('Failed to create project folder: ' + error.message);
        }
        // Wait before retrying (exponential backoff)
        await new Promise(resolve => setTimeout(resolve, Math.pow(2, retryCount) * 1000));
      }
    }
  }

  public async getProjectDocuments(projectId: string): Promise<IDocument[]> {
    try {
      // Check cache first
      const cached = this.documentCache.get(projectId);
      if (cached && Date.now() - cached.timestamp < this.CACHE_DURATION) {
        return cached.documents;
      }

      await this.ensureDocumentLibrary();
      await this.ensureProjectFolder(projectId);
      
      const maxRetries = 3;
      let retryCount = 0;
      let lastError: Error | null = null;

      while (retryCount < maxRetries) {
        try {
          const list =  await this.sp.web.lists.getByTitle(this.documentsLibrary);
          // console.log('the list is: ', list.items());
          const items = await list.items
            .select(
              "Id",
              "Title",
              "FileLeafRef",
              "FileRef",
              "ProjectId",
              "DocumentType",
              "Status",
              "Editor/Title",
              "Created",
              "Modified",
              "UniclassCode1",
              "UniclassTitle1",
              "UniclassCode2",
              "UniclassTitle2",
              "UniclassCode3",
              "UniclassTitle3"
            )
            .expand("Editor")
            .filter(`ProjectId eq '${projectId}'`)();

          const documents = items.map(item => ({
            Id: item.Id,
            Title: item.Title,
            FileLeafRef: item.FileLeafRef,
            FileRef: item.FileRef,
            ProjectId: item.ProjectId,
            DocumentType: item.DocumentType,
            Status: item.Status,
            ModifiedBy: item.Editor?.Title || '',
            Created: item.Created,
            Modified: item.Modified,
            UniclassCode1: item.UniclassCode1,
            UniclassTitle1: item.UniclassTitle1,
            UniclassCode2: item.UniclassCode2,
            UniclassTitle2: item.UniclassTitle2,
            UniclassCode3: item.UniclassCode3,
            UniclassTitle3: item.UniclassTitle3
          }));

          // Update cache
          this.documentCache.set(projectId, {
            documents,
            timestamp: Date.now()
          });

          return documents;
        } catch (error) {
          lastError = error;
          retryCount++;
          if (retryCount === maxRetries) {
            console.error('Error getting project documents after retries:', error);
            throw new Error('Failed to get project documents: ' + error.message);
          }
          // Wait before retrying (exponential backoff)
          await new Promise(resolve => setTimeout(resolve, Math.pow(2, retryCount) * 1000));
        }
      }

      throw lastError || new Error('Failed to get project documents after retries');
    } catch (error) {
      console.error('Error getting project documents:', error);
      throw new Error('Failed to get project documents: ' + error.message);
    }
  }

  public async uploadDocument(
    projectId: string,
    file: File,
    metadata: Partial<IDocument>,
    onProgress?: IUploadProgressCallback
  ): Promise<IDocument> {
    try {
      await this.ensureDocumentLibrary();
      const result = await this.ensureProjectFolder(projectId);
      console.log('the result is: ', result);
      
      const web = this.sp.web;
      const webData = await web.select('ServerRelativeUrl')();
      const folderPath = `${webData.ServerRelativeUrl}/${this.documentsLibraryPath}/${projectId}`;
      console.log('the project  id is: ', projectId);
      console.log('Uploading to path:', folderPath);
      
      // For large files (> 10MB), use chunked upload
      if (file.size > 10 * 1024 * 1024) {
        return await this.uploadLargeFile(folderPath, file, metadata, projectId, onProgress);
      }

      // For smaller files, use regular upload
      console.log('Attempting to upload file:', file.name);
      const fileResult = await this.sp.web.getFolderByServerRelativePath(folderPath)
        .files.addUsingPath(file.name, file, { Overwrite: true });

      console.log('File uploaded successfully:', fileResult.ServerRelativeUrl);

      // Update metadata
      const item = await this.sp.web.getFileByServerRelativePath(fileResult.ServerRelativeUrl).getItem();
      await item.update({
        ...metadata,
        ProjectId: projectId,
        Title: file.name
      });

      return await item();
    } catch (error) {
      console.error('Error uploading document:', error);
      throw new Error('Failed to upload document: ' + error.message);
    }
  }

  private async uploadLargeFile(
    folderPath: string,
    file: File,
    metadata: Partial<IDocument>,
    projectId: string,
    onProgress?: IUploadProgressCallback
  ): Promise<IDocument> {
    try {
      const totalSize = file.size;

      // Start chunked upload with progress tracking
      const result = await this.sp.web.getFolderByServerRelativePath(folderPath)
        .files.addChunked(file.name, file, {
          progress: (data: IFileUploadProgressData) => {
            if (onProgress) {
              onProgress((data.offset / totalSize) * 100, file);
            }
          }
        });

      // Update metadata
      const item = await this.sp.web.getFileByServerRelativePath(result.ServerRelativeUrl).getItem();
      await item.update({
        ...metadata,
        ProjectId: projectId,
        Title: file.name
      });

      return await item();
    } catch (error) {
      console.error('Error uploading large file:', error);
      throw new Error('Failed to upload large file: ' + error.message);
    }
  }

  public async deleteDocument(documentId: number): Promise<void> {
    try {
      await this.ensureDocumentLibrary();
      await this.sp.web.lists
        .getByTitle(this.documentsLibrary)
        .items.getById(documentId)
        .delete();
    } catch (error) {
      console.error('Error deleting document:', error);
      throw new Error('Failed to delete document: ' + error.message);
    }
  }

  public async checkoutDocument(documentId: number): Promise<void> {
    try {
      await this.ensureDocumentLibrary();
      const item = await this.sp.web.lists
        .getByTitle(this.documentsLibrary)
        .items.getById(documentId)();
      
      await this.sp.web.getFileByServerRelativePath(item.FileRef).checkout();
    } catch (error) {
      console.error('Error checking out document:', error);
      throw new Error('Failed to check out document: ' + error.message);
    }
  }

  public async checkinDocument(documentId: number, comment: string): Promise<void> {
    try {
      await this.ensureDocumentLibrary();
      const item = await this.sp.web.lists
        .getByTitle(this.documentsLibrary)
        .items.getById(documentId)();
      
      await this.sp.web.getFileByServerRelativePath(item.FileRef).checkin(comment);
    } catch (error) {
      console.error('Error checking in document:', error);
      throw new Error('Failed to check in document: ' + error.message);
    }
  }

  public async getDocumentVersions(documentId: number): Promise<IFileInfo[]> {
    try {
      await this.ensureDocumentLibrary();
      const item = await this.sp.web.lists
        .getByTitle(this.documentsLibrary)
        .items.getById(documentId)();
      
      return await this.sp.web.getFileByServerRelativePath(item.FileRef).versions();
    } catch (error) {
      console.error('Error getting document versions:', error);
      throw new Error('Failed to get document versions: ' + error.message);
    }
  }

  public async getFolderContents(folderPath: string): Promise<IDocument[]> {
    try {
      const web = this.sp.web;
      const folder = web.getFolderByServerRelativePath(folderPath);
      
      // Get folders
      const folders = await folder.folders();
      const folderItems = await Promise.all(folders.map(async (f) => {
        return {
          Id: 0,
          Title: f.Name,
          Description: "",
          FileLeafRef: f.Name,
          FileRef: f.ServerRelativeUrl,
          ProjectId: "",
          DocumentType: "Folder",
          Status: "",
          ModifiedBy: "",
          Created: f.TimeCreated,
          Modified: f.TimeLastModified,
          IsFolder: true,
          ServerRelativeUrl: f.ServerRelativeUrl,
          Children: [],
          UniclassCode1: "",
          UniclassTitle1: "",
          UniclassCode2: "",
          UniclassTitle2: "",
          UniclassCode3: "",
          UniclassTitle3: ""
        };
      }));

      // Get files
      const files = await folder.files();
      const fileItems = await Promise.all(files.map(async (f) => {
        const item = await this.sp.web.getFileByServerRelativePath(f.ServerRelativeUrl).listItemAllFields();
        return {
          Id: item.Id,
          Title: f.Name,
          Description: item.Title || "",
          FileLeafRef: f.Name,
          FileRef: f.ServerRelativeUrl,
          ProjectId: item.ProjectId || "",
          DocumentType: item.DocumentType || "",
          Status: item.Status || "",
          ModifiedBy: "",
          Created: f.TimeCreated,
          Modified: f.TimeLastModified,
          IsFolder: false,
          ServerRelativeUrl: f.ServerRelativeUrl,
          UniclassCode1: item.UniclassCode1 || "",
          UniclassTitle1: item.UniclassTitle1 || "",
          UniclassCode2: item.UniclassCode2 || "",
          UniclassTitle2: item.UniclassTitle2 || "",
          UniclassCode3: item.UniclassCode3 || "",
          UniclassTitle3: item.UniclassTitle3 || ""
        };
      }));

      return [...folderItems, ...fileItems];
    } catch (error) {
      console.error('Error getting folder contents:', error);
      throw new Error('Failed to get folder contents: ' + error.message);
    }
  }

  public async createFolder(folderPath: string): Promise<IFolderInfo> {
    try {
      const web = this.sp.web;
      const folder = await web.folders.addUsingPath(folderPath);
      return {
        Name: folder.Name,
        ServerRelativeUrl: folder.ServerRelativeUrl,
        ItemCount: 0,
        TimeCreated: new Date().toISOString(),
        TimeLastModified: new Date().toISOString()
      };
    } catch (error) {
      console.error('Error creating folder:', error);
      throw new Error('Failed to create folder: ' + error.message);
    }
  }
} 