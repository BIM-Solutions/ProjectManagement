// Import required PnP JS and SharePoint dependencies
import { spfi, SPFx } from "@pnp/sp/presets/all";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SPFI } from "@pnp/sp";
import { IFileInfo, IFileUploadProgressData } from "@pnp/sp/files";
import "@pnp/sp/files";
import "@pnp/sp/folders";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/webs";

// Interface defining document metadata structure
export interface IDocument {
  Id: number;
  Title: string;
  // Title: string;
  FileLeafRef: string; // Name of file without path
  FileRef: string; // Full server relative path
  ProjectId: string; // Associated project identifier
  DocumentType: string; // Type of document
  Status: string; // Document status
  Version?: string; // Optional version number
  ModifiedBy: string; // Last modified by user
  Created: string; // Creation date
  Modified: string; // Last modified date
  IsFolder?: boolean; // Flag indicating if item is folder
  ServerRelativeUrl?: string; // Server relative URL
  Children?: IDocument[]; // Optional nested documents
  UniclassCode1?: string; // Level 1 Uniclass code
  UniclassTitle1?: string; // Level 1 Uniclass title
  UniclassCode2?: string; // Level 2 Uniclass code
  UniclassTitle2?: string; // Level 2 Uniclass title
  UniclassCode3?: string; // Level 3 Uniclass code
  UniclassTitle3?: string; // Level 3 Uniclass title
}

// Interface for folder information
export interface IFolderInfo {
  Name: string; // Folder name
  ServerRelativeUrl: string; // Server relative path
  ItemCount: number; // Number of items in folder
  TimeCreated: string; // Creation timestamp
  TimeLastModified: string; // Last modified timestamp
}

// Interface for upload progress callback function
export interface IUploadProgressCallback {
  (progress: number, file: File): void;
}

// Main service class for document management operations
export class DocumentService {
  private sp: SPFI; // SharePoint Framework instance
  private readonly documentsLibrary = "Projects"; // Library name
  private readonly documentsLibraryPath = "Projects"; // Library path
  private libraryChecked: boolean = false; // Flag to track library verification
  private checkedFolders: Set<string> = new Set(); // Cache of verified folders
  private documentCache: Map<string, { documents: IDocument[], timestamp: number }> = new Map(); // Cache for documents
  private readonly CACHE_DURATION = 30000; // Cache duration in milliseconds (30s)

  // Initialize service with SharePoint context
  constructor(context: WebPartContext) {
    this.sp = spfi().using(SPFx(context));
  }

  // Ensures document library exists with required fields
  public async ensureDocumentLibrary(documentLibrary: string): Promise<void> {
    if (this.libraryChecked) return;
    
    const maxRetries = 3;
    let retryCount = 0;

    while (retryCount < maxRetries) {
      try {
        // Check if library exists
        let listExists = false;
        try {
          await this.sp.web.lists.getByTitle(documentLibrary).select('Id')();
          listExists = true;
        } catch {
          listExists = false;
        }
        
        // Create library if it doesn't exist
        if (!listExists) {
          await this.sp.web.lists.ensure(documentLibrary, 'Documents Library', 101, true);
        }
        
        const list = this.sp.web.lists.getByTitle(documentLibrary);
        
        // Get existing fields
        const fields = await list.fields();
        const existingFields = new Set(fields.map(f => f.InternalName));

        // Define required metadata fields
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

        // Add missing fields to library
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
        // Exponential backoff retry delay
        await new Promise(resolve => setTimeout(resolve, Math.pow(2, retryCount) * 1000));
      }
    }
  }

  // Ensures project folder exists in document library
  private async ensureProjectFolder(projectId: string): Promise<void> {
    if (this.checkedFolders.has(projectId)) return;

    const maxRetries = 3;
    let retryCount = 0;

    while (retryCount < maxRetries) {
      try {
        const web = this.sp.web;
        const webData = await web.select('ServerRelativeUrl')();
        const folderPath = `${webData.ServerRelativeUrl}/${this.documentsLibraryPath}/${projectId}`;
        
        // Check if folder exists
        const result = await web.getFolderByServerRelativePath(folderPath).select('Exists')();
        
        // Create folder if it doesn't exist
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
        // Exponential backoff retry delay
        await new Promise(resolve => setTimeout(resolve, Math.pow(2, retryCount) * 1000));
      }
    }
  }

  // Retrieves all documents for a specific project
  public async getProjectDocuments(projectId: string): Promise<IDocument[]> {
    try {
      // Check cache before making API call
      const cached = this.documentCache.get(projectId);
      if (cached && Date.now() - cached.timestamp < this.CACHE_DURATION) {
        return cached.documents;
      }

      await this.ensureDocumentLibrary(this.documentsLibrary);
      await this.ensureProjectFolder(projectId);
      
      const maxRetries = 3;
      let retryCount = 0;
      let lastError: Error | null = null;

      while (retryCount < maxRetries) {
        try {
          // Fetch documents with metadata
          const list =  await this.sp.web.lists.getByTitle(this.documentsLibrary);
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

          // Map SharePoint items to IDocument interface
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

          // Update cache with fresh data
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
          // Exponential backoff retry delay
          await new Promise(resolve => setTimeout(resolve, Math.pow(2, retryCount) * 1000));
        }
      }

      throw lastError || new Error('Failed to get project documents after retries');
    } catch (error) {
      console.error('Error getting project documents:', error);
      throw new Error('Failed to get project documents: ' + error.message);
    }
  }

  // Uploads a document with metadata to a project folder
  public async uploadDocument(
    projectId: string,
    file: File,
    metadata: Partial<IDocument>,
    onProgress?: IUploadProgressCallback
  ): Promise<IDocument> {
    try {
      await this.ensureDocumentLibrary(this.documentsLibrary);
      const result = await this.ensureProjectFolder(projectId);
      console.log('the result is: ', result);
      
      const web = this.sp.web;
      const webData = await web.select('ServerRelativeUrl')();
      const folderPath = `${webData.ServerRelativeUrl}/${this.documentsLibraryPath}/${projectId}`;
      console.log('the project  id is: ', projectId);
      console.log('Uploading to path:', folderPath);
      
      // Use chunked upload for large files
      if (file.size > 10 * 1024 * 1024) {
        return await this.uploadLargeFile(folderPath, file, metadata, projectId, onProgress);
      }

      // Regular upload for smaller files
      console.log('Attempting to upload file:', file.name);
      const fileResult = await this.sp.web.getFolderByServerRelativePath(folderPath)
        .files.addUsingPath(file.name, file, { Overwrite: true });

      console.log('File uploaded successfully:', fileResult.ServerRelativeUrl);

      // Update file metadata
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

  // Handles chunked upload for large files
  private async uploadLargeFile(
    folderPath: string,
    file: File,
    metadata: Partial<IDocument>,
    projectId: string,
    onProgress?: IUploadProgressCallback
  ): Promise<IDocument> {
    try {
      const totalSize = file.size;

      // Upload file in chunks with progress tracking
      const result = await this.sp.web.getFolderByServerRelativePath(folderPath)
        .files.addChunked(file.name, file, {
          progress: (data: IFileUploadProgressData) => {
            if (onProgress) {
              onProgress((data.offset / totalSize) * 100, file);
            }
          }
        });

      // Update metadata after upload
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

  // Deletes a document by ID
  public async deleteDocument(documentId: number): Promise<void> {
    try {
      await this.ensureDocumentLibrary(this.documentsLibrary);
      await this.sp.web.lists
        .getByTitle(this.documentsLibrary)
        .items.getById(documentId)
        .delete();
    } catch (error) {
      console.error('Error deleting document:', error);
      throw new Error('Failed to delete document: ' + error.message);
    }
  }

  // Checks out a document for editing
  public async checkoutDocument(documentId: number): Promise<void> {
    try {
      await this.ensureDocumentLibrary(this.documentsLibrary);
      const item = await this.sp.web.lists
        .getByTitle(this.documentsLibrary)
        .items.getById(documentId)();
      
      await this.sp.web.getFileByServerRelativePath(item.FileRef).checkout();
    } catch (error) {
      console.error('Error checking out document:', error);
      throw new Error('Failed to check out document: ' + error.message);
    }
  }

  // Checks in a document with comment
  public async checkinDocument(documentId: number, comment: string): Promise<void> {
    try {
      await this.ensureDocumentLibrary(this.documentsLibrary);
      const item = await this.sp.web.lists
        .getByTitle(this.documentsLibrary)
        .items.getById(documentId)();
      
      await this.sp.web.getFileByServerRelativePath(item.FileRef).checkin(comment);
    } catch (error) {
      console.error('Error checking in document:', error);
      throw new Error('Failed to check in document: ' + error.message);
    }
  }

  // Gets version history for a document
  public async getDocumentVersions(documentId: number): Promise<IFileInfo[]> {
    try {
      await this.ensureDocumentLibrary(this.documentsLibrary);
      const item = await this.sp.web.lists
        .getByTitle(this.documentsLibrary)
        .items.getById(documentId)();
      
      return await this.sp.web.getFileByServerRelativePath(item.FileRef).versions();
    } catch (error) {
      console.error('Error getting document versions:', error);
      throw new Error('Failed to get document versions: ' + error.message);
    }
  }

  // Gets contents of a folder including files and subfolders
  public async getFolderContents(folderPath: string): Promise<IDocument[]> {
    try {
      const web = this.sp.web;
      const folder = web.getFolderByServerRelativePath(folderPath);
      
      // Get folders and map to IDocument structure
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

      // Get files and map to IDocument structure
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

  // Creates a new folder at specified path
  public async createFolder(folderPath: string): Promise<IFolderInfo | undefined> {
    try {
      const web = this.sp.web;
      let folder;
      try {
        folder = await web.getFolderByServerRelativePath(folderPath)();
        // Folder exists, use it
      } catch {
        // Folder doesn't exist, create it
        folder = await web.folders.addUsingPath(folderPath);
      }
      return {
        Name: folder.Name,
        ServerRelativeUrl: folder.ServerRelativeUrl,
        ItemCount: 0,
        TimeCreated: new Date().toISOString(),
        TimeLastModified: new Date().toISOString()
      };
    } catch {
      // Silently handle errors
    }
  }
} 