import { spfi, SPFx } from "@pnp/sp/presets/all";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SPFI } from "@pnp/sp";
import { IFileInfo } from "@pnp/sp/files";
import "@pnp/sp/files";
import "@pnp/sp/folders";

export interface IDocument {
  Id: number;
  Title: string;
  FileLeafRef: string;
  FileRef: string;
  ProjectId: string;
  DocumentType: string;
  Status: string;
  Version: string;
  ModifiedBy: string;
  Created: string;
  Modified: string;
}

export class DocumentService {
  private sp: SPFI;
  private readonly documentsLibrary = "ProjectDocuments";

  constructor(context: WebPartContext) {
    this.sp = spfi().using(SPFx(context));
  }

  public async getProjectDocuments(projectId: string): Promise<IDocument[]> {
    return await this.sp.web.lists
      .getByTitle(this.documentsLibrary)
      .items.filter(`ProjectId eq '${projectId}'`)
      .select(
        "Id",
        "Title",
        "FileLeafRef",
        "FileRef",
        "ProjectId",
        "DocumentType",
        "Status",
        "Version",
        "ModifiedBy",
        "Created",
        "Modified"
      )();
  }

  public async uploadDocument(
    projectId: string,
    file: File,
    metadata: Partial<IDocument>
  ): Promise<IDocument> {
    const projectFolder = `${this.documentsLibrary}/${projectId}`;
    
    // Ensure project folder exists
    try {
      await this.sp.web.folders.getByUrl(projectFolder).select("Exists")();
    } catch {
      await this.sp.web.folders.addUsingPath(projectFolder);
    }

    // Upload file
    const fileResult = await this.sp.web.getFolderByServerRelativePath(projectFolder)
      .files.addUsingPath(file.name, file, { Overwrite: true });

    // Update metadata
    const item = await this.sp.web.getFileByServerRelativePath(fileResult.ServerRelativeUrl).getItem();
    await item.update({
      ...metadata,
      ProjectId: projectId
    });

    return await item();
  }

  public async deleteDocument(documentId: number): Promise<void> {
    await this.sp.web.lists
      .getByTitle(this.documentsLibrary)
      .items.getById(documentId)
      .delete();
  }

  public async checkoutDocument(documentId: number): Promise<void> {
    const item = await this.sp.web.lists
      .getByTitle(this.documentsLibrary)
      .items.getById(documentId)();
    
    await this.sp.web.getFileByServerRelativePath(item.FileRef).checkout();
  }

  public async checkinDocument(documentId: number, comment: string): Promise<void> {
    const item = await this.sp.web.lists
      .getByTitle(this.documentsLibrary)
      .items.getById(documentId)();
    
    await this.sp.web.getFileByServerRelativePath(item.FileRef).checkin(comment);
  }

  public async getDocumentVersions(documentId: number): Promise<IFileInfo[]> {
    const item = await this.sp.web.lists
      .getByTitle(this.documentsLibrary)
      .items.getById(documentId)();
    
    return await this.sp.web.getFileByServerRelativePath(item.FileRef).versions();
  }
} 