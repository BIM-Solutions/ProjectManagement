import { SPFI } from "@pnp/sp";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { IFilePickerResult } from "@pnp/spfx-controls-react";

export interface IStandard {
  Id: number;
  Title: string;
  Client: string;
  Version: string;
  uniclassCode: string;
  FileRef: string;
}

export class StandardsService {
  private sp: SPFI;
  //   private context: WebPartContext;
  private readonly STANDARDS_LIBRARY = "Standards";
  private readonly TEMPLATES_LIBRARY = "Templates";

  // Initialize the service with SharePoint context and SPFI instance
  constructor(context: WebPartContext, sp: SPFI) {
    this.sp = sp;
    this.documentService = new DocumentService(context);
  }

  public async uploadNewStandard(
    file: IFilePickerResult,
    projectNumber: string
  ): Promise<void> {
    try {
      // Upload to Standards library
      const standardsFolder = `${this.STANDARDS_LIBRARY}/${projectNumber}`;
      await this.sp.web
        .getFolderByServerRelativePath(standardsFolder)
        .files.addUsingPath(file.fileName, await file.downloadFileContent(), {
          Overwrite: true,
        });

      // Copy to project folder
      const projectFolder = `Projects/${projectNumber}/Standards`;
      await this.copyToProjectFolder(
        file.fileName,
        standardsFolder,
        projectFolder
      );

      // Update metadata
      await this.updateProjectMetadata(projectNumber, file.fileName);
    } catch (error) {
      console.error("Error uploading standard:", error);
      throw error;
    }
  }

  public async getExistingStandards(): Promise<IStandard[]> {
    try {
      const items = await this.sp.web.lists
        .getByTitle(this.STANDARDS_LIBRARY)
        .items.select(
          "Id",
          "Title",
          "Client",
          "Version",
          "UniclassCode",
          "FileRef"
        )();

      return items as unknown as IStandard[];
    } catch (error) {
      console.error("Error fetching standards:", error);
      throw error;
    }
  }

  public async copyExistingStandard(
    standard: IStandard,
    projectNumber: string
  ): Promise<void> {
    try {
      const sourceFolder = `${this.STANDARDS_LIBRARY}/${standard.Client}/${standard.Version}`;
      const projectFolder = `Projects/${projectNumber}/Standards`;
      await this.copyToProjectFolder(
        standard.Title,
        this.STANDARDS_LIBRARY,
        projectFolder
      );
      await this.updateProjectMetadata(projectNumber, standard.Title);
    } catch (error) {
      console.error("Error copying standard:", error);
      throw error;
    }
  }

  private async copyToProjectFolder(
    fileName: string,
    sourceFolder: string,
    targetFolder: string
  ): Promise<void> {
    try {
      const sourceFile = await this.sp.web.getFileByServerRelativePath(
        `${sourceFolder}/${fileName}`
      );
      await sourceFile.copyTo(`${targetFolder}/${fileName}`, true);
    } catch (error) {
      console.error("Error copying file:", error);
      throw error;
    }
  }

  private async updateProjectMetadata(
    projectNumber: string,
    fileName: string
  ): Promise<void> {
    try {
      const file = await this.sp.web.getFileByServerRelativePath(
        `Projects/${projectNumber}/Standards/${fileName}`
      );
      const metadata = await file.getItem();
      const metadataValues = await metadata();

      // Update project's report metadata
      await this.sp.web.lists
        .getByTitle("Projects")
        .items.getById(parseInt(projectNumber))
        .update({
          StandardClient: metadataValues.Client,
          StandardVersion: metadataValues.Version,
          StandardUniclass: metadataValues.UniclassCode,
        });
    } catch (error) {
      console.error("Error updating metadata:", error);
      throw error;
    }
  }

  public async copyTemplates(
    projectNumber: string,
    uniclassCode: string
  ): Promise<void> {
    try {
      // Get templates matching the Uniclass code
      const templates = await this.sp.web.lists
        .getByTitle(this.TEMPLATES_LIBRARY)
        .items.filter(`UniclassCode eq '${uniclassCode}'`)();

      // Copy each matching template
      for (const template of templates) {
        const sourceFile = await this.sp.web.getFileByServerRelativePath(
          template.FileRef
        );
        const newFileName = `${projectNumber}_${uniclassCode}_${template.Title}.docx`;
        await sourceFile.copyTo(
          `Projects/${projectNumber}/Reports/${newFileName}`,
          true
        );
      }
    } catch (error) {
      console.error("Error copying templates:", error);
      throw error;
    }
  }

  // Fetch Uniclass data from JSON file
  private async fetchUniclass(
    context: WebPartContext
  ): Promise<IUniclassItem[]> {
    try {
      const response = await fetch(
        `${context.pageContext.web.absoluteUrl}/SiteAssets/Uniclass2015_PM.json`
      );
      if (!response.ok) throw new Error("Failed to load Uniclass data");
      const data = await response.json();
      return data;
    } catch (e) {
      console.error("Error fetching Uniclass data:", e);
      throw e;
    }
  }

  // Map Uniclass code to full hierarchy of codes and titles
  public async mapUniclassCodes(
    uniclassCode: string,
    context: WebPartContext
  ): Promise<{
    UniclassTitle1: string;
    UniclassCode1: string;
    UniclassTitle2: string;
    UniclassCode2: string;
    UniclassTitle3: string;
    UniclassCode3: string;
  }> {
    const uniclassData = await this.fetchUniclass(context);

    // Navigate through hierarchy to find matching code
    for (const l1 of uniclassData) {
      if (l1.subgroups) {
        for (const l2 of l1.subgroups) {
          if (l2.sections) {
            for (const l3 of l2.sections) {
              if (l3.code === uniclassCode) {
                return {
                  UniclassTitle1: l1.title,
                  UniclassCode1: l1.code,
                  UniclassTitle2: l2.title,
                  UniclassCode2: l2.code,
                  UniclassTitle3: l3.title,
                  UniclassCode3: l3.code,
                };
              }
            }
          }
        }
      }
    }
    throw new Error("Uniclass code not found");
  }

  // Upload a native file with metadata
  public async uploadNativeFile(
    file: File,
    version: string,
    client: string,
    projectNumber: string,
    metadata: { uniclassCode?: string; title?: string; description?: string },
    context: WebPartContext
  ): Promise<void> {
    try {
      // Map Uniclass codes and ensure library exists
      const uniclassCodes = await this.mapUniclassCodes(
        metadata.uniclassCode || "",
        context
      );
      await this.documentService.ensureDocumentLibrary(this.STANDARDS_LIBRARY);

      // Create folder structure
      const clientFolder = `${this.STANDARDS_LIBRARY}/${client}`;
      const versionFolder = `${clientFolder}/${version}`;
      try {
        await this.documentService.createFolder(clientFolder);
      } catch {
        // Do nothing
      }
      try {
        await this.documentService.createFolder(versionFolder);
      } catch {
        // Do nothing
      }

      // Upload file and update metadata
      const fileResult = await this.sp.web
        .getFolderByServerRelativePath(versionFolder)
        .files.addUsingPath(file.name, file, { Overwrite: true });
      const item = await this.sp.web
        .getFileByServerRelativePath(fileResult.ServerRelativeUrl)
        .getItem();
      await item.update({
        Title: metadata.title || file.name,
        UniclassCode1: uniclassCodes.UniclassCode1 || "",
        UniclassTitle1: uniclassCodes.UniclassTitle1 || "",
        UniclassCode2: uniclassCodes.UniclassCode2 || "",
        UniclassTitle2: uniclassCodes.UniclassTitle2 || "",
        UniclassCode3: uniclassCodes.UniclassCode3 || "",
        UniclassTitle3: uniclassCodes.UniclassTitle3 || "",
        ProjectId: projectNumber,
      });
    } catch (error) {
      console.error("Error uploading native file:", error);
      throw error;
    }
  }
}
