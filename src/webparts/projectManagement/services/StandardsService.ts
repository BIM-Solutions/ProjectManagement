// Import required dependencies
import { SPFI } from "@pnp/sp";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { IFilePickerResult } from "@pnp/spfx-controls-react";
import { DocumentService } from "./DocumentService";

interface IFile {
  Name: string;
}

// Interface defining the structure of a Standard item
export interface IStandard {
  Id: number;
  Title: string;
  Client: string;
  Version: string;
  uniclassCode: string;
  FileRef: string;
  UniclassCode1?: string;
  UniclassTitle1?: string;
  UniclassCode2?: string;
  UniclassTitle2?: string;
  UniclassCode3?: string;
  UniclassTitle3?: string;
  File?: IFile;
}

// Interface defining the structure of a Uniclass item with nested levels
interface IUniclassItem {
  code: string;
  title: string;
  subgroups?: IUniclassItem[];
  sections?: IUniclassItem[];
}

// Service class for managing standards-related operations
export class StandardsService {
  private sp: SPFI;
  private documentService: DocumentService;
  private readonly STANDARDS_LIBRARY = "Standards";
  private readonly TEMPLATES_LIBRARY = "Templates";

  // Initialize the service with SharePoint context and SPFI instance
  constructor(context: WebPartContext, sp: SPFI) {
    this.sp = sp;
    this.documentService = new DocumentService(context);
  }

  // Upload new standards file to SharePoint
  public async uploadNewStandards(
    file: IFilePickerResult,
    nextVersion: string,
    clientName: string,
    projectNumber: string
  ): Promise<void> {
    try {
      // Ensure the Standards library exists
      await this.documentService.ensureDocumentLibrary(this.STANDARDS_LIBRARY);

      // Create folder paths
      const clientFolder = `${this.STANDARDS_LIBRARY}/${clientName}`;
      const versionFolder = `${clientFolder}/${nextVersion}`;

      // Create necessary folders
      await this.documentService.createFolder(clientFolder);
      await this.documentService.createFolder(versionFolder);

      // Upload the file to the version folder
      await this.sp.web
        .getFolderByServerRelativePath(versionFolder)
        .files.addUsingPath(file.fileName, await file.downloadFileContent(), {
          Overwrite: true,
        });
    } catch (error) {
      console.error("Error uploading standard:", error);
      throw error;
    }
  }

  // Get the current version for a client's standards
  public async GetCurrentVersion(clientName: string): Promise<string> {
    const clientFolder = `${this.STANDARDS_LIBRARY}/${clientName}`;
    const folders = await this.sp.web
      .getFolderByServerRelativePath(clientFolder)
      .folders();
    // Sort versions in descending order
    const versions = folders
      .map((f) => f.Name)
      .sort((a, b) => {
        const aNum = parseInt(a.replace(/^V/, ""));
        const bNum = parseInt(b.replace(/^V/, ""));
        return bNum - aNum;
      });

    return versions[0];
  }

  // Calculate the next version number based on current version
  public async getNextVersion(clientName: string): Promise<string> {
    const currentVersion = await this.GetCurrentVersion(clientName);

    // If no current version exists, start with V1
    if (!currentVersion) {
      return "V1";
    }

    // Extract prefix and number from current version
    const prefix = currentVersion.match(/^[A-Za-z]+/)?.[0] || "";
    const num = parseInt(currentVersion.replace(/^[A-Za-z]+/, ""));

    if (isNaN(num)) {
      return "V1";
    }

    const nextNum = num + 1;
    // Special handling for 'P' prefix (pad with zeros)
    if (prefix === "P") {
      return `${prefix}${nextNum.toString().padStart(2, "0")}`;
    }
    return `${prefix}${nextNum}`;
  }

  // Retrieve all existing standards from SharePoint
  public async getExistingStandards(): Promise<IStandard[]> {
    try {
      await this.documentService.ensureDocumentLibrary(this.STANDARDS_LIBRARY);

      // Fetch items with selected fields
      const items = await this.sp.web.lists
        .getByTitle(this.STANDARDS_LIBRARY)
        .items.select(
          "Id",
          "Title",
          "FileRef",
          "File/Name",
          "UniclassCode1",
          "UniclassTitle1",
          "UniclassCode2",
          "UniclassTitle2",
          "UniclassCode3",
          "UniclassTitle3"
        )
        .expand("File")();

      // Map items and extract client/version from FileRef
      console.log("items", items);
      return items
        .map((item) => ({
          ...item,
          Client:
            item.FileRef.split("/")[
              item.FileRef.split("/").indexOf(this.STANDARDS_LIBRARY) + 1
            ],
          Version:
            item.FileRef.split("/")[
              item.FileRef.split("/").indexOf(this.STANDARDS_LIBRARY) + 2
            ],
        }))
        .filter(
          (item) => !!item.Title && !!item.UniclassCode1
        ) as unknown as IStandard[];
    } catch (error) {
      console.error("Error fetching standards:", error);
      throw error;
    }
  }

  // Copy a standard to a project folder
  public async copyExistingStandard(
    standard: IStandard,
    projectNumber: string
  ): Promise<void> {
    try {
      const sourceFolder = `${this.STANDARDS_LIBRARY}/${standard.Client}/${standard.Version}`;
      const projectFolder = `Projects/${projectNumber}/Standards`;

      // Create folder structure using Uniclass hierarchy
      const uniclassFolder = `${projectFolder}/${standard.UniclassCode1}/${standard.UniclassCode2}/${standard.UniclassCode3}`;
      await this.documentService.createFolder(uniclassFolder);

      // Copy file and update metadata
      await this.copyToProjectFolder(
        standard.Title,
        sourceFolder,
        uniclassFolder
      );
      await this.updateProjectMetadata(projectNumber, standard.Title);
    } catch (error) {
      console.error("Error copying standard:", error);
      throw error;
    }
  }

  // Helper method to copy a file to project folder
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

  // Update project metadata after copying standards
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

      // Update project list item with standard metadata
      await this.sp.web.lists
        .getByTitle("Projects")
        .items.getById(parseInt(projectNumber))
        .update({
          StandardClient: metadataValues.Client,
          StandardVersion: metadataValues.Version,
          StandardUniclass: metadataValues.UniclassCode,
          StandardUniclass1: metadataValues.UniclassCode1,
          StandardUniclass2: metadataValues.UniclassCode2,
          StandardUniclass3: metadataValues.UniclassCode3,
        });
    } catch (error) {
      console.error("Error updating metadata:", error);
      throw error;
    }
  }

  // Copy templates based on Uniclass code
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

  // Add multiple standards to a project with progress tracking
  public async addStandardsToProjectWithProgress(
    files: {
      fileName: string;
      uniclassCode: string;
      client: string;
      version: string;
    }[],
    projectNumber: string,
    context: WebPartContext,
    sp: SPFI,
    onProgress?: (
      progress: number,
      summary: { fileName: string; targetPath: string }[]
    ) => void
  ): Promise<{ fileName: string; targetPath: string }[]> {
    const summary: { fileName: string; targetPath: string }[] = [];
    console.log("projectNumber", projectNumber);
    // Process each file
    for (let i = 0; i < files.length; i++) {
      const { fileName, uniclassCode, client, version } = files[i];
      console.log(uniclassCode);
      // Get Uniclass codes and create target folder
      const codes = await this.mapUniclassCodes(uniclassCode, context);
      const targetFolder = `Projects/${projectNumber}/${codes.UniclassTitle1}/${codes.UniclassTitle2}`;
      await this.documentService.createFolder(targetFolder);

      // Get web data and construct paths
      const webData = await this.sp.web.select("ServerRelativeUrl")();
      const siteRoot = webData.ServerRelativeUrl.replace(/\/$/, "");
      const sourcePath = `${siteRoot}/${this.STANDARDS_LIBRARY}/${client}/${version}/${fileName}`;
      const sourceFile = this.sp.web.getFileByServerRelativePath(sourcePath);
      const targetPath = `${siteRoot}/${targetFolder}/${fileName}`
        .replace(/\\/g, "/")
        .replace(/\/{2,}/g, "/");
      const relativeFolderPath = targetFolder.replace(`Projects/`, "");

      // Create folder hierarchy
      const folderParts = relativeFolderPath.split("/");
      let currentPath = "";
      for (const part of folderParts) {
        if (currentPath === "") {
          currentPath = part;
        } else {
          currentPath += "/" + part;
        }
        try {
          await sp.web.folders.addUsingPath(`Projects/${currentPath}`);
        } catch {
          // Do nothing
        }
      }

      // Copy file to target location
      try {
        await sourceFile.copyByPath(targetPath, false, {
          KeepBoth: false,
          ResetAuthorAndCreatedOnCopy: true,
          ShouldBypassSharedLocks: false,
        });

        // Get the copied file and update its metadata
        const targetFile = await sp.web.getFileByServerRelativePath(targetPath);
        const listItem = await targetFile.getItem();
        await listItem.update({
          ProjectId: projectNumber,
        });
      } catch (error) {
        console.error("Error copying file:", error);
        throw error;
      }

      // Update progress
      summary.push({ fileName, targetPath });
      if (onProgress) {
        onProgress((i + 1) / files.length, [...summary]);
      }
    }
    return summary;
  }
}
