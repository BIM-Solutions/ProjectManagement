import { SPFI } from '@pnp/sp';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { IFilePickerResult } from '@pnp/spfx-controls-react';

export interface IStandard {
  Id: number;
  Title: string;
  Client: string;
  Version: string;
  UniclassCode: string;
  FileRef: string;
}

export class StandardsService {
  private sp: SPFI;
//   private context: WebPartContext;
  private readonly STANDARDS_LIBRARY = 'Standards';
  private readonly TEMPLATES_LIBRARY = 'Templates';

  constructor(context: WebPartContext, sp: SPFI) {
    // this.context = context;
    this.sp = sp;
  }

  public async uploadNewStandard(file: IFilePickerResult, projectNumber: string): Promise<void> {
    try {
      // Upload to Standards library
      const standardsFolder = `${this.STANDARDS_LIBRARY}/${projectNumber}`;
      await this.sp.web.getFolderByServerRelativePath(standardsFolder)
        .files.addUsingPath(file.fileName, await file.downloadFileContent(), { Overwrite: true });

      // Copy to project folder
      const projectFolder = `Projects/${projectNumber}/Standards`;
      await this.copyToProjectFolder(file.fileName, standardsFolder, projectFolder);

      // Update metadata
      await this.updateProjectMetadata(projectNumber, file.fileName);
    } catch (error) {
      console.error('Error uploading standard:', error);
      throw error;
    }
  }

  public async getExistingStandards(): Promise<IStandard[]> {
    try {
      const items = await this.sp.web.lists
        .getByTitle(this.STANDARDS_LIBRARY)
        .items
        .select('Id', 'Title', 'Client', 'Version', 'UniclassCode', 'FileRef')();

      
      return items as unknown as IStandard[];
    } catch (error) {
      console.error('Error fetching standards:', error);
      throw error;
    }
  }

  public async copyExistingStandard(standard: IStandard, projectNumber: string): Promise<void> {
    try {
      const projectFolder = `Projects/${projectNumber}/Standards`;
      await this.copyToProjectFolder(standard.Title, this.STANDARDS_LIBRARY, projectFolder);
      await this.updateProjectMetadata(projectNumber, standard.Title);
    } catch (error) {
      console.error('Error copying standard:', error);
      throw error;
    }
  }

  private async copyToProjectFolder(fileName: string, sourceFolder: string, targetFolder: string): Promise<void> {
    try {
      const sourceFile = await this.sp.web.getFileByServerRelativePath(`${sourceFolder}/${fileName}`);
      await sourceFile.copyTo(`${targetFolder}/${fileName}`, true);
    } catch (error) {
      console.error('Error copying file:', error);
      throw error;
    }
  }

  private async updateProjectMetadata(projectNumber: string, fileName: string): Promise<void> {
    try {
      const file = await this.sp.web.getFileByServerRelativePath(`Projects/${projectNumber}/Standards/${fileName}`);
      const metadata = await file.getItem();
      const metadataValues = await metadata();
      
      // Update project's report metadata
      await this.sp.web.lists
        .getByTitle('Projects')
        .items
        .getById(parseInt(projectNumber))
        .update({
          StandardClient: metadataValues.Client,
          StandardVersion: metadataValues.Version,
          StandardUniclass: metadataValues.UniclassCode
        });
    } catch (error) {
      console.error('Error updating metadata:', error);
      throw error;
    }
  }

  public async copyTemplates(projectNumber: string, uniclassCode: string): Promise<void> {
    try {
      const templates = await this.sp.web.lists
        .getByTitle(this.TEMPLATES_LIBRARY)
        .items
        .filter(`UniclassCode eq '${uniclassCode}'`)();


      for (const template of templates) {
        const sourceFile = await this.sp.web.getFileByServerRelativePath(template.FileRef);
        const newFileName = `${projectNumber}_${uniclassCode}_${template.Title}.docx`;
        await sourceFile.copyTo(`Projects/${projectNumber}/Reports/${newFileName}`, true);
      }
    } catch (error) {
      console.error('Error copying templates:', error);
      throw error;
    }
  }
}