// ProjectForm.tsx
import * as React from 'react';
import { useContext, useEffect, useState } from 'react';
import { PeoplePicker, PrincipalType, IPeoplePickerContext } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { FilePicker, IFilePickerResult } from '@pnp/spfx-controls-react/lib/FilePicker';
import {
  Button,
   Combobox, 
   Field, 
   Input, 
   MessageBar,
   MessageBarIntent,
   Textarea,
   makeStyles,
   Option
  //  useId
} from '@fluentui/react-components';
import { Image, ImageFit } from '@fluentui/react/lib/Image';
import { SPContext } from './SPContext';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { Project, ProjectSelectionService } from '../../services/ProjectSelectionServices';
import { DEBUG } from './DevVariables';
import { EventService } from '../../services/EventService';
import { projectStatusOptions, sectorOptions } from '../..//services/ListService';


interface IProjectFormProps {
  context: WebPartContext;
  mode: 'create' | 'edit';
  project?: Project;
  onSuccess?: () => void;
  onCancel?: () => void;
}

const useStyles = makeStyles({
  container: {
    gap: '10',
    childrenGap: '10', 
    padding: '20'
  },
  rowArea: {
    gap: '10'
  },
  button: {
    marginRight: '10'
  },
  buttonRow: {
    display: 'flex',
    justifyContent: 'center',
    gap: '16px',
    marginTop: '24px',
  },
  
  labelsRow: {
    gap: '10'
  }


})

const ProjectForm: React.FC<IProjectFormProps> = ({ context, mode, project, onSuccess, onCancel }) => {
  const sp = useContext(SPContext);
  const [title, setTitle] = useState('');
  const [projectName, setProjectName] = useState('');
  const [projectImage, setProjectImage] = useState<{ fileName: string; serverRelativeUrl: string } | null>(null);
  const [sector, setSector] = useState<string>('');
  const [status, setStatus] = useState<string>('');
  const [client, setClient] = useState('');
  const [manager, setManager] = useState('');
  const [pm, setPM] = useState('');
  const [checker, setChecker] = useState('');
  const [approver, setApprover] = useState('');
  const [description, setDescription] = useState('');
  const [subCodes, setSubCodes] = useState('');
  const [clientContact, setClientContact] = useState('');
  const [message, setMessage] = useState<{ type: MessageBarIntent; text: string } | null>(null);
  const styles = useStyles();
  // const dropdownId = useId("dropdown");

  const peoplePickerContext: IPeoplePickerContext = {
    spHttpClient: context.spHttpClient,
    msGraphClientFactory: context.msGraphClientFactory,
    absoluteUrl: context.pageContext.web.absoluteUrl
  };

  useEffect(() => {
    if (mode === 'edit' && project) {
      setTitle(project.ProjectNumber || '');
      setProjectName(project.ProjectName || '');
      setSector(project.Sector);
      setStatus(project.Status);
      setClient(project.Client);
      setManager(project.Manager?.Email || '');
      setPM(project.PM?.Email || '');
      setChecker(project.Checker?.Email || '');
      setApprover(project.Approver?.Email || '');
      setDescription(project.ProjectDescription || '');
      setSubCodes(project.DeltekSubCodes || '');
      setClientContact(project.ClientContact || '');
      try {
        if (project.ProjectImage) {
          const parsed = JSON.parse(project.ProjectImage);
          setProjectImage(parsed);
        }
      } catch {
        // Intentionally ignored
      }
    }
  }, [mode, project]);

  const safeEnsureUserId = async (email: string): Promise<number | null> => {
    if (!email) return null;
    try {
      const user = await sp.web.ensureUser(email);
      return user.Id;
    } catch {
      return null;
    }
  };
  interface ProjectPayload {
    Title: string;
    ProjectName: string;
    Sector: string;
    Status: string;
    Client: string;
    PMId: number | null;
    ManagerId: number | null;
    CheckerId: number | null;
    ApproverId: number | null;
    ProjectDescription: string;
    DeltekSubCodes: string;
    ClientContact: string;
    ProjectImage?: string;
  }  

  const handleSubmit = async (): Promise<void> => {
    if (!title || !projectName || !sector || !status) {
      setMessage({ type: 'error', text: 'Please fill in all required fields.' });
      return;
    }

    const payload:  ProjectPayload = {
      Title: title,
      ProjectName: projectName,
      Sector: sector,
      Status: status,
      Client: client,
      PMId: await safeEnsureUserId(pm),
      ManagerId: await safeEnsureUserId(manager),
      CheckerId: await safeEnsureUserId(checker),
      ApproverId: await safeEnsureUserId(approver),
      ProjectDescription: description,
      DeltekSubCodes: subCodes,
      ClientContact: clientContact,
      ProjectImage: projectImage ? JSON.stringify(projectImage) : undefined
    };

    try {
      if (!DEBUG) {
        console.log('Debug mode: Project payload', payload);
        console.log('Debug mode: Project id', project?.id);
      }
      
      if (mode === 'edit' && project?.id) {
        await sp.web.lists.getByTitle("9719_ProjectInformationDatabase").items.getById(project.id).update(payload);
        const updatedProject = await sp.web.lists
          .getByTitle("9719_ProjectInformationDatabase")
          .items
          .getById(project.id)
          .expand("PM", "Manager", "Checker", "Approver")
          .select(
            "Id", "Title", "ProjectName", "Status", "Sector", "Client", "ProjectImage", "ProjectDescription", "DeltekSubCodes", "ClientContact",
            "PM/Id", "PM/Title", "PM/EMail", "PM/JobTitle", "PM/Department",
            "Manager/Id", "Manager/Title", "Manager/EMail", "Manager/JobTitle", "Manager/Department",
            "Checker/Id", "Checker/Title", "Checker/EMail", "Checker/JobTitle", "Checker/Department",
            "Approver/Id", "Approver/Title", "Approver/EMail", "Approver/JobTitle", "Approver/Department"
          )
          ();

        ProjectSelectionService.setSelectedProject({
          id: updatedProject.Id,
          ProjectNumber: updatedProject.Title,
          ProjectName: updatedProject.ProjectName,
          Status: updatedProject.Status,
          Sector: updatedProject.Sector,
          Client: updatedProject.Client,
          ProjectImage: updatedProject.ProjectImage,
          ProjectDescription: updatedProject.ProjectDescription || '',
          DeltekSubCodes: updatedProject.DeltekSubCodes || '',
          ClientContact: updatedProject.ClientContact || '',
          PM: updatedProject.PM ? {
            Id: updatedProject.PM.Id,
            Title: updatedProject.PM.Title,
            Email: updatedProject.PM.EMail,
            JobTitle: updatedProject.PM.JobTitle,
            Department: updatedProject.PM.Department
          } : undefined,
          Manager: updatedProject.Manager ? {
            Id: updatedProject.Manager.Id,
            Title: updatedProject.Manager.Title,
            Email: updatedProject.Manager.EMail,
            JobTitle: updatedProject.Manager.JobTitle,
            Department: updatedProject.Manager.Department
          } : undefined,
          Checker: updatedProject.Checker ? {
            Id: updatedProject.Checker.Id,
            Title: updatedProject.Checker.Title,
            Email: updatedProject.Checker.EMail,
            JobTitle: updatedProject.Checker.JobTitle,
            Department: updatedProject.Checker.Department
          } : undefined,
          Approver: updatedProject.Approver ? {
            Id: updatedProject.Approver.Id,
            Title: updatedProject.Approver.Title,
            Email: updatedProject.Approver.EMail,
            JobTitle: updatedProject.Approver.JobTitle,
            Department: updatedProject.Approver.Department
          } : undefined
        });


      } else {
        await sp.web.lists.getByTitle("9719_ProjectInformationDatabase").items.add(payload);
        
      }
      setMessage({ type: 'success', text: `Project ${mode === 'edit' ? 'updated' : 'created'} successfully!` });
      if (onSuccess) onSuccess();
      EventService.publishProjectUpdated();
    
      } catch (err: unknown) {
        if (err instanceof Error) {
          console.error(err);
          setMessage({ type: 'error', text: err.message || 'Submission failed.' });
        } else {
          console.error('Unknown error:', err);
          setMessage({ type: 'error', text: 'An unknown error occurred.' });
        }
      }
  };
  

  return (
    <div className={styles.container}>
      <h2>{mode === 'edit' ? 'Edit Project' : 'Create New Project'}</h2>
      {message && <MessageBar intent={message.type}>{message.text}</MessageBar>}
      <div className={styles.labelsRow}>
      <Field label="Project Number" required>
        <Input
          value={title}
          onChange={(event) => setTitle((event.target as HTMLInputElement).value || '')}
          style={{ flexGrow: 1 }}
        />
      </Field>
      <Field label="Project Name" required>
        <Input
          value={projectName}
          onChange={(event) => setProjectName((event.target as HTMLInputElement).value || '')}
          style={{ flexGrow: 1}}
          />
      </Field>
      </div >
      <div style={{ marginTop: 16 }}>
        <FilePicker
          context={context}
          buttonLabel="Upload Project Image"
          hideStockImages={true}
          hideLinkUploadTab={true}
          hideOneDriveTab={true}
          hideSiteFilesTab={false}
          onSave={async (results: IFilePickerResult[]) => {
            const result = results[0];
            const blob = await result.downloadFileContent();
            const arrayBuffer = await blob.arrayBuffer();
            const upload = await sp.web.getFolderByServerRelativePath("/sites/DevelopmentSite/SiteAssets/ProjectImages")
              .files.addUsingPath(result.fileName, arrayBuffer, { Overwrite: true });
            const serverRelativeUrl = upload.ServerRelativeUrl;
            setProjectImage({ fileName: result.fileName, serverRelativeUrl });
          }}
          onCancel={() => console.log("Image selection cancelled")}
          
        />
        {projectImage && (
          <Image
            imageFit={ImageFit.contain}
            src={projectImage.serverRelativeUrl}
            alt={projectImage.fileName}
            style={{ height: 200, width: 'auto', marginTop: 16 }}
          />
      )}
      </div>
      <div className={styles.rowArea}>
      <Field label="Sector" required>
        <Combobox
          value={sector}
          onOptionSelect={(_, data) => setSector(data.optionText || '')}
          placeholder="Select sector"
        >
          {sectorOptions.map((text) => (
            <Option key={text.key} text={text.value}>
              {text.value}
            </Option>
          ))}
        </Combobox>
      </Field>
      <Field label="Status" required>
        <Combobox
          value={status}
          onOptionSelect={(_, data) => setStatus(data.optionText || '')}
          placeholder="Select status"
        >
          {projectStatusOptions.map((text) => (
            <Option key={text.key} text={text.value}>
              {text.value}
            </Option>
          ))}
        </Combobox>
      </Field>

      </div>
      <Field label="Client" required>
        <Input
          value={client}
          onChange={(event) => setClient((event.target as HTMLInputElement).value || '')}
          style={{ flexGrow: 1}}
          />
      </Field>
      <PeoplePicker context={peoplePickerContext} titleText="Information Manager" personSelectionLimit={1} defaultSelectedUsers={[manager]} principalTypes={[PrincipalType.User]} onChange={(items) => setManager(items[0]?.secondaryText || '')} />
      <PeoplePicker context={peoplePickerContext} titleText="Project Manager" personSelectionLimit={1} defaultSelectedUsers={[pm]} principalTypes={[PrincipalType.User]} onChange={(items) => setPM(items[0]?.secondaryText || '')} />
      <PeoplePicker context={peoplePickerContext} titleText="Checker" personSelectionLimit={1} defaultSelectedUsers={[checker]} principalTypes={[PrincipalType.User]} onChange={(items) => setChecker(items[0]?.secondaryText || '')} />
      <PeoplePicker context={peoplePickerContext} titleText="Approver" personSelectionLimit={1} defaultSelectedUsers={[approver]}  principalTypes={[PrincipalType.User]} onChange={(items) => setApprover(items[0]?.secondaryText || '')} />
      <Field label="Project Decription" required>
        <Textarea
          value={description}
          onChange={(event) => setDescription((event.target as HTMLTextAreaElement).value || '')}
          style={{ flexGrow: 1}}
          />
      </Field>
      <Field label="Deltek SubCodes" required>
        <Textarea
          value={subCodes}
          onChange={(event) => setSubCodes((event.target as HTMLTextAreaElement).value || '')}
          style={{ flexGrow: 1}}
          />
      </Field>
      <Field label="Client Contact" required>
        <Input
          value={clientContact}
          onChange={(event) => setClientContact((event.target as HTMLInputElement).value || '')}
          style={{ flexGrow: 1}}
          />
      </Field>
      <div className={styles.buttonRow}>
        {mode !== 'create' && <Button appearance='primary' onClick={onCancel} className={styles.button}>Cancel</Button>}
        <Button appearance='primary' onClick={() => {
          setTitle('');
          setProjectName('');
          setSector('');
          setStatus('');
          setClient('');
          setManager('');
          setPM('');
          setChecker('');
          setApprover('');
          setDescription('');
          setSubCodes('');
          setClientContact('');
          setProjectImage(null);
        }} className={styles.button}>Reset</Button>
      <Button appearance='primary'  onClick={handleSubmit} >{mode === 'edit' ? 'Update Project' : 'Create Project'}</Button>
      </div>
    </div>
  );
};

export default ProjectForm;
