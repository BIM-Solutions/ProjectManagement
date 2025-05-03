// ProjectForm.tsx
import * as React from 'react';
import { useContext, useEffect, useState } from 'react';
import { PeoplePicker, PrincipalType, IPeoplePickerContext } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { FilePicker, IFilePickerResult } from '@pnp/spfx-controls-react/lib/FilePicker';
import {
  TextField, Dropdown, IDropdownOption, Stack,
  PrimaryButton, MessageBar, MessageBarType
} from '@fluentui/react';
import { Image, ImageFit } from '@fluentui/react/lib/Image';
import { SPContext } from '../SPContext';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { Project } from '../services/ProjectSelectionServices';

const Debug = true; // Set to false in production

const sectorOptions: IDropdownOption[] = [
  { key: 'Public', text: 'Public' },
  { key: 'Defence', text: 'Defence' },
  { key: 'Power', text: 'Power' }
];

const statusOptions: IDropdownOption[] = [
  { key: 'Enquiry', text: 'Enquiry' },
  { key: 'Active', text: 'Active' },
  { key: 'Complete', text: 'Complete' },
  { key: 'Lost', text: 'Lost' },
  { key: 'Cancelled', text: 'Cancelled' },
  { key: 'Inactive', text: 'Inactive' }
];

interface IProjectFormProps {
  context: WebPartContext;
  mode: 'create' | 'edit';
  project?: Project;
  onSuccess?: () => void;
}

const ProjectForm: React.FC<IProjectFormProps> = ({ context, mode, project, onSuccess }) => {
  const sp = useContext(SPContext);
  const [title, setTitle] = useState('');
  const [projectName, setProjectName] = useState('');
  const [projectImage, setProjectImage] = useState<{ fileName: string; serverRelativeUrl: string } | null>(null);
  const [sector, setSector] = useState<string | undefined>();
  const [status, setStatus] = useState<string | undefined>();
  const [client, setClient] = useState('');
  const [manager, setManager] = useState('');
  const [pm, setPM] = useState('');
  const [checker, setChecker] = useState('');
  const [approver, setApprover] = useState('');
  const [description, setDescription] = useState('');
  const [subCodes, setSubCodes] = useState('');
  const [clientContact, setClientContact] = useState('');
  const [message, setMessage] = useState<{ type: MessageBarType; text: string } | null>(null);

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
      setMessage({ type: MessageBarType.error, text: 'Please fill in all required fields.' });
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
      if (Debug) {
        console.log('Debug mode: Project payload', payload);
        console.log('Debug mode: Project id', project?.id);
      }
      
      if (mode === 'edit' && project?.id) {
        await sp.web.lists.getByTitle("9719_ProjectInformationDatabase").items.getById(project.id).update(payload);
      } else {
        await sp.web.lists.getByTitle("9719_ProjectInformationDatabase").items.add(payload);
      }
      setMessage({ type: MessageBarType.success, text: `Project ${mode === 'edit' ? 'updated' : 'created'} successfully!` });
      if (onSuccess) onSuccess();
    
      } catch (err: unknown) {
        if (err instanceof Error) {
          console.error(err);
          setMessage({ type: MessageBarType.error, text: err.message || 'Submission failed.' });
        } else {
          console.error('Unknown error:', err);
          setMessage({ type: MessageBarType.error, text: 'An unknown error occurred.' });
        }
      }
  };

  return (
    <Stack tokens={{ childrenGap: 10, padding: 20 }}>
      <h2>{mode === 'edit' ? 'Edit Project' : 'Create New Project'}</h2>
      {message && <MessageBar messageBarType={message.type}>{message.text}</MessageBar>}
      <Stack horizontal tokens={{ childrenGap: 10 }}>
        <TextField label="Project Number" required value={title} onChange={(_, v) => setTitle(v || '')} styles={{ root: { flexGrow: 1 } }} />
        <TextField label="Project Name" required value={projectName} onChange={(_, v) => setProjectName(v || '')} styles={{ root: { flexGrow: 1 } }} />
      </Stack>
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
          src={`${context.pageContext.web.absoluteUrl}${projectImage.serverRelativeUrl}`}
          alt="Project Image"
          style={{ height: 200, width: 400 }}
        />
      )}
      <Stack horizontal tokens={{ childrenGap: 10 }}>
        <Dropdown label="Sector" options={sectorOptions} selectedKey={sector} required onChange={(_, o) => setSector(o?.key as string)} styles={{ root: { flexGrow: 1 } }} />
        <Dropdown label="Status" options={statusOptions} selectedKey={status} required onChange={(_, o) => setStatus(o?.key as string)} styles={{ root: { flexGrow: 1 } }} />
      </Stack>
      <TextField label="Client" value={client} onChange={(_, v) => setClient(v || '')} />
      <PeoplePicker context={peoplePickerContext} titleText="Information Manager" personSelectionLimit={1} principalTypes={[PrincipalType.User]} onChange={(items) => setManager(items[0]?.secondaryText || '')} />
      <PeoplePicker context={peoplePickerContext} titleText="Project Manager" personSelectionLimit={1} principalTypes={[PrincipalType.User]} onChange={(items) => setPM(items[0]?.secondaryText || '')} />
      <PeoplePicker context={peoplePickerContext} titleText="Checker" personSelectionLimit={1} principalTypes={[PrincipalType.User]} onChange={(items) => setChecker(items[0]?.secondaryText || '')} />
      <PeoplePicker context={peoplePickerContext} titleText="Approver" personSelectionLimit={1} principalTypes={[PrincipalType.User]} onChange={(items) => setApprover(items[0]?.secondaryText || '')} />
      <TextField label="Project Description" multiline rows={3} value={description} onChange={(_, v) => setDescription(v || '')} />
      <TextField label="Deltek SubCodes" multiline rows={2} value={subCodes} onChange={(_, v) => setSubCodes(v || '')} />
      <TextField label="Client Contact" multiline rows={2} value={clientContact} onChange={(_, v) => setClientContact(v || '')} />
      <PrimaryButton text={mode === 'edit' ? 'Update Project' : 'Create Project'} onClick={handleSubmit} />
    </Stack>
  );
};

export default ProjectForm;
