import * as React from 'react';
import { useContext, useState } from 'react';
import { PeoplePicker, PrincipalType, IPeoplePickerContext} from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { FilePicker, IFilePickerResult } from '@pnp/spfx-controls-react/lib/FilePicker';
import {
  TextField, Dropdown, IDropdownOption, Stack,
  PrimaryButton, MessageBar, MessageBarType
} from '@fluentui/react';
import { Image, ImageFit } from '@fluentui/react/lib/Image';
import { SPContext } from '../../common/SPContext';
import { WebPartContext } from '@microsoft/sp-webpart-base';

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

// import { SPFI } from '@pnp/sp';

// /**
//  * Ensures that a user exists in SharePoint and returns their ID.
//  * @param sp The SharePoint context to use.
//  * @param email The email address of the user to ensure.
//  * @returns A promise that resolves with the ID of the user in SharePoint.
//  */
// const ensureUserId = async (sp: SPFI, email: string): Promise<number> => {
//   const user = await sp.web.ensureUser(email);
//   return user.Id;
// };

interface ICreateProjectFormProps {
  onSuccess?: () => void; // Callback function to be called on successful project creation
  context: WebPartContext; // Context from SharePoint
}

/**
 * CreateProjectForm component allows users to create a new project by filling out a form.
 * The form includes fields for project name, project number, sector, status, client,
 * project manager, manager, checker, approver, project description, Deltek subcodes, 
 * and client contact.
 * 
 * The component uses Fluent UI controls for input fields and displays messages upon 
 * successful creation or errors.
 * 
 * Props:
 * - onSuccess: An optional callback function executed upon successful project creation.
 * 
 * The component interacts with the SharePoint context to add new project items to 
 * the "9719_ProjectInformationDatabase" list.
 */

const CreateProjectForm: React.FC<ICreateProjectFormProps> = ({ onSuccess, context }) => {
  const sp = useContext(SPContext);

  const [title, setProjectNumber] = useState('');
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

  /**
   * Ensures that a user exists in SharePoint and returns their ID.
   * If the user does not exist or there is an error, it returns null.
   * @param email The email address of the user to ensure.
   * @returns A promise that resolves with the ID of the user in SharePoint, or null.
   */
  const safeEnsureUserId = async (email: string): Promise<number | null> => {
    if (!email) return null;
    try {
      const user = await sp.web.ensureUser(email);
      return user.Id;
    } catch (e) {
      console.warn(`Could not resolve user: ${email}`, e);
      return null;
    }
  };

  /**
   * Handles the form submission by adding a new project item to the
   * "9719_ProjectInformationDatabase" list in SharePoint.
   * If the submission is successful, it resets the form fields and
   * displays a success message. If there are any errors, it displays
   * an error message.
   * If the onSuccess prop is provided, it calls the callback function
   * upon successful project creation.
   */
  const handleSubmit = async (): Promise<void> => {
    if (!title || !projectName || !sector || !status) {
      setMessage({ type: MessageBarType.error, text: 'Please fill in all required fields.' });
      return;
    }

    const pmId = await safeEnsureUserId(pm);
    const managerId = await safeEnsureUserId(manager);
    const checkerId = await safeEnsureUserId(checker);
    const approverId = await safeEnsureUserId(approver);
    
    try {
      // Add the new project item to the SharePoint list
      await sp.web.lists.getByTitle("9719_ProjectInformationDatabase").items.add({
        Title: title,
        ProjectName: projectName,
        Sector: sector,
        Status: status,
        Client: client,
        PMId: pmId,
        ManagerId: managerId,
        CheckerId: checkerId,
        ApproverId: approverId,
        ProjectDescription: description,
        DeltekSubCodes: subCodes,
        ClientContact: clientContact,
        ProjectImage: projectImage ? JSON.stringify(projectImage) : undefined

      });
      
      setMessage({ type: MessageBarType.success, text: 'Project created successfully!' });
      setProjectNumber('');
      setProjectName('');
      setSector(undefined);
      setStatus(undefined);
      setClient('');
      setManager('');
      setPM('');
      setChecker('');
      setApprover('');
      setDescription('');
      setSubCodes('');
      setClientContact('');
      setProjectImage(null);

      if (onSuccess) onSuccess();
    } catch (err) {
      console.error(err);
      setMessage({ type: MessageBarType.error, text: err.message || 'An error occurred while creating the project.' });
    }
  };

  return (
    <Stack tokens={{ childrenGap: 10, padding: 20 }} >
      <h2>Create New Project</h2>
      {message && <MessageBar messageBarType={message.type}>{message.text}</MessageBar>}
      <Stack horizontal tokens={{ childrenGap: 10 }}>
        <TextField label="Project Number" required value={title} onChange={(_, v) => setProjectNumber(v || '')} styles={{ root: { flexGrow: 1 } }} />
        <TextField label="Project Name" required value={projectName} onChange={(_, v) => setProjectName(v || '')} styles={{ root: { flexGrow: 1 } }}/>
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

          try {
            const blob: Blob = await result.downloadFileContent();

            const arrayBuffer = await blob.arrayBuffer(); // ← native API

            const upload = await sp.web
              .getFolderByServerRelativePath("/sites/DevelopmentSite/SiteAssets/ProjectImages")
              .files.addUsingPath(result.fileName, arrayBuffer, { Overwrite: true });

            const serverRelativeUrl = upload.ServerRelativeUrl;

            setProjectImage({
              fileName: result.fileName,
              serverRelativeUrl
            });
          } catch (err) {
            console.error("Image upload failed", err);
            setMessage({ type: MessageBarType.error, text: "Image upload failed." });
          }
        }}
        onCancel={() => console.log("Cancelled image selection")}
      />


      {projectImage && (
        <Image
          imageFit={ImageFit.contain}
          src={`${context.pageContext.web.absoluteUrl}${projectImage.serverRelativeUrl}`}
          alt="Uploaded Project"
          style={{ height: 200, width: 400, marginTop: 10 }}
        />
      )}
      <Stack horizontal tokens={{ childrenGap: 10 }}>
        <Dropdown label="Sector" options={sectorOptions} selectedKey={sector} required onChange={(_, o) => setSector(o?.key as string)} styles={{ root: { flexGrow: 1 } }}/>
        <Dropdown label="Status" options={statusOptions} selectedKey={status} required onChange={(_, o) => setStatus(o?.key as string)} styles={{ root: { flexGrow: 1 } }}/>
      </Stack>
      <TextField label="Client" value={client} onChange={(_, v) => setClient(v || '')} />
      <PeoplePicker
        context={peoplePickerContext}
        titleText="Information Manager"
        personSelectionLimit={1}
        principalTypes={[PrincipalType.User]}
        onChange={(items) => setManager(items[0]?.secondaryText || '')}
        required={false}
      />
      <PeoplePicker
        context={peoplePickerContext}
        titleText="Project Manager"
        personSelectionLimit={1}
        principalTypes={[PrincipalType.User]}
        onChange={(items) => setPM(items[0]?.secondaryText || '')}
        required={false}
      />
      <PeoplePicker
        context={peoplePickerContext}
        titleText="Checker"
        personSelectionLimit={1}
        principalTypes={[PrincipalType.User]}
        onChange={(items) => setChecker(items[0]?.secondaryText || '')}
        required={false}
      />
      <PeoplePicker
        context={peoplePickerContext}
        titleText="Approver"
        personSelectionLimit={1}
        principalTypes={[PrincipalType.User]}        
        onChange={(items) => setApprover(items[0]?.secondaryText || '')}
        required={false}
      />
      <TextField label="Project Description" multiline rows={3} value={description} onChange={(_, v) => setDescription(v || '')} />
      <TextField label="Deltek SubCodes" multiline rows={2} value={subCodes} onChange={(_, v) => setSubCodes(v || '')} />
      <TextField label="Client Contact" multiline rows={2} value={clientContact} onChange={(_, v) => setClientContact(v || '')} />

      <PrimaryButton text="Create Project" onClick={handleSubmit} />
    </Stack>
  );
};

export default CreateProjectForm;
