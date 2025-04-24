import * as React from 'react';
import { useContext, useState } from 'react';
import { PeoplePicker, PrincipalType, IPeoplePickerContext} from "@pnp/spfx-controls-react/lib/PeoplePicker";

import {
  TextField, Dropdown, IDropdownOption, Stack,
  PrimaryButton, MessageBar, MessageBarType
} from '@fluentui/react';
import { SPContext } from '../SPContext';
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

import { SPFI } from '@pnp/sp';

/**
 * Ensures that a user exists in SharePoint and returns their ID.
 * @param sp The SharePoint context to use.
 * @param email The email address of the user to ensure.
 * @returns A promise that resolves with the ID of the user in SharePoint.
 */
const ensureUserId = async (sp: SPFI, email: string): Promise<number> => {
  const user = await sp.web.ensureUser(email);
  return user.Id;
};

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

    try {
      const pmId = await ensureUserId(sp, pm);
      const managerId = await ensureUserId(sp, manager);
      const checkerId = await ensureUserId(sp, checker);
      const approverId = await ensureUserId(sp, approver);
    
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
        ClientContact: clientContact
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

      if (onSuccess) onSuccess();
    } catch (err) {
      console.error(err);
      setMessage({ type: MessageBarType.error, text: 'Failed to create project.' });
    }
  };

  return (
    <Stack tokens={{ childrenGap: 10, padding: 20 }} styles={{ root: { maxWidth: 600 } }}>
      <h2>Create New Project</h2>
      {message && <MessageBar messageBarType={message.type}>{message.text}</MessageBar>}

      <TextField label="Project Number" required value={title} onChange={(_, v) => setProjectNumber(v || '')} />
      <TextField label="Project Name" required value={projectName} onChange={(_, v) => setProjectName(v || '')} />
      <Dropdown label="Sector" options={sectorOptions} selectedKey={sector} required onChange={(_, o) => setSector(o?.key as string)} />
      <Dropdown label="Status" options={statusOptions} selectedKey={status} required onChange={(_, o) => setStatus(o?.key as string)} />
      <TextField label="Client" value={client} onChange={(_, v) => setClient(v || '')} />
      <PeoplePicker
        context={peoplePickerContext}
        titleText="Information Manager"
        personSelectionLimit={1}
        principalTypes={[PrincipalType.User]}
        onChange={(items) => setManager(items[0]?.secondaryText || '')}
        required={false}
      />
      <TextField label="Project Manager" value={pm} onChange={(_, v) => setPM(v || '')} />
      <TextField label="Manager" value={manager} onChange={(_, v) => setManager(v || '')} />
      <TextField label="Checker" value={checker} onChange={(_, v) => setChecker(v || '')} />
      <TextField label="Approver" value={approver} onChange={(_, v) => setApprover(v || '')} />
      <TextField label="Project Description" multiline rows={3} value={description} onChange={(_, v) => setDescription(v || '')} />
      <TextField label="Deltek SubCodes" multiline rows={2} value={subCodes} onChange={(_, v) => setSubCodes(v || '')} />
      <TextField label="Client Contact" multiline rows={2} value={clientContact} onChange={(_, v) => setClientContact(v || '')} />

      <PrimaryButton text="Create Project" onClick={handleSubmit} />
    </Stack>
  );
};

export default CreateProjectForm;
