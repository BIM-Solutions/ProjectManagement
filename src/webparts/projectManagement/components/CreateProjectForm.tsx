import * as React from 'react';
import { useContext, useState } from 'react';
import { TextField, Dropdown, IDropdownOption, Stack, PrimaryButton, MessageBar, MessageBarType } from '@fluentui/react';
import { SPContext } from '../SPContext';

const sectorOptions: IDropdownOption[] = [
  { key: 'Commercial', text: 'Commercial' },
  { key: 'Residential', text: 'Residential' },
  { key: 'Infrastructure', text: 'Infrastructure' },
];

interface ICreateProjectFormProps {
    onSuccess?: () => void;
  }

const CreateProjectForm: React.FC<ICreateProjectFormProps> = ({ onSuccess }) => {
  const sp = useContext(SPContext);
  const [title, setTitle] = useState('');
  const [projectNumber, setProjectNumber] = useState('');
  const [sector, setSector] = useState<string | undefined>();
  const [keyPersonnel, setKeyPersonnel] = useState('');
  const [message, setMessage] = useState<{ type: MessageBarType; text: string } | null>(null);

  const handleSubmit = async (): Promise<void> => {
    if (!title || !projectNumber || !sector) {
      setMessage({ type: MessageBarType.error, text: 'Please fill in all required fields.' });
      return;
    }

    try {
      await sp.web.lists.getByTitle("ProjectInformation").items.add({
        Title: title,
        ProjectNumber: projectNumber,
        Sector: sector,
        KeyPersonnel: keyPersonnel
      });

      setMessage({ type: MessageBarType.success, text: 'Project created successfully!' });
      setTitle('');
      setProjectNumber('');
      setSector(undefined);
      setKeyPersonnel('');

      if (onSuccess) {
        onSuccess();
      }

    } catch (err: unknown) {
      console.error(err);
      setMessage({ type: MessageBarType.error, text: 'Failed to create project.' });
    }
  };

  return (
    <Stack tokens={{ childrenGap: 15, padding: 20 }} styles={{ root: { maxWidth: 500 } }}>
      <h2>Create New Project</h2>

      {message && <MessageBar messageBarType={message.type}>{message.text}</MessageBar>}

      <TextField label="Project Name" required value={title} onChange={(_, v) => setTitle(v || '')} />
      <TextField label="Project Number" required value={projectNumber} onChange={(_, v) => setProjectNumber(v || '')} />
      <Dropdown
        label="Sector"
        options={sectorOptions}
        selectedKey={sector}
        required
        onChange={(_, option) => setSector(option?.key as string)}
      />
      <TextField label="Key Personnel" value={keyPersonnel} onChange={(_, v) => setKeyPersonnel(v || '')} />

      <PrimaryButton text="Create Project" onClick={handleSubmit} />
    </Stack>
  );
};

export default CreateProjectForm;
