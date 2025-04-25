import * as React from 'react';
import { useEffect, useState, useContext } from 'react';
import {
    DetailsList, IColumn, Stack, Text, TextField, Dropdown, ComboBox, DefaultButton,
    DetailsListLayoutMode, SelectionMode
  } from '@fluentui/react';
import { SPContext } from '../SPContext';
import ProjectDetails from './ProjectDetails';
 export interface IProjectInfo {
  Id: number;
  ProjectNumber: string;
  ProjectName: string;
  Status: string;
  Client: string;
  PM: string;
  Manager: string;
  Sector: string;
}

const columns: IColumn[] = [
  { key: 'col1', name: 'Project Number', fieldName: 'ProjectNumber', minWidth: 150, isResizable: true },
  { key: 'col2', name: 'Project Name', fieldName: 'ProjectName', minWidth: 150, isResizable: true },
  { key: 'col3', name: 'Sector', fieldName: 'Sector', minWidth: 100, isResizable: true },
  { key: 'col4', name: 'Status', fieldName: 'Status', minWidth: 100, isResizable: true },
  { key: 'col5', name: 'Client', fieldName: 'Client', minWidth: 100, isResizable: true },
  { key: 'col6', name: 'Project Manager', fieldName: 'PM', minWidth: 100, isResizable: true },
  { key: 'col7', name: 'Information Manager', fieldName: 'Manager', minWidth: 100, isResizable: true },
];



/**
 * A component that displays a list of projects in a details list.
 * It provides text fields to search by project number and name, and dropdowns to filter by sector and status.
 * It also provides a combo box to filter by client, and a 'Clear Filters' button to clear all filters.
 * When a project is selected, it displays the project details in a separate component.
 * The component uses the SPContext to load the list of projects from the "9719_ProjectInformationDatabase" list.
 * If the component is unable to load the project list, it displays an error message.
 */
const ProjectList: React.FC = () => {
  const sp = useContext(SPContext);
  const [items, setItems] = useState<IProjectInfo[]>([]);
  const [loading, setLoading] = useState(true);
  const [selectedProject, setSelectedProject] = useState<IProjectInfo | null>(null);
  const [searchNumber, setSearchNumber] = useState('');
  const [searchName, setSearchName] = useState('');
  const [filterSector, setFilterSector] = useState('');
  const [filterStatus, setFilterStatus] = useState('');
  const [filterClient, setFilterClient] = useState('');



  useEffect(() => {
    const fetchItems = async (): Promise<void> => {
      try {
        const data: { Id: number; Title: string; ProjectName: string; Status: string; Sector: string; Client: string; PM?: { Title: string }; Manager?: { Title: string }; Checker?: { Title: string }; Approver?: { Title: string } }[] = await sp.web.lists
          .getByTitle("9719_ProjectInformationDatabase")
          .items
          .expand("PM", "Manager", "Checker", "Approver")
          .select("Id", "Title", "ProjectName", "Status", "Sector", "Client", 
                  "PM/Title", "Manager/Title", "Checker/Title", "Approver/Title")
          ();

        const formattedData: IProjectInfo[] = data.map(item => ({
          Id: item.Id,
          ProjectNumber: item.Title,
          ProjectName: item.ProjectName,
          Status: item.Status,
          Client: item.Client,
          Sector: item.Sector,
          PM: item.PM?.Title || "",
          Manager: item.Manager?.Title || "",
        }));
        setItems(formattedData);
      } catch (error) {
        console.error("Failed to load project list items", error);
      } finally {
        setLoading(false);
      }
    };

    fetchItems().catch((error) => {
      console.error("Error fetching project items:", error);      
    });
  }, []);

  return (
    <Stack tokens={{ childrenGap: 20 }} styles={{ root: { width: '100%', height: '100vh' } }}>
      <Stack
        horizontal
        wrap
        tokens={{ childrenGap: 15, padding: 10 }}
        styles={{
          root: {
            marginBottom: 20,
            alignItems: 'center',
            justifyContent: 'space-between',
            columnGap: '15px',
            rowGap: '10px',
            display: 'flex',
            flexWrap: 'wrap'
          }
        }}
      >
        <TextField
          label="Project Number"
          value={searchNumber}
          onChange={(_, v) => setSearchNumber(v || '')}
          styles={{ root: { minWidth: 180, margin: '7.5px 7.5px'} }}
        />
        <TextField
          label="Project Name"
          value={searchName}
          onChange={(_, v) => setSearchName(v || '')}
          styles={{ root: { minWidth: 180, margin: '7.5px 7.5px'} }}
        />
        <Dropdown
          label="Sector"
          options={[
            { key: '', text: 'All' },
            { key: 'Public', text: 'Public' },
            { key: 'Defence', text: 'Defence' },
            { key: 'Power', text: 'Power' }
          ]}
          selectedKey={filterSector}
          onChange={(_, option) => setFilterSector(option?.key as string)}
          styles={{ root: { minWidth: 100 } }}
        />
        <Dropdown
          label="Status"
          options={[
            { key: '', text: 'All' },
            { key: 'Enquiry', text: 'Enquiry' },
            { key: 'Active', text: 'Active' },
            { key: 'Complete', text: 'Complete' },
            { key: 'Lost', text: 'Lost' },
            { key: 'Cancelled', text: 'Cancelled' },
            { key: 'Inactive', text: 'Inactive' }
          ]}
          selectedKey={filterStatus}
          onChange={(_, option) => setFilterStatus(option?.key as string)}
          styles={{ root: { minWidth: 100 } }}
        />
        <ComboBox
          label="Client"
          allowFreeform
          autoComplete="on"
          options={[...Array.from(new Set(items.map(i => i.Client))).map(c => ({ key: c, text: c }))]}
          selectedKey={filterClient}
          onChange={(_, option) => setFilterClient(option?.key as string)}
          onPendingValueChanged={(val) => setFilterClient(val?.key as string)}
          styles={{ root: { minWidth: 180 } }}
        />

        {/* Push Clear Button to the end */}
        <Stack.Item grow align="end" styles={{ root: { marginLeft: 'auto', marginTop: 8 } }}>
          <DefaultButton
            text="Clear Filters"
            onClick={() => {
              setSearchNumber('');
              setSearchName('');
              setFilterSector('');
              setFilterStatus('');
              setFilterClient('');
            }}
            primary
            styles={{
              root: {
                backgroundColor: '#b2531a', // Your theme color (e.g., same as your + Add Project button)
                borderColor: '#b2531a',
                color: 'white'
              },
              rootHovered: {
                backgroundColor: '#934311', // darker on hover
                borderColor: '#934311'
              }
            }}
          />
        </Stack.Item>
      </Stack>

      <Stack styles={{ root: { width: '100%' } }}>
        <Text variant="xLarge">📁 All Projects</Text>
        {loading ? (
          <Text>Loading...</Text>
        ) : (
          
          <DetailsList
          items={items.filter(item =>
                  item.ProjectNumber?.toLowerCase().includes(searchNumber.toLowerCase()) &&
                  item.ProjectName?.toLowerCase().includes(searchName.toLowerCase()) &&
                  (filterSector === '' || item.Sector === filterSector) &&
                  (filterStatus === '' || item.Status === filterStatus) &&
                  (filterClient === '' || item.Client?.toLowerCase().includes(filterClient.toLowerCase()))
                )}
          
            columns={columns}
            setKey="items"
            layoutMode={DetailsListLayoutMode.fixedColumns}
            selectionMode={SelectionMode.none}
            onActiveItemChanged={(item) => setSelectedProject(item)}
          />
        )}
      </Stack>

      <Stack styles={{ root: { width: '40%', background: '#f3f2f1', padding: 10, borderRadius: 4 } }}>
        <ProjectDetails
          project={selectedProject ?? undefined}
          onEdit={() => console.log("Edit not implemented yet")}
          onDelete={() => console.log("Delete not implemented yet")}
        />
      </Stack>
    </Stack>
  );
};



export default ProjectList;
