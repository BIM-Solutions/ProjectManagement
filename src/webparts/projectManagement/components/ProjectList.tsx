import * as React from 'react';
import { useEffect, useState, useContext } from 'react';
import {
    DetailsList, IColumn, Stack, Text, TextField, Dropdown, ComboBox, DefaultButton,
    DetailsListLayoutMode, SelectionMode, Spinner, SpinnerSize
  } from '@fluentui/react';
import { SPContext } from '../../common/SPContext';
import { Project, ProjectSelectionService } from '../../common/services/ProjectSelectionServices';
import { useLoading } from '../services/LoadingContext';

const DEBUG = true; // Set to false in production



const columns: IColumn[] = [
  { key: 'col1', name: 'Project Number', fieldName: 'ProjectNumber', minWidth: 150, isResizable: true },
  { key: 'col2', name: 'Project Name', fieldName: 'ProjectName', minWidth: 150, isResizable: true },
  { key: 'col3', name: 'Sector', fieldName: 'Sector', minWidth: 100, isResizable: true },
  { key: 'col4', name: 'Status', fieldName: 'Status', minWidth: 100, isResizable: true },
  { key: 'col5', name: 'Client', fieldName: 'Client', minWidth: 100, isResizable: true },
  {
    key: 'col6',
    name: 'Project Manager',
    minWidth: 150,
    onRender: item => <span>{item.PM?.Title || '—'}</span>
  },
  {
    key: 'col7',
    name: 'Information Manager',
    minWidth: 150,
    onRender: item => <span>{item.Manager?.Title || '—'}</span>
  }
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
  const [items, setItems] = useState<Project[]>([]);
  // Removed unused loading state
  const [searchNumber, setSearchNumber] = useState('');
  const [searchName, setSearchName] = useState('');
  const [filterSector, setFilterSector] = useState('');
  const [filterStatus, setFilterStatus] = useState('');
  const [filterClient, setFilterClient] = useState('');
  const { isLoading, setIsLoading } = useLoading();



  useEffect(() => {


    const fetchItems = async (): Promise<void> => {
      try {
        const data: { Id: number; Title: string; ProjectName: string; Status: string; Sector: string; Client: string;
                ProjectDescription?: string; DeltekSubCodes?: string; SubCodes?: string; ClientContact?: string; ProjectImage?: string;
               PM?: { Id: number, Title: string, EMail: string, JobTitle?: string, Department?: string }; 
               Manager?: { Id: number, Title: string, EMail: string, JobTitle?: string, Department?: string  }; 
               Checker?: { Id: number, Title: string, EMail: string, JobTitle?: string, Department?: string }; 
               Approver?: { Id: number, Title: string, EMail: string, JobTitle?: string, Department?: string} }[] = await sp.web.lists
          .getByTitle("9719_ProjectInformationDatabase")
          .items
          .expand("PM", "Manager", "Checker", "Approver")
          .select(
            "Id", "Title", "ProjectName", "Status", "Sector", "Client","ProjectImage", "ProjectDescription", "DeltekSubCodes", "ClientContact",
            "PM/Id", "PM/Title", "PM/EMail", "PM/JobTitle", "PM/Department",
            "Manager/Id", "Manager/Title", "Manager/EMail", "Manager/JobTitle", "Manager/Department",
            "Checker/Id", "Checker/Title", "Checker/EMail", "Checker/JobTitle", "Checker/Department",
            "Approver/Id", "Approver/Title", "Approver/EMail", "Approver/JobTitle", "Approver/Department"
          )
          ();
          // console.log("Raw item sample:", data[0]); // to check the full structure

            // data.forEach((item, index) => {
            //   console.log(`PM for item ${index}:`, item.PM);
            //   console.log(`Manager for item ${index}:`, item.Manager);
            // });

            const formattedData = data.map<Project>(item => ({
            id: item.Id,
            ProjectNumber: item.Title,
            ProjectName: item.ProjectName,
            Status: item.Status,
            Client: item.Client,
            Sector: item.Sector,
            ProjectDescription: item.ProjectDescription || 'No Project Description',
            DeltekSubCodes: item.DeltekSubCodes || 'Deltek Sub Codes',
            // SubCodes: item.SubCodes || 'Sub Codes',
            ClientContact: item.ClientContact || 'Client Contact',
            ProjectImage: item.ProjectImage || 'Project Image',
            
            PM: item.PM ? {
              Id: item.PM.Id,
              Title: item.PM.Title,
              Email: item.PM.EMail,
              JobTitle: item.PM.JobTitle,
              Department: item.PM.Department
            } : undefined,
            Manager: item.Manager ? {
              Id: item.Manager.Id,
              Title: item.Manager.Title,
              Email: item.Manager.EMail,
              JobTitle: item.Manager.JobTitle,
              Department: item.Manager.Department
            } : undefined,
            Checker: item.Checker ? {
              Id: item.Checker.Id,
              Title: item.Checker.Title,
              Email: item.Checker.EMail,
              JobTitle: item.Checker.JobTitle,
              Department: item.Checker.Department
            } : undefined,
            Approver: item.Approver ? {
              Id: item.Approver.Id,
              Title: item.Approver.Title,
              Email: item.Approver.EMail,
              JobTitle: item.Approver.JobTitle,
              Department: item.Approver.Department
            }
            : undefined
          }))
        
        setItems(formattedData);
      } catch (error) {
        console.error("Failed to load project list items", error);
      } finally {
        setIsLoading(false);
      }
    };

    fetchItems().catch((error) => {
      console.error("Error fetching project items:", error);      
    });
  }, []);

  return (
    <Stack tokens={{ childrenGap: 20 }} styles={{ root: { width: '100%', height: '40%' } }}>
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
              
              rootHovered: {
                backgroundColor: '#934311', // darker on hover
                borderColor: '#934311'
              }
            }}
          />
        </Stack.Item>
      </Stack>

      <Stack styles={{ root: { width: '100%', height: '500px', overflowY: 'auto' } }}>
        <Text variant="xLarge">📁 All Projects</Text>
        {isLoading ? (
           <Spinner label="Updating SharePoint Lists..." size={SpinnerSize.medium} />
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
            checkButtonAriaLabel="select row"
            selectionMode={SelectionMode.single}
            onActiveItemChanged={(item) => {
              if (DEBUG) {
                console.log("Active item changed:", item); // to check the active item
              }

              const selected: Project = {
                id: item.id,
                ProjectNumber: item.ProjectNumber,
                ProjectName: item.ProjectName,
                ProjectDescription: item.ProjectDescription || '',
                DeltekSubCodes: item.DeltekSubCodes || '',
                SubCodes: item.SubCodes || '',
                ClientContact: item.ClientContact || '',
                Status: item.Status,
                Client: item.Client,
                Sector: item.Sector,
                PM: item.PM,
                Manager: item.Manager,
                Checker: item.Checker,
                Approver: item.Approver,
                ProjectImage: item.ProjectImage || ''
              };
              if (DEBUG) {
                console.log("Selected project:", selected); // to check the selected project
              }
              ProjectSelectionService.setSelectedProject(selected);
            }}
          />
        )}
      </Stack>
    </Stack>
  );
};



export default ProjectList;
