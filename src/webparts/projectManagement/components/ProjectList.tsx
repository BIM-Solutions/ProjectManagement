import * as React from 'react';
import { useEffect, useState, useContext } from 'react';
import {
  DetailsList, IColumn, Stack, Text, TextField, Dropdown, ComboBox, DefaultButton,
  DetailsListLayoutMode, SelectionMode, Spinner, SpinnerSize
} from '@fluentui/react';
import { SPContext } from '../../common/SPContext';
import { Project, ProjectSelectionService } from '../../common/services/ProjectSelectionServices';
import { useLoading } from '../services/LoadingContext';
import { DEBUG } from '../../common/DevVariables';
import { EventService } from '../../common/services/EventService';
import { projectStatusOptions, sectorOptions } from '../services/ListService';



const ProjectList: React.FC = () => {
  const sp = useContext(SPContext);
  const allItemsRef = React.useRef<Project[]>([]);

  const [items, setItems] = useState<Project[]>([]);
  const [allItems, setAllItems] = useState<Project[]>([]); // ✅ master list
  const [searchNumber, setSearchNumber] = useState('');
  const [searchName, setSearchName] = useState('');
  const [filterSector, setFilterSector] = useState('');
  const [filterStatus, setFilterStatus] = useState('');
  const [filterClient, setFilterClient] = useState('');
  const { isLoading, setIsLoading } = useLoading();
  const [columns, setColumns] = useState<IColumn[]>([]);

  useEffect(() => {
    const fetchItems = async (): Promise<void> => {
      try {
        const data = await sp.web.lists
          .getByTitle("9719_ProjectInformationDatabase")
          .items
          .expand("PM", "Manager", "Checker", "Approver")
          .select(
            "Id", "Title", "ProjectName", "Status", "Sector", "Client", "ProjectImage", "ProjectDescription", "DeltekSubCodes", "ClientContact",
            "PM/Id", "PM/Title", "PM/EMail", "PM/JobTitle", "PM/Department",
            "Manager/Id", "Manager/Title", "Manager/EMail", "Manager/JobTitle", "Manager/Department",
            "Checker/Id", "Checker/Title", "Checker/EMail", "Checker/JobTitle", "Checker/Department",
            "Approver/Id", "Approver/Title", "Approver/EMail", "Approver/JobTitle", "Approver/Department"
          )();

        const formattedData = data.map<Project>(item => ({
          id: item.Id,
          ProjectNumber: item.Title,
          ProjectName: item.ProjectName,
          Status: item.Status,
          Client: item.Client,
          Sector: item.Sector,
          ProjectDescription: item.ProjectDescription || 'No Project Description',
          DeltekSubCodes: item.DeltekSubCodes || 'Deltek Sub Codes',
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
          } : undefined
        }));

        // Sort initially by ProjectNumber ascending
        const defaultSortField: keyof Project = 'ProjectNumber';
        const sortedData = [...formattedData].sort((a, b) => {
          const aVal = a[defaultSortField];
          const bVal = b[defaultSortField];
          if (aVal === null || bVal === null) return 0;
          return aVal > bVal ? 1 : -1;
        });

        setItems(sortedData);
        setAllItems(sortedData);
        allItemsRef.current = sortedData; // Store the master list in the ref
        const builtColumns: IColumn[] = [];
        const createColumn = (key: string, name: string, fieldName: keyof Project): IColumn => ({
          key,
          name,
          fieldName,
          minWidth: 150,
          isResizable: true,
          isSorted: fieldName === 'ProjectNumber', // Default sort by ProjectNumber
          isSortedDescending: false,
          onColumnClick: (_, col) => {
            if (!DEBUG) console.log("Column clicked:", col);
            if (!DEBUG) console.log("BuiltColumns:", builtColumns);
            const isSortedDescending = !col.isSortedDescending;
            const newCols = builtColumns.map(c => ({
              ...c,
              isSorted: c.key === col.key,
              isSortedDescending: c.key === col.key ? isSortedDescending : false,
            }));
            if (!DEBUG) console.log("All items:", allItems);
            const sorted = [...allItemsRef.current].sort((a, b) => {
              const aVal = a[fieldName];
              const bVal = b[fieldName];
              if (!DEBUG) console.log("Sorting:", aVal, bVal);
              if (aVal === null || bVal === null || aVal === undefined || bVal === undefined) return 0;
              return isSortedDescending ? (aVal < bVal ? 1 : -1) : (aVal > bVal ? 1 : -1);
            });
            if (!DEBUG) console.log("Sorted items:", sorted);
            if (!DEBUG) console.log("Sorted columns:", newCols);
            setItems(sorted);
            setAllItems(sorted);
            allItemsRef.current = sorted; // Update the master list in the ref
            setColumns(newCols);
          }
        });

        

        builtColumns.push(
          createColumn('col1', 'Project Number', 'ProjectNumber'),
          createColumn('col2', 'Project Name', 'ProjectName'),
          createColumn('col3', 'Sector', 'Sector'),
          createColumn('col4', 'Status', 'Status'),
          createColumn('col5', 'Client', 'Client'),
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
        );
        if (!DEBUG) console.log("Built columns after build:", builtColumns);

        setColumns(builtColumns);
      } catch (error) {
        console.error("Failed to load project list items", error);
      } finally {
        setIsLoading(false);
      }
    };

    fetchItems().catch(console.error);
    const unsubscribe = EventService.subscribeToProjectUpdates(() => {
      if (!DEBUG) console.log("Project updated event received, re-fetching items...");
      fetchItems().catch(console.error);
    });
    return () => unsubscribe();
  }, []);

  return (
    <Stack tokens={{ childrenGap: 20 }} styles={{ root: { width: '100%', height: 'auto' } }}>
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
          styles={{ root: { minWidth: 180, margin: '7.5px' } }}
        />
        <TextField
          label="Project Name"
          value={searchName}
          onChange={(_, v) => setSearchName(v || '')}
          styles={{ root: { minWidth: 180, margin: '7.5px' } }}
        />
        <Dropdown
          label="Sector"
          options={sectorOptions}
          selectedKey={filterSector}
          onChange={(_, option) => setFilterSector(option?.key as string)}
          styles={{ root: { minWidth: 100 } }}
        />
        <Dropdown
          label="Status"
          options={projectStatusOptions}
          selectedKey={filterStatus}
          onChange={(_, option) => setFilterStatus(option?.key as string)}
          styles={{ root: { minWidth: 100 } }}
        />
        <ComboBox
          label="Client"
          allowFreeform
          autoComplete="on"
          options={[...Array.from(new Set(allItems.map(i => i.Client))).map(c => ({ key: c, text: c }))]}
          selectedKey={filterClient}
          onChange={(_, option) => setFilterClient(option?.key as string)}
          onPendingValueChanged={(val) => setFilterClient(val?.key as string)}
          styles={{ root: { minWidth: 180 } }}
        />

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
          />
        </Stack.Item>
      </Stack>

      <Stack styles={{ root: { width: '100%', height: 'auto', overflowY: 'auto', overflowX: 'auto' } }}>
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
            selectionMode={SelectionMode.single}
            onActiveItemChanged={(item) => {
              if (!DEBUG) console.log("Selected project:", item);
              ProjectSelectionService.setSelectedProject(item);
            }}
          />
        )}
      </Stack>
    </Stack>
  );
};

export default ProjectList;
