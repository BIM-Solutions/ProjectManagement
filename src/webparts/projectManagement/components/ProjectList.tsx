import * as React from 'react';
import { useEffect, useState, useContext } from 'react';
import {
  Input,
  Spinner,
  Table,
  TableBody,
  TableCell,
  TableHeader,
  TableHeaderCell,
  TableRow,
  Select,
  Button,
  makeStyles,
  Persona,
  tokens,
  mergeClasses,


} from '@fluentui/react-components';
import { SPContext } from './common/SPContext';
import { Project, ProjectSelectionService } from '../services/ProjectSelectionServices';
import { useLoading } from '../services/LoadingContext';
// import { DEBUG } from './common/DevVariables';
import { eventService } from '../services/EventService';
import { projectStatusOptions, sectorOptions } from '../services/ListService';


const useStyles = makeStyles({
  root: {
    display: 'flex',
    flexDirection: 'column',
    rowGap: '16px',
    width: '100%',
    overflowX: 'auto',
    overflowY: 'auto'
  },
  filters: {
    display: 'flex',
    flexWrap: 'wrap',
    gap: '12px',
    alignItems: 'flex-end',
  },
  tableContainer: {
    overflow: 'auto',
    border: `1px solid ${tokens.colorNeutralStroke1}`,
    borderRadius: '8px',
    backgroundColor: tokens.colorNeutralBackground1,
    tableLayout: 'auto'
  },
  
  tableHeader: {
    borderBottom: `1px solid ${tokens.colorNeutralStroke1}`,
    backgroundColor: tokens.colorNeutralBackground2,
  },
  
  tableRow: {
    borderBottom: `1px solid ${tokens.colorNeutralStroke1}`,
  },
  
  tableCell: {
    borderRight: `1px solid ${tokens.colorNeutralStroke1}`,
    padding: '8px',
    whiteSpace: 'nowrap',
    overflow: 'hidden',
    textOverflow: 'ellipsis',
    fontSize: '14px',
  },
  clickableRow: {
    cursor: 'pointer',
    '&:hover': {
      backgroundColor: tokens.colorNeutralBackground2Hover,
    },
  },
  selectedRow: {
    backgroundColor: tokens.colorBrandBackground2,
    '&:hover': {
      backgroundColor: tokens.colorBrandBackground2Hover,
    },
  },
});

const ProjectList: React.FC = () => {
  const styles = useStyles();
  const sp = useContext(SPContext);
  const { isLoading, setIsLoading } = useLoading();
  const [items, setItems] = useState<Project[]>([]);
  const [searchNumber, setSearchNumber] = useState('');
  const [searchName, setSearchName] = useState('');
  const [filterSector, setFilterSector] = useState<string>('');
  const [filterStatus, setFilterStatus] = useState<string>('');
  const [sortField, setSortField] = useState<keyof Project>('ProjectNumber');
  const [sortDirection, setSortDirection] = useState<'asc' | 'desc'>('asc');
  const [selectedProjectId, setSelectedProjectId] = useState<number | null>(null);

  // useEffect(() => {
  //   const handleKeyDown = (e: KeyboardEvent): void => {
  //     // Don't deselect if focused on an input or textarea or select
  //     const target = e.target as HTMLElement;
  //     const tag = target?.tagName?.toLowerCase();
  //     const isTypingField = tag === 'input' || tag === 'textarea' || tag === 'select';
  
  //     if (e.key === 'Escape' && !isTypingField) {
  //       setSelectedProjectId(null);
  //       ProjectSelectionService.setSelectedProject(undefined);
  //     }
  //   };
  
  //   document.addEventListener('keydown', handleKeyDown);
  //   return () => {
  //     document.removeEventListener('keydown', handleKeyDown);
  //   };
  // }, []);

  useEffect(() => {
    const fetchItems = async (): Promise<void> => {
      setIsLoading(true);
      try {
        const data = await sp.web.lists.getByTitle("9719_ProjectInformationDatabase").items
          .select("Id", "Title", "ProjectName", "Status", "Sector", "Client", "PM", "Manager", "Checker", "Approver")
          .expand("PM", "Manager", "Checker", "Approver")
          .select(
            "Id", "Title", "ProjectName", "Status", "Sector", "Client", "ProjectImage", "ProjectDescription", "DeltekSubCodes", "ClientContact",
            "PM/Id", "PM/Title", "PM/EMail", "PM/JobTitle", "PM/Department",
            "Manager/Id", "Manager/Title", "Manager/EMail", "Manager/JobTitle", "Manager/Department",
            "Checker/Id", "Checker/Title", "Checker/EMail", "Checker/JobTitle", "Checker/Department",
            "Approver/Id", "Approver/Title", "Approver/EMail", "Approver/JobTitle", "Approver/Department"
          )();

          const formatted = data.map<Project>(item => ({
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
        const sortedData = [...formatted].sort((a, b) => {
          const aVal = a[defaultSortField];
          const bVal = b[defaultSortField];
          if (aVal === null || bVal === null) return 0;
          return aVal > bVal ? 1 : -1;
        });

        setItems(sortedData);
        setSelectedProjectId(ProjectSelectionService.getSelectedProject()?.id || null)
      } catch (err) {
        console.error("Error loading items:", err);
      } finally {
        setIsLoading(false);
      }
    };

    fetchItems().catch(err => {
      console.error("Error fetching items:", err);
      setIsLoading(false);
    });
    const unsubscribe = eventService.subscribeToProjectUpdates(fetchItems);
    return () => unsubscribe();
  }, []);

  const filteredItems = items.filter(item =>
    item.ProjectNumber?.toLowerCase().includes(searchNumber.toLowerCase()) &&
    item.ProjectName?.toLowerCase().includes(searchName.toLowerCase()) &&
    (filterSector === '' || item.Sector === filterSector) &&
    (filterStatus === '' || item.Status === filterStatus)
  );

  const handleSort = (field: keyof Project): void => {
    const isAsc = sortField === field && sortDirection === 'asc';
    setSortField(field);
    setSortDirection(isAsc ? 'desc' : 'asc');
  };

  const sortedItems = [...filteredItems].sort((a, b) => {
    const aVal = a[sortField];
    const bVal = b[sortField];
    if (aVal === null || aVal === undefined || bVal === null || bVal === undefined) return 0;
  
    if (typeof aVal === 'string' && typeof bVal === 'string') {
      return sortDirection === 'asc'
        ? aVal.localeCompare(bVal)
        : bVal.localeCompare(aVal);
    }
  
    return sortDirection === 'asc' ? (aVal > bVal ? 1 : -1) : (aVal < bVal ? 1 : -1);
  });
  

  return (
    <div className={styles.root}>
      <div className={styles.filters}>
        <Input
          placeholder="Search Project Number"
          value={searchNumber}
          onChange={(_, data) => setSearchNumber(data.value)}
        />
        <Input
          placeholder="Search Project Name"
          value={searchName}
          onChange={(_, data) => setSearchName(data.value)}
        />
        <Select
          style={{ minWidth: '150px' }}
          value={filterSector ?? ''}
          onChange={(_, data) => setFilterSector(data.value)}
        >
          <option value="">All Sectors</option>
          {sectorOptions.map(option => (
            <option key={option.key} value={option.value}>
              {option.value}
            </option>
          ))}
        </Select>

        <Select
          style={{ minWidth: '150px' }}
          value={filterSector ?? ''}
          onChange={(_, data) => setFilterStatus(data.value)}
        >
          <option value="">All Statuses</option>
          {projectStatusOptions.map(option => (
            <option key={option.key} value={option.value}>
              {option.value}
            </option>
          ))}
        </Select>

        <Button
          onClick={() => {
            setSearchNumber('');
            setSearchName('');
            setFilterSector('');
            setFilterStatus('');
            ProjectSelectionService.setSelectedProject(undefined);
            setSelectedProjectId(null);
          }}
        >
          Clear Filters
        </Button>
      </div>

      {isLoading ? (
        <Spinner label="Loading projects..." />
      ) : (
        <div className={styles.tableContainer}>
          <Table style={{ tableLayout: 'auto' }} sortable>
            <TableHeader className={styles.tableHeader}>
              <TableRow >
                <TableHeaderCell style={{ minWidth: '150px', maxWidth: '300px' }} onClick={() => handleSort('ProjectNumber')}>
                  Project Number {sortField === 'ProjectNumber' ? (sortDirection === 'asc' ? '↑' : '↓') : ''}
                </TableHeaderCell>
                <TableHeaderCell style={{ minWidth: '150px', maxWidth: '300px' }} onClick={() => handleSort('ProjectName')}>
                  Project Name {sortField === 'ProjectName' ? (sortDirection === 'asc' ? '↑' : '↓') : ''}
                </TableHeaderCell>
                <TableHeaderCell style={{ minWidth: '100px', maxWidth: '150px' }} onClick={() => handleSort('Sector')}>
                  Sector {sortField === 'Sector' ? (sortDirection === 'asc' ? '↑' : '↓') : ''}
                </TableHeaderCell>
                <TableHeaderCell style={{ minWidth: '100px', maxWidth: '150px' }} onClick={() => handleSort('Status')}>
                  Status {sortField === 'Status' ? (sortDirection === 'asc' ? '↑' : '↓') : ''}
                </TableHeaderCell>
                <TableHeaderCell style={{ minWidth: '100px', maxWidth: '150px' }} onClick={() => handleSort('Client')}>
                  Client {sortField === 'Client' ? (sortDirection === 'asc' ? '↑' : '↓') : ''}
                </TableHeaderCell>
                <TableHeaderCell style={{ minWidth: '110px', maxWidth: '250px' }}>Project Manager</TableHeaderCell>
                <TableHeaderCell style={{ minWidth: '110px', maxWidth: '250px' }}>Information Manager</TableHeaderCell>
                <TableHeaderCell style={{ minWidth: '110px', maxWidth: '250px' }}>Checker</TableHeaderCell>
                <TableHeaderCell style={{ minWidth: '110px', maxWidth: '250px' }}>Approver</TableHeaderCell>
              </TableRow>
            </TableHeader>
            <TableBody>
              {sortedItems.map(project => (
                <TableRow
                  key={project.id}
                  className={mergeClasses(
                    styles.tableRow,
                    styles.clickableRow,
                    selectedProjectId === project.id ? styles.selectedRow : ''
                  )}
                  onClick={() => {
                    ProjectSelectionService.setSelectedProject(project);
                    setSelectedProjectId(project.id);
                  }}
                >
                  <TableCell className={styles.tableCell} style={{ minWidth: '150px', maxWidth: '300px' }}>{project.ProjectNumber}</TableCell>
                  <TableCell className={styles.tableCell} style={{ minWidth: '150px', maxWidth: '300px' }}>{project.ProjectName}</TableCell>
                  <TableCell className={styles.tableCell} style={{ minWidth: '100px', maxWidth: '150px' }} >{project.Sector}</TableCell>
                  <TableCell className={styles.tableCell} style={{ minWidth: '100px', maxWidth: '150px' }} >{project.Status}</TableCell>
                  <TableCell className={styles.tableCell} style={{ minWidth: '100px', maxWidth: '150px' }} >{project.Client}</TableCell>
                  <TableCell className={styles.tableCell} style={{ minWidth: '110px', maxWidth: '250px' }}>
                    <Persona
                      size="small"
                      name={project.PM?.Title}
                      avatar={{ image: {
                        src: `/_layouts/15/userphoto.aspx?accountname=${encodeURIComponent(project.PM?.Email || '')}`,
                        alt: project.PM?.Title,
                      } }}
                      // secondaryText={project.PM?.JobTitle}
                      // tertiaryText={project.PM?.Department}
                  />
                  </TableCell>

                  <TableCell className={styles.tableCell} style={{ minWidth: '110px', maxWidth: '250px' }}>
                    <Persona
                      size="small"
                      name={project.Manager?.Title}
                      avatar={{ image: {
                        src: `/_layouts/15/userphoto.aspx?accountname=${encodeURIComponent(project.Manager?.Email || '')}`,
                        alt: project.Manager?.Title,
                      } }}
                      // secondaryText={project.Manager?.JobTitle}
                      // tertiaryText={project.Manager?.Department}
                  />
                  </TableCell>

                  <TableCell className={styles.tableCell} style={{ minWidth: '110px', maxWidth: '250px' }}>
                    <Persona
                      size="small"
                      name={project.Checker?.Title}
                      avatar={{ image: {
                        src: `/_layouts/15/userphoto.aspx?accountname=${encodeURIComponent(project.Checker?.Email || '')}`,
                        alt: project.Checker?.Title,
                      } }}
                      // secondaryText={project.Checker?.JobTitle}
                      // tertiaryText={project.Checker?.Department}
                  />
                  </TableCell>

                  <TableCell className={styles.tableCell} style={{ minWidth: '110px', maxWidth: '250px' }}>
                    <Persona
                      size="small"
                      name={project.Approver?.Title}
                      avatar={{ image: {
                        src: `/_layouts/15/userphoto.aspx?accountname=${encodeURIComponent(project.Approver?.Email || '')}`,
                        alt: project.Approver?.Title,
                      } }}
                      // secondaryText={project.Approver?.JobTitle}
                      // tertiaryText={project.Approver?.Department}
                  />
                  </TableCell>

                </TableRow>
              ))}
            </TableBody>
          </Table>
        </div>
      )}
    </div>
  );
};

export default ProjectList;
