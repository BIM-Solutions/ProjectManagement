import * as React from 'react';
import { useEffect, useState, useContext } from 'react';
import {
  DetailsList, IColumn, Stack, Text,
  DetailsListLayoutMode, SelectionMode
} from '@fluentui/react';
import { SPContext } from '../SPContext';
import ProjectDetails from './ProjectDetails';

export interface IProjectInfo {
  Id: number;
  Title: string;
  ProjectNumber: string;
  Sector: string;
  KeyPersonnel: string;
}

const columns: IColumn[] = [
  { key: 'col1', name: 'Project Name', fieldName: 'Title', minWidth: 150, isResizable: true },
  { key: 'col2', name: 'Project Number', fieldName: 'ProjectNumber', minWidth: 100, isResizable: true },
  { key: 'col3', name: 'Sector', fieldName: 'Sector', minWidth: 100, isResizable: true },
  { key: 'col4', name: 'Key Personnel', fieldName: 'KeyPersonnel', minWidth: 150, isResizable: true },
];

const ProjectList: React.FC = () => {
  const sp = useContext(SPContext);
  const [items, setItems] = useState<IProjectInfo[]>([]);
  const [loading, setLoading] = useState(true);
  const [selectedProject, setSelectedProject] = useState<IProjectInfo | null>(null);

  useEffect(() => {
    const fetchItems = async (): Promise<void> => {
      try {
        const data: IProjectInfo[] = await sp.web.lists.getByTitle("ProjectInformation").items();
        setItems(data);
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
    <Stack horizontal tokens={{ childrenGap: 20 }} styles={{ root: { width: '100%' } }}>
      <Stack styles={{ root: { width: '60%' } }}>
        <Text variant="xLarge">📁 All Projects</Text>
        {loading ? (
          <Text>Loading...</Text>
        ) : (
          <DetailsList
            items={items}
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
