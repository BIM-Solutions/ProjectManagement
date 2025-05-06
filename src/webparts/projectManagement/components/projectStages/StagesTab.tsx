// StagesTab.tsx
import * as React from 'react';
import { useEffect, useState } from 'react';
import { Stack, Text, DetailsList, DetailsListLayoutMode, SelectionMode } from '@fluentui/react';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { spfi } from '@pnp/sp';
import { SPFx } from '@pnp/sp/presets/all';
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";

import { Project } from '../../services/ProjectSelectionServices';

interface StagesTabProps {
  project: Project;
  context: WebPartContext;
}

interface StageItem {
  Id: number;
  Title: string;
  StartDate?: string;
  EndDate?: string;
  Status?: string;
  Notes?: string;
}

const listName = 'ProjectStages';

const StagesTab: React.FC<StagesTabProps> = ({ project, context }) => {
  const sp = spfi().using(SPFx(context));
  const [stages, setStages] = useState<StageItem[]>([]);
  const [selectedStage, setSelectedStage] = useState<StageItem | null>(null);

  const ensureProjectStage = async (): Promise<void> => {
    const items = await sp.web.lists.getByTitle(listName).items
      .filter(`ProjectNumber eq '${project.ProjectNumber}'`).top(5000)();
    if (!items.length) {
      await sp.web.lists.getByTitle(listName).items.add({
        Title: 'Placeholder Stage',
        ProjectNumber: project.ProjectNumber
      });
    }
  };

  const fetchStages = async (): Promise<void> => {
    await ensureProjectStage();
    const results = await sp.web.lists.getByTitle(listName).items
      .filter(`ProjectNumber eq '${project.ProjectNumber}'`).top(5000)();
    setStages(results);
    if (results.length) setSelectedStage(results[0]);
  };

  useEffect(() => {
    fetchStages().catch(console.error);
  }, [project]);

  return (
    <Stack horizontal tokens={{ childrenGap: 20 }} styles={{ root: { paddingTop: 20 } }}>
      {/* Left panel: list of stages */}
      <Stack tokens={{ childrenGap: 10 }} styles={{ root: { width: '40%' } }}>
        <Text variant="large">Stages</Text>
        <DetailsList
          items={stages}
          columns={[
            { key: 'title', name: 'Stage Name', fieldName: 'Title', minWidth: 100, isResizable: true },
            { key: 'start', name: 'Start Date', fieldName: 'StartDate', minWidth: 80, isResizable: true },
            { key: 'end', name: 'End Date', fieldName: 'EndDate', minWidth: 80, isResizable: true },
            { key: 'status', name: 'Status', fieldName: 'Status', minWidth: 80, isResizable: true },
          ]}
          layoutMode={DetailsListLayoutMode.fixedColumns}
          selectionMode={SelectionMode.single}
          onItemInvoked={(item) => setSelectedStage(item)}
        />
      </Stack>

      {/* Right panel: selected stage details */}
      <Stack tokens={{ childrenGap: 10 }} styles={{ root: { flexGrow: 1 } }}>
        <Text variant="large">Stage Details</Text>
        {selectedStage ? (
          <>
            <Text><strong>Stage:</strong> {selectedStage.Title}</Text>
            <Text><strong>Status:</strong> {selectedStage.Status}</Text>
            <Text><strong>Start Date:</strong> {selectedStage.StartDate}</Text>
            <Text><strong>End Date:</strong> {selectedStage.EndDate}</Text>
            <Text><strong>Notes:</strong> {selectedStage.Notes}</Text>
          </>
        ) : (
          <Text>No stage selected.</Text>
        )}
      </Stack>
    </Stack>
  );
};

export default StagesTab;
