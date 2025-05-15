import * as React from 'react';
import { useEffect, useState } from 'react';
import {
  Text,
  makeStyles,
  tokens,
  DataGrid,
  DataGridHeader,
  DataGridRow,
  DataGridBody,
  DataGridCell,
  TableColumnDefinition,
  createTableColumn,
} from '@fluentui/react-components';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { spfi } from '@pnp/sp';
import { SPFx } from '@pnp/sp/presets/all';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
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

const useStyles = makeStyles({
  root: {
    display: 'flex',
    flexDirection: 'row',
    gap: tokens.spacingHorizontalXL,
    paddingTop: tokens.spacingVerticalL,
  },
  leftPanel: {
    width: '40%',
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalM,
  },
  rightPanel: {
    flexGrow: 1,
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalM,
  },
});

const StagesTab: React.FC<StagesTabProps> = ({ project, context }) => {
  const sp = spfi().using(SPFx(context));
  const styles = useStyles();

  const [stages, setStages] = useState<StageItem[]>([]);
  const [selectedStage, setSelectedStage] = useState<StageItem | null>(null);

  const ensureProjectStage = async (): Promise<void> => {
    const items = await sp.web.lists.getByTitle(listName).items
      .filter(`ProjectNumber eq '${project.ProjectNumber}'`).top(5000)();
    if (!items.length) {
      await sp.web.lists.getByTitle(listName).items.add({
        Title: 'Placeholder Stage',
        ProjectNumber: project.ProjectNumber,
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

  const columns: TableColumnDefinition<StageItem>[] = [
    createTableColumn<StageItem>({
      columnId: 'Title',
      renderHeaderCell: () => 'Stage Name',
      renderCell: (item) => item.Title,
    }),
    createTableColumn<StageItem>({
      columnId: 'StartDate',
      renderHeaderCell: () => 'Start Date',
      renderCell: (item) => item.StartDate ?? '',
    }),
    createTableColumn<StageItem>({
      columnId: 'EndDate',
      renderHeaderCell: () => 'End Date',
      renderCell: (item) => item.EndDate ?? '',
    }),
    createTableColumn<StageItem>({
      columnId: 'Status',
      renderHeaderCell: () => 'Status',
      renderCell: (item) => item.Status ?? '',
    }),
  ];

  return (
    <div className={styles.root}>
      <div className={styles.leftPanel}>
        <Text size={500} weight="semibold">Stages</Text>
        <DataGrid
          items={stages}
          columns={columns}
          selectionMode="single"
          getRowId={(item) => item.Id.toString()}
          onSelectionChange={(event, data) => {
            const selected = stages.find(s => s.Id.toString() === Array.from(data.selectedItems)[0]);
            setSelectedStage(selected ?? null);
          }}
        >
          <DataGridHeader>
            {(ctx: {columns: TableColumnDefinition<StageItem>[]}) => (
              <DataGridRow>
                {(column) => (
                  ctx.columns.map((column) => (

                  <DataGridCell key={column.columnId}>{column.renderHeaderCell()}</DataGridCell>
                )))}
              </DataGridRow>
            )}
          </DataGridHeader>
          <DataGridBody<StageItem>>
            {(row) => (
              <React.Fragment>
                {columns.map((column, index) => (
                  <DataGridCell key={column.columnId} aria-colindex={index + 1}>
                    {column.renderCell(row.item)}
                  </DataGridCell>
                ))}
              </React.Fragment>
            )}
          </DataGridBody>
        </DataGrid>
      </div>

      <div className={styles.rightPanel}>
        <Text size={500} weight="semibold">Stage Details</Text>
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
      </div>
    </div>
  );
};

export default StagesTab;
