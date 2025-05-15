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
  // DataGridSelectionCell,
  TableColumnDefinition,
  createTableColumn,
} from '@fluentui/react-components';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { spfi } from '@pnp/sp';
import { SPFx } from '@pnp/sp/presets/all';
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { Project } from '../../services/ProjectSelectionServices';

interface FeesTabProps {
  project: Project;
  context: WebPartContext;
}

interface FeeItem {
  Id: number;
  FeeName: string;
  FeeAmount?: number;
  BudgetFee?: number;
  SpendToDate?: number;
  ForecastHours?: number;
  ActualHours?: number;
  RemianingBudgetOverspend?: number;
}

const useStyles = makeStyles({
  root: {
    display: 'flex',
    flexDirection: 'row',
    gap: tokens.spacingHorizontalXL,
    paddingTop: tokens.spacingVerticalL,
  },
  leftPanel: {
    width: '50%',
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

const listName = '9719_ProjectFees';

const FeesTab: React.FC<FeesTabProps> = ({ context, project }) => {
  const sp = spfi().using(SPFx(context));
  const styles = useStyles();

  const [fees, setFees] = useState<FeeItem[]>([]);
  const [selectedFee, setSelectedFee] = useState<FeeItem | null>(null);

  const ensureFeeEntry = async (): Promise<void> => {
    const items = await sp.web.lists.getByTitle(listName).items
      .filter(`ProjectNumber eq '${project.ProjectNumber}'`).top(1)();
    if (!items.length) {
      const result = await sp.web.lists.getByTitle(listName).items.add({
        FeeName: 'Initial Budget',
        ProjectNumber: project.ProjectNumber
      });
      setSelectedFee(result.data);
      setFees([result.data]);
    } else {
      setFees(items);
      setSelectedFee(items[0]);
    }
  };

  useEffect(() => {
    ensureFeeEntry().catch(console.error);
  }, [project]);

  const columns: TableColumnDefinition<FeeItem>[] = [
    createTableColumn<FeeItem>({
      columnId: 'FeeName',
      renderHeaderCell: () => 'Fee Name',
      renderCell: (item) => item.FeeName,
    }),
    createTableColumn<FeeItem>({
      columnId: 'BudgetFee',
      renderHeaderCell: () => 'Budget (£)',
      renderCell: (item) => `£${item.BudgetFee?.toFixed(2) ?? '0.00'}`,
    }),
    createTableColumn<FeeItem>({
      columnId: 'SpendToDate',
      renderHeaderCell: () => 'Spend (£)',
      renderCell: (item) => `£${item.SpendToDate?.toFixed(2) ?? '0.00'}`,
    }),
    createTableColumn<FeeItem>({
      columnId: 'ForecastHours',
      renderHeaderCell: () => 'Forecast Hrs',
      renderCell: (item) => `${item.ForecastHours ?? 0}`,
    }),
  ];

  return (
    <div className={styles.root}>
      <div className={styles.leftPanel}>
        <Text size={500} weight="semibold">Fee Summary</Text>
        <DataGrid
          items={fees}
          columns={columns}
          selectionMode="single"
          getRowId={(item) => item.Id.toString()}
          onSelectionChange={(event, data) => {
            const selectedId = Array.from(data.selectedItems.values())[0];
            const fee = fees.find(f => f.Id.toString() === selectedId);
            setSelectedFee(fee ?? null);
          }}
        >
          <DataGridHeader>
            {(ctx: { columns: TableColumnDefinition<FeeItem>[] }) => (
              <DataGridRow>
                {(column) => (
                  ctx.columns.map((column) => (
                    <DataGridCell key={column.columnId}>{column.renderHeaderCell()}</DataGridCell>
                  ))
                )}
              </DataGridRow>
            )}
          </DataGridHeader>
          <DataGridBody<FeeItem>>
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
        <Text size={500} weight="semibold">Fee Details</Text>
        {selectedFee ? (
          <>
            <Text><strong>Fee Name:</strong> {selectedFee.FeeName}</Text>
            <Text><strong>Fee Amount:</strong> £{selectedFee.FeeAmount?.toFixed(2) ?? '0.00'}</Text>
            <Text><strong>Budget Fee:</strong> £{selectedFee.BudgetFee?.toFixed(2) ?? '0.00'}</Text>
            <Text><strong>Spend to Date:</strong> £{selectedFee.SpendToDate?.toFixed(2) ?? '0.00'}</Text>
            <Text><strong>Forecast Hours:</strong> {selectedFee.ForecastHours ?? 0}</Text>
            <Text><strong>Actual Hours:</strong> {selectedFee.ActualHours ?? 0}</Text>
            <Text><strong>Remaining/Overspend:</strong> £{selectedFee.RemianingBudgetOverspend?.toFixed(2) ?? '0.00'}</Text>
          </>
        ) : (
          <Text>No fee selected.</Text>
        )}
      </div>
    </div>
  );
};

export default FeesTab;
