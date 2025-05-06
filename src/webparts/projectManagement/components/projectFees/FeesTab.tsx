// FeesTab.tsx
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

const listName = '9719_ProjectFees';

const FeesTab: React.FC<FeesTabProps> = ({ context, project }) => {
  const sp = spfi().using(SPFx(context));
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

  return (
    <Stack horizontal tokens={{ childrenGap: 20 }} styles={{ root: { paddingTop: 20 } }}>
      {/* Left Panel - Fee Summary */}
      <Stack tokens={{ childrenGap: 10 }} styles={{ root: { width: '50%' } }}>
        <Text variant="large">Fee Summary</Text>
        <DetailsList
          items={fees}
          columns={[
            { key: 'feeName', name: 'Fee Name', fieldName: 'FeeName', minWidth: 100, isResizable: true },
            { key: 'budget', name: 'Budget (£)', fieldName: 'BudgetFee', minWidth: 100, isResizable: true },
            { key: 'spend', name: 'Spend (£)', fieldName: 'SpendToDate', minWidth: 100, isResizable: true },
            { key: 'hours', name: 'Forecast Hrs', fieldName: 'ForecastHours', minWidth: 80, isResizable: true },
          ]}
          layoutMode={DetailsListLayoutMode.fixedColumns}
          selectionMode={SelectionMode.single}
          onItemInvoked={(item) => setSelectedFee(item)}
        />
      </Stack>

      {/* Right Panel - Fee Details */}
      <Stack tokens={{ childrenGap: 10 }} styles={{ root: { flexGrow: 1 } }}>
        <Text variant="large">Fee Details</Text>
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
      </Stack>
    </Stack>
  );
};

export default FeesTab;
