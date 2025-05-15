import * as React from 'react';
import { useEffect, useState } from 'react';
import {
  Text,
  makeStyles,
  tokens,

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
  table: {
    width: '100%',
  },
  clickableRow: {
    cursor: 'pointer',
    '&:hover': {
      backgroundColor: tokens.colorNeutralBackground2,
    },
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

  return (
    <div className={styles.root}>
      <div className={styles.leftPanel}>
        <Text size={500} weight="semibold">Fee Summary</Text>
        <table className={styles.table}>
          <thead>
            <tr>
              <th>Fee Name</th>
              <th>Budget (£)</th>
              <th>Spend (£)</th>
              <th>Forecast Hrs</th>
            </tr>
          </thead>
          <tbody>
            {fees.map(fee => (
              <tr 
                key={fee.Id} 
                className={styles.clickableRow}
                onClick={() => setSelectedFee(fee)}
              >
                <td>{fee.FeeName}</td>
                <td>£{fee.BudgetFee?.toFixed(2) ?? '0.00'}</td>
                <td>£{fee.SpendToDate?.toFixed(2) ?? '0.00'}</td>
                <td>{fee.ForecastHours ?? 0}</td>
              </tr>
            ))}
          </tbody>
        </table>
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
