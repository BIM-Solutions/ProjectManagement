import * as React from 'react';
import { useEffect, useState } from 'react';
import {
  Text,
  makeStyles,
  tokens,
  Button,
} from '@fluentui/react-components';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { spfi } from '@pnp/sp';
import { SPFx } from '@pnp/sp/presets/all';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import { Project } from '../../services/ProjectSelectionServices';

interface StageOverviewProps {
  project: Project | undefined;
  context: WebPartContext;
  onStageSelect?: (stageId: number) => void;
}

interface StageItem {
  Id: number;
  Title: string;
  StartDate: string;
  EndDate: string;
  Status: string;
  StageColor: string;
  ProjectNumber: string;
}

const listName = '9719_ProjectStages';
const defaultColor = '#0078D4';

const useStyles = makeStyles({
  root: {
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalM,
    height: '100%',
  },
  calendar: {
    flexGrow: 1,
    backgroundColor: tokens.colorNeutralBackground1,
    padding: tokens.spacingVerticalM,
    borderRadius: tokens.borderRadiusMedium,
    boxShadow: tokens.shadow4,
  },
  calendarGrid: {
    display: 'grid',
    gridTemplateColumns: 'repeat(4, 1fr)',
    gap: '1px',
    backgroundColor: tokens.colorNeutralStroke1,
    padding: '1px',
  },
  calendarHeader: {
    display: 'flex',
    justifyContent: 'space-between',
    alignItems: 'center',
    marginBottom: tokens.spacingVerticalM,
  },
  calendarCell: {
    backgroundColor: tokens.colorNeutralBackground1,
    padding: tokens.spacingVerticalS,
    minHeight: '120px',
    position: 'relative',
    cursor: 'pointer',
    '&:hover': {
      backgroundColor: tokens.colorNeutralBackground1Hover,
    },
  },
  monthName: {
    position: 'absolute',
    top: '4px',
    left: '4px',
    fontWeight: 'bold',
  },
  stageContainer: {
    position: 'absolute',
    top: '30px',
    left: '4px',
    right: '4px',
    bottom: '4px',
    display: 'flex',
    flexDirection: 'column',
    gap: '4px',
  },
  stageBar: {
    height: '24px',
    borderRadius: '4px',
    display: 'flex',
    alignItems: 'center',
    padding: '0 8px',
    fontSize: '12px',
    color: 'white',
    whiteSpace: 'nowrap',
    overflow: 'hidden',
    textOverflow: 'ellipsis',
  },
});

const StageOverview: React.FC<StageOverviewProps> = ({ project, context, onStageSelect }) => {
  const sp = spfi().using(SPFx(context));
  const styles = useStyles();

  const [stages, setStages] = useState<StageItem[]>([]);
  const [currentYear, setCurrentYear] = useState(new Date().getFullYear());

  const fetchStages = async (): Promise<void> => {
    try {
      const results = await sp.web.lists.getByTitle(listName).items
        .filter(`ProjectNumber eq '${project?.ProjectNumber}'`).top(5000)();
      console.log('Project number:', project?.ProjectNumber);
      console.log('List Name:', listName);
      console.log('Fetched stages:', results);
      setStages(results);
    } catch (error) {
      console.error('Error fetching stages:', error);
    }
  };

  useEffect(() => {
    fetchStages().catch(console.error);
  }, [project]);

  const getStagesForMonth = (year: number, month: number): StageItem[] => {
    const monthStart = new Date(year, month, 1);
    const monthEnd = new Date(year, month + 1, 0);

    return stages.filter(stage => {
      const stageStart = new Date(stage.StartDate);
      const stageEnd = new Date(stage.EndDate);
      return stageStart <= monthEnd && stageEnd >= monthStart;
    });
  };

  const renderCalendar = (): React.ReactNode => {
    const months = [];
    const monthNames = [
      'January', 'February', 'March', 'April',
      'May', 'June', 'July', 'August',
      'September', 'October', 'November', 'December'
    ];

    for (let month = 0; month < 12; month++) {
      const stagesForMonth = getStagesForMonth(currentYear, month);
      
      months.push(
        <div key={month} className={styles.calendarCell}>
          <Text className={styles.monthName}>{monthNames[month]}</Text>
          <div className={styles.stageContainer}>
            {stagesForMonth.map((stage) => (
              <div
                key={stage.Id}
                className={styles.stageBar}
                style={{ backgroundColor: stage.StageColor || defaultColor }}
                onClick={() => onStageSelect?.(stage.Id)}
                title={`${stage.Title}\nStart: ${new Date(stage.StartDate).toLocaleDateString()}\nEnd: ${new Date(stage.EndDate).toLocaleDateString()}`}
              >
                {stage.Title}
              </div>
            ))}
          </div>
        </div>
      );
    }

    return (
      <div className={styles.calendar}>
        <div className={styles.calendarHeader}>
          <Button
            onClick={() => setCurrentYear(currentYear - 1)}
          >
            Previous
          </Button>
          <Text size={400} weight="semibold">{currentYear}</Text>
          <Button
            onClick={() => setCurrentYear(currentYear + 1)}
          >
            Next
          </Button>
        </div>
        <div className={styles.calendarGrid}>
          {months}
        </div>
      </div>
    );
  };

  return (
    <div className={styles.root}>
      {renderCalendar()}
    </div>
  );
};

export default StageOverview; 