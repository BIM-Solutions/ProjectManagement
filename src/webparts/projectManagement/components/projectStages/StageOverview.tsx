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
    padding: 0,
    borderRadius: 0,
    boxShadow: tokens.shadow8,
    border: `1px solid ${tokens.colorNeutralStroke1}`,
  },
  calendarGrid: {
    display: 'grid',
    gridTemplateColumns: 'repeat(4, 1fr)',
    gap: 0,
    backgroundColor: 'transparent',
    padding: 0,
  },
  calendarHeader: {
    display: 'flex',
    justifyContent: 'space-between',
    alignItems: 'center',
    marginBottom: tokens.spacingVerticalM,
  },
  calendarCell: {
    backgroundColor: tokens.colorNeutralBackground1,
    padding: 0,
    aspectRatio: '1 / 1',
    border: `1px solid ${tokens.colorNeutralStroke1}`,
    borderRadius: 0,
    position: 'relative',
    cursor: 'pointer',
    boxShadow: tokens.shadow2,
    transition: 'box-shadow 0.2s',
    '&:hover': {
      boxShadow: tokens.shadow8,
      zIndex: 2,
    },
  },
  monthName: {
    position: 'absolute',
    top: '8px',
    left: '12px',
    fontWeight: 'bold',
    fontSize: '16px',
    color: tokens.colorNeutralForeground1,
  },
  stageContainer: {
    position: 'absolute',
    top: '36px',
    left: '8px',
    right: '8px',
    bottom: '8px',
    display: 'flex',
    flexDirection: 'column',
    gap: '8px',
  },
  stageBar: {
    borderRadius: '16px',
    display: 'flex',
    flexDirection: 'column',
    alignItems: 'flex-start',
    padding: '8px 12px',
    fontSize: '14px',
    color: 'white',
    marginBottom: '4px',
    boxShadow: tokens.shadow2,
    cursor: 'pointer',
    transition: 'transform 0.1s, box-shadow 0.2s',
    '&:hover': {
      transform: 'scale(1.03)',
      boxShadow: tokens.shadow8,
    },
  },
  stageTitle: {
    fontWeight: 700,
    fontSize: '15px',
    marginBottom: '2px',
    lineHeight: 1.2,
  },
  stageDates: {
    fontWeight: 400,
    fontSize: '12px',
    opacity: 0.85,
    lineHeight: 1.2,
  },
  legend: {
    display: 'flex',
    gap: '24px',
    marginTop: '20px',
    alignItems: 'center',
    flexWrap: 'wrap',
  },
  legendItem: {
    display: 'flex',
    alignItems: 'center',
    gap: '8px',
  },
  legendColor: {
    width: '18px',
    height: '18px',
    borderRadius: '50%',
    display: 'inline-block',
    border: `2px solid ${tokens.colorNeutralStroke1}`,
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
      // console.log('Project number:', project?.ProjectNumber);
      // console.log('List Name:', listName);
      // console.log('Fetched stages:', results);
      setStages(results);
    } catch (error) {
      console.error('Error fetching stages:', error);
    }
  };

  useEffect(() => {
    fetchStages().catch(console.error);
  }, [project]);

  // const getStagesForMonth = (year: number, month: number): StageItem[] => {
  //   const monthStart = new Date(year, month, 1);
  //   const monthEnd = new Date(year, month + 1, 0);

  //   return stages.filter(stage => {
  //     const stageStart = new Date(stage.StartDate);
  //     const stageEnd = new Date(stage.EndDate);
  //     return stageStart <= monthEnd && stageEnd >= monthStart;
  //   });
  // };

  const renderCalendar = (): React.ReactNode => {
    const months = [];
    const monthNames = [
      'January', 'February', 'March', 'April',
      'May', 'June', 'July', 'August',
      'September', 'October', 'November', 'December'
    ];
    // Render empty month cells for the grid
    for (let month = 0; month < 12; month++) {
      months.push(
        <div key={month} className={styles.calendarCell}>
          <Text className={styles.monthName}>{monthNames[month]}</Text>
        </div>
      );
    }

    // Calculate bar positions for overlay
    // const monthCount = 12;
    const columns = 4;
    const rows = 3;
    const barHeight = 36;
    const barStacking: { [row: number]: number } = {};

    // Helper to split a stage into row segments
    function getRowSegments(startMonth: number, endMonth: number): { row: number, segStart: number, segEnd: number }[] {
      const segments = [];
      const row = Math.floor(startMonth / columns);
      const endRow = Math.floor(endMonth / columns);
      for (let r = row; r <= endRow; r++) {
        const rowStartMonth = r * columns;
        const rowEndMonth = rowStartMonth + columns - 1;
        const segStart = r === row ? startMonth : rowStartMonth;
        const segEnd = r === endRow ? endMonth : rowEndMonth;
        segments.push({ row: r, segStart, segEnd });
      }
      return segments;
    }

    // For stacking bars within a row
    function getStackIndex(row: number): number {
      if (barStacking[row] === undefined) barStacking[row] = 0;
      return barStacking[row]++;
    }

    // Only render bars for stages that overlap the current year
    const yearStart = new Date(currentYear, 0, 1);
    const yearEnd = new Date(currentYear, 11, 31);
    // Get the width of a cell to use for rowHeight (square cells)
    const gridRef = React.useRef<HTMLDivElement>(null);
    const [cellSize, setCellSize] = React.useState(140);
    React.useEffect(() => {
      function updateCellSize(): void {
        if (gridRef.current) {
          const firstCell = gridRef.current.querySelector('div');
          if (firstCell) {
            setCellSize((firstCell as HTMLElement).offsetWidth);
          }
        }
      }
      updateCellSize();
      window.addEventListener('resize', updateCellSize);
      return () => window.removeEventListener('resize', updateCellSize);
    }, [gridRef, months.length]);

    const rowHeight = cellSize;
    const bars = stages.reduce((acc: JSX.Element[], stage: StageItem) => {
      const start = new Date(stage.StartDate);
      const end = new Date(stage.EndDate);
      if (end < yearStart || start > yearEnd) return acc; // skip if not overlapping
      // Clamp to current year for display
      const startMonth = Math.max(0, start.getFullYear() === currentYear ? start.getMonth() : 0);
      const endMonth = Math.min(11, end.getFullYear() === currentYear ? end.getMonth() : 11);
      const segments = getRowSegments(startMonth, endMonth);
      segments.forEach(({ row, segStart, segEnd }) => {
        const left = ((segStart % columns) / columns) * 100;
        const widthPercent = ((segEnd - segStart + 1) / columns) * 100;
        const stackIdx = getStackIndex(row);
        acc.push(
          <div
            key={`${stage.Id}-r${row}`}
            className={styles.stageBar}
            style={{
              backgroundColor: stage.StageColor || defaultColor,
              position: 'absolute',
              left: `calc(${left}% + 2px)`,
              width: `calc(${widthPercent}% - 28px)`,
              top: `${row * rowHeight + 32 + stackIdx * (barHeight + 20)}px`,
              height: `${barHeight}px`,
              zIndex: 2,
              pointerEvents: 'auto',
            }}
            onClick={() => onStageSelect?.(stage.Id)}
            title={`${stage.Title}\nStart: ${new Date(stage.StartDate).toLocaleDateString()}\nEnd: ${new Date(stage.EndDate).toLocaleDateString()}`}
          >
            <span className={styles.stageTitle}>{stage.Title}</span>
            <span className={styles.stageDates}>
              {new Date(stage.StartDate).toLocaleDateString()} - {new Date(stage.EndDate).toLocaleDateString()}
            </span>
          </div>
        );
      });
      return acc;
    }, [] as JSX.Element[]);

    return (
      <div className={styles.calendar} style={{ position: 'relative', minHeight: `${rowHeight * rows + 48}px` }}>
        <div className={styles.calendarHeader}>
          <Button
            appearance="primary"
            onClick={() => setCurrentYear(currentYear - 1)}
          >
            {currentYear - 1}
          </Button>
          <Text size={400} weight="semibold">{currentYear}</Text>
          <Button
            appearance="primary"
            onClick={() => setCurrentYear(currentYear + 1)}
          >
            {currentYear + 1}
          </Button>
        </div>
        <div className={styles.calendarGrid} ref={gridRef} style={{ position: 'relative', zIndex: 1 }}>
          {months}
        </div>
        <div style={{ position: 'absolute', left: 0, right: 0, top: 48, height: 'calc(100% - 48px)', pointerEvents: 'none' }}>
          {bars}
        </div>
      </div>
    );
  };

  return (
    <div className={styles.root}>
      {renderCalendar()}
      {/* Legend for stage colors */}
      <div className={styles.legend}>
        {Array.from(new Set(stages.map(s => s.StageColor))).map((color, idx) => {
          const stage = stages.find(s => s.StageColor === color);
          return (
            <div key={color} className={styles.legendItem}>
              <span className={styles.legendColor} style={{ backgroundColor: color }} />
              <span>{stage?.Title}</span>
            </div>
          );
        })}
      </div>
    </div>
  );
};

export default StageOverview; 