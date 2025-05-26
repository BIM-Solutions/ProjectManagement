import * as React from "react";
import { useEffect, useState, useRef } from "react";
import { Text, makeStyles, tokens, Button } from "@fluentui/react-components";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { spfi } from "@pnp/sp";
import { SPFx } from "@pnp/sp/presets/all";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { Project } from "../../services/ProjectSelectionServices";

interface StageOverviewProps {
  project: Project | undefined;
  context: WebPartContext;
  onStageSelect?: (stageId: number) => void;
  stagesChanged?: number;
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

const listName = "9719_ProjectStages";
const defaultColor = "#0078D4";

const useStyles = makeStyles({
  root: {
    display: "flex",
    flexDirection: "column",
    gap: tokens.spacingVerticalM,
    height: "100%",
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
    display: "grid",
    gridTemplateColumns: "repeat(4, 1fr)",
    gap: 0,
    backgroundColor: "transparent",
    padding: 0,
  },
  calendarHeader: {
    display: "flex",
    justifyContent: "space-between",
    alignItems: "center",
    marginBottom: tokens.spacingVerticalM,
  },
  calendarCell: {
    backgroundColor: tokens.colorNeutralBackground1,
    padding: 0,
    aspectRatio: "1 / 1",
    border: `1px solid ${tokens.colorNeutralStroke1}`,
    borderRadius: 0,
    position: "relative",
    cursor: "pointer",
    boxShadow: tokens.shadow2,
    transition: "box-shadow 0.2s",
    "&:hover": {
      boxShadow: tokens.shadow8,
      zIndex: 2,
    },
  },
  monthName: {
    position: "absolute",
    top: "8px",
    left: "12px",
    fontWeight: "bold",
    fontSize: "16px",
    color: tokens.colorNeutralForeground1,
  },
  stageBar: {
    borderRadius: "16px",
    display: "flex",
    flexDirection: "column",
    alignItems: "flex-start",
    padding: "8px 12px",
    fontSize: "14px",
    color: "white",
    boxShadow: tokens.shadow2,
    cursor: "pointer",
    transition: "transform 0.1s, box-shadow 0.2s",
    "&:hover": {
      transform: "scale(1.03)",
      boxShadow: tokens.shadow8,
    },
  },
  inline: {
    flexDirection: "row",
    alignItems: "center",
    gap: "8px",
    "& .stageTitle": {
      marginBottom: 0,
      fontSize: "13px",
    },
    "& .stageDates": {
      fontSize: "11px",
    },
  },
  stageTitle: {
    fontWeight: 700,
    fontSize: "15px",
    marginBottom: "2px",
    lineHeight: 1.2,
    whiteSpace: "nowrap",
    overflow: "hidden",
    textOverflow: "ellipsis",
  },
  stageDates: {
    fontWeight: 400,
    fontSize: "12px",
    opacity: 0.85,
    lineHeight: 1.2,
    whiteSpace: "nowrap",
    overflow: "hidden",
    textOverflow: "ellipsis",
  },
  legend: {
    display: "flex",
    gap: "24px",
    marginTop: "20px",
    alignItems: "center",
    flexWrap: "wrap",
  },
  legendItem: {
    display: "flex",
    alignItems: "center",
    gap: "8px",
  },
  legendColor: {
    width: "18px",
    height: "18px",
    borderRadius: "50%",
    display: "inline-block",
    border: `2px solid ${tokens.colorNeutralStroke1}`,
  },
});

const StageOverview: React.FC<StageOverviewProps> = ({
  project,
  context,
  onStageSelect,
  stagesChanged,
}) => {
  const sp = spfi().using(SPFx(context));
  const styles = useStyles();

  const [stages, setStages] = useState<StageItem[]>([]);
  const [currentYear, setCurrentYear] = useState(new Date().getFullYear());

  const fetchStages = async (): Promise<void> => {
    try {
      const results = await sp.web.lists
        .getByTitle(listName)
        .items.filter(`ProjectNumber eq '${project?.ProjectNumber}'`)
        .top(5000)();
      setStages(results);
    } catch (error) {
      console.error("Error fetching stages:", error);
    }
  };

  useEffect(() => {
    fetchStages().catch(console.error);
  }, [project, stagesChanged]);

  const renderCalendar = (): React.ReactNode => {
    const columns = 4;
    const rows = 3;
    const monthNames = [
      "January",
      "February",
      "March",
      "April",
      "May",
      "June",
      "July",
      "August",
      "September",
      "October",
      "November",
      "December",
    ];

    // cell sizing
    const gridRef = useRef<HTMLDivElement>(null);
    const [cellSize, setCellSize] = useState(140);
    useEffect(() => {
      const updateSize = (): void => {
        if (gridRef.current && gridRef.current.children.length > 0) {
          const firstCell = gridRef.current.children[0] as HTMLElement;
          setCellSize(firstCell.offsetWidth);
          // Optionally log for debugging
          // console.log('cellSize', firstCell.offsetWidth);
        }
      };

      updateSize(); // Initial call

      window.addEventListener("resize", updateSize);

      // Optionally, update after a short delay to catch late renders
      const timeout = setTimeout(updateSize, 100);

      return () => {
        window.removeEventListener("resize", updateSize);
        clearTimeout(timeout);
      };
    }, [gridRef]);

    const rowHeight = cellSize;
    const TOP_OFFSET = 32;
    const BOTTOM_OFFSET = 8;
    const ROW_GAP = 4;
    const availableHeight = rowHeight - TOP_OFFSET - BOTTOM_OFFSET;

    // split into segments
    type Segment = {
      row: number;
      segStart: number;
      segEnd: number;
      stage: StageItem;
    };
    const segments: Segment[] = [];
    const yearStart = new Date(currentYear, 0, 1);
    const yearEnd = new Date(currentYear, 11, 31);

    // helper to split across grid rows
    const getRowSegments = (
      startMonth: number,
      endMonth: number
    ): { row: number; segStart: number; segEnd: number }[] => {
      const segs: { row: number; segStart: number; segEnd: number }[] = [];
      const row = Math.floor(startMonth / columns);
      const endRow = Math.floor(endMonth / columns);
      for (let r = row; r <= endRow; r++) {
        const rowStart = r * columns;
        const rowEnd = rowStart + columns - 1;
        segs.push({
          row: r,
          segStart: r === row ? startMonth : rowStart,
          segEnd: r === endRow ? endMonth : rowEnd,
        });
      }
      return segs;
    };

    // collect segments
    stages.forEach((stage) => {
      const start = new Date(stage.StartDate);
      const end = new Date(stage.EndDate);
      if (end < yearStart || start > yearEnd) return;
      const sm = start.getFullYear() === currentYear ? start.getMonth() : 0;
      const em = end.getFullYear() === currentYear ? end.getMonth() : 11;
      getRowSegments(sm, em).forEach(({ row, segStart, segEnd }) => {
        segments.push({ row, segStart, segEnd, stage });
      });
    });

    // Group segments by row
    const rowSegments: Record<number, Segment[]> = {};
    segments.forEach((seg) => {
      if (!rowSegments[seg.row]) rowSegments[seg.row] = [];
      rowSegments[seg.row].push(seg);
    });

    // render month cells
    const months = monthNames.map((name, i) => (
      <div key={i} className={styles.calendarCell}>
        <Text className={styles.monthName}>{name}</Text>
      </div>
    ));

    // Calculate heights and assign stack index per row
    const bars: React.ReactNode[] = [];
    Object.entries(rowSegments).forEach(([rowStr, segs]) => {
      const row = +rowStr;
      const count = segs.length;
      // Calculate bar height with a max of 36px
      const barHeight = Math.min(
        36,
        (availableHeight - (count - 1) * ROW_GAP) / count - 16
      );
      segs.forEach((seg, idx) => {
        const { segStart, segEnd, stage } = seg;
        const leftPct = ((segStart % columns) / columns) * 100;
        const widthPct = ((segEnd - segStart + 1) / columns) * 100;
        bars.push(
          <div
            key={`${stage.Id}-${row}-${idx}`}
            className={`${styles.stageBar} ${
              barHeight < 36 ? styles.inline : ""
            }`}
            style={{
              backgroundColor: stage.StageColor || defaultColor,
              position: "absolute",
              left: `calc(${leftPct}% + 2px)`,
              width: `calc(${widthPct}% - 28px)`,
              top: `${
                row * rowHeight + TOP_OFFSET + idx * (barHeight + 15 + ROW_GAP)
              }px`,
              height: `${barHeight}px`,
              pointerEvents: "auto",
              zIndex: 2,
            }}
            onClick={(e) => {
              e.stopPropagation();
              onStageSelect?.(stage.Id);
            }}
            title={`${stage.Title}\n${new Date(
              stage.StartDate
            ).toLocaleDateString()} - ${new Date(
              stage.EndDate
            ).toLocaleDateString()}`}
          >
            <span className={styles.stageTitle}>{stage.Title}</span>
            <span className={styles.stageDates}>
              {new Date(stage.StartDate).toLocaleDateString()} -{" "}
              {new Date(stage.EndDate).toLocaleDateString()}
            </span>
          </div>
        );
      });
    });

    return (
      <div
        className={styles.calendar}
        style={{
          position: "relative",
          minHeight: `${rowHeight * rows + 48}px`,
        }}
      >
        <div className={styles.calendarHeader}>
          <Button
            appearance="primary"
            onClick={() => setCurrentYear(currentYear - 1)}
          >
            {currentYear - 1}
          </Button>
          <Text size={400} weight="semibold">
            {currentYear}
          </Text>
          <Button
            appearance="primary"
            onClick={() => setCurrentYear(currentYear + 1)}
          >
            {currentYear + 1}
          </Button>
        </div>
        <div
          className={styles.calendarGrid}
          ref={gridRef}
          style={{ position: "relative", zIndex: 1 }}
        >
          {months}
        </div>
        <div
          style={{
            position: "absolute",
            top: 48,
            left: 0,
            right: 0,
            height: `calc(100% - 48px)`,
            pointerEvents: "none",
          }}
        >
          {bars}
        </div>
      </div>
    );
  };

  return (
    <div className={styles.root}>
      {renderCalendar()}
      <div className={styles.legend}>
        {Array.from(new Set(stages.map((s) => s.StageColor))).map((color) => {
          const st = stages.find((s) => s.StageColor === color);
          return (
            <div key={color} className={styles.legendItem}>
              <span
                className={styles.legendColor}
                style={{ backgroundColor: color }}
              />
              <span>{st?.Title}</span>
            </div>
          );
        })}
      </div>
    </div>
  );
};

export default StageOverview;
