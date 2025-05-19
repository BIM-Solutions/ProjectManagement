import * as React from 'react';
import { useEffect, useState } from 'react';
import {
  Text,
  makeStyles,
  tokens,
  Button,
  Dialog,
  DialogTrigger,
  DialogSurface,
  DialogBody,
  DialogContent,
  DialogActions,
  Input,
  Field,
  shorthands,
} from '@fluentui/react-components';
import { Add24Regular } from '@fluentui/react-icons';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { spfi } from '@pnp/sp';
import { SPFx } from '@pnp/sp/presets/all';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import { Project } from '../../services/ProjectSelectionServices';
import { DatePicker } from '@fluentui/react-datepicker-compat';

interface StagesTabProps {
  project: Project;
  context: WebPartContext;
  selectedStageId?: number;
}

interface StageItem {
  Id: number;
  Title: string;
  StartDate: string;
  EndDate: string;
  Status: string;
  Notes?: string;
  StageColor: string;
  ProjectNumber: string;
}

type NewStageItem = Omit<StageItem, 'Id'>;

const listName = '9719_ProjectStages';
const defaultColor = '#0078D4';

const useStyles = makeStyles({
  root: {
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalM,
    padding: tokens.spacingVerticalM,
    backgroundColor: tokens.colorNeutralBackground1,
    borderRadius: tokens.borderRadiusMedium,
    boxShadow: tokens.shadow4,
  },
  header: {
    display: 'flex',
    justifyContent: 'space-between',
    alignItems: 'center',
  },
  stageList: {
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalS,
    maxHeight: 'calc(100vh - 300px)',
    overflowY: 'auto',
  },
  stageItem: {
    display: 'flex',
    alignItems: 'center',
    padding: tokens.spacingVerticalS,
    borderRadius: tokens.borderRadiusMedium,
    backgroundColor: tokens.colorNeutralBackground2,
    cursor: 'pointer',
    '&:hover': {
      backgroundColor: tokens.colorNeutralBackground2Hover,
    },
  },
  colorDot: {
    width: '12px',
    height: '12px',
    borderRadius: '50%',
    marginRight: tokens.spacingHorizontalS,
  },
  dialogContent: {
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalM,
  },
  colorPickerWrapper: {
    position: 'relative',
    width: '100%',
  },
  colorButton: {
    width: '100%',
    height: '32px',
    ...shorthands.border('1px', 'solid', tokens.colorNeutralStroke1),
    ...shorthands.borderRadius(tokens.borderRadiusMedium),
    cursor: 'pointer',
    padding: 0,
    '&:hover': {
      ...shorthands.borderColor(tokens.colorNeutralStroke1Hover),
    },
  },
  colorInput: {
    position: 'absolute',
    top: 0,
    left: 0,
    width: '100%',
    height: '100%',
    opacity: 0,
    cursor: 'pointer',
  },
  stageDetails: {
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalS,
    padding: tokens.spacingVerticalM,
    backgroundColor: tokens.colorNeutralBackground2,
    borderRadius: tokens.borderRadiusMedium,
  },
});

const StagesTab: React.FC<StagesTabProps> = ({ project, context, selectedStageId }) => {
  const sp = spfi().using(SPFx(context));
  const styles = useStyles();

  const [stages, setStages] = useState<StageItem[]>([]);
  const [selectedStage, setSelectedStage] = useState<StageItem | null>(null);
  const [showAddDialog, setShowAddDialog] = useState(false);
  const [newStage, setNewStage] = useState<Partial<NewStageItem>>({
    StageColor: defaultColor,
    ProjectNumber: project.ProjectNumber
  });

  const ensureProjectStage = async (): Promise<void> => {
    const items = await sp.web.lists.getByTitle(listName).items
      .filter(`ProjectNumber eq '${project.ProjectNumber}'`).top(5000)();
    if (!items.length) {
      await sp.web.lists.getByTitle(listName).items.add({
        Title: 'Initial Stage',
        StageColor: defaultColor,
        StartDate: new Date().toISOString(),
        EndDate: new Date().toISOString(),
        Status: 'Not Started',
      });
    }
  };

  const fetchStages = async (): Promise<void> => {
    try {
      await ensureProjectStage();
      const results = await sp.web.lists.getByTitle(listName).items
        .filter(`ProjectNumber eq '${project.ProjectNumber}'`).top(5000)();
      setStages(results);
      
      if (selectedStageId) {
        const selected = results.find(stage => stage.Id === selectedStageId);
        if (selected) {
          setSelectedStage(selected);
        }
      }
    } catch (error) {
      console.error('Error fetching stages:', error);
    }
  };

  useEffect(() => {
    fetchStages().catch(console.error);
  }, [project, selectedStageId]);

  const handleAddStage = async (): Promise<void> => {
    if (newStage.Title && newStage.StartDate && newStage.EndDate) {
      try {
        await sp.web.lists.getByTitle(listName).items.add({
          Title: newStage.Title,
          ProjectNumber: project.ProjectNumber,
          StartDate: newStage.StartDate,
          EndDate: newStage.EndDate,
          StageColor: newStage.StageColor || defaultColor,
          Notes: newStage.Notes,
          Status: 'Not Started',
        });
        setShowAddDialog(false);
        setNewStage({ 
          StageColor: defaultColor,
          ProjectNumber: project.ProjectNumber 
        });
        await fetchStages();
      } catch (error) {
        console.error('Error adding stage:', error);
      }
    }
  };

  return (
    <div className={styles.root}>
      <div className={styles.header}>
        <Text size={500} weight="semibold">Project Stages</Text>
        <Dialog open={showAddDialog} onOpenChange={(_, { open }) => setShowAddDialog(open)}>
          <DialogTrigger disableButtonEnhancement>
            <Button icon={<Add24Regular />}>Add Stage</Button>
          </DialogTrigger>
          <DialogSurface>
            <DialogBody>
              <DialogContent className={styles.dialogContent}>
                <Field label="Stage Name" required>
                  <Input
                    value={newStage.Title || ''}
                    onChange={(e) => setNewStage({ ...newStage, Title: e.target.value })}
                  />
                </Field>
                <Field label="Start Date" required>
                  <DatePicker
                    value={newStage.StartDate ? new Date(newStage.StartDate) : null}
                    onSelectDate={(date) => {
                      if (date) {
                        setNewStage({ ...newStage, StartDate: date.toISOString() });
                      }
                    }}
                  />
                </Field>
                <Field label="End Date" required>
                  <DatePicker
                    value={newStage.EndDate ? new Date(newStage.EndDate) : null}
                    onSelectDate={(date) => {
                      if (date) {
                        setNewStage({ ...newStage, EndDate: date.toISOString() });
                      }
                    }}
                  />
                </Field>
                <Field label="Stage Color">
                  <div className={styles.colorPickerWrapper}>
                    <div
                      className={styles.colorButton}
                      style={{ backgroundColor: newStage.StageColor || defaultColor }}
                    />
                    <input
                      type="color"
                      className={styles.colorInput}
                      value={newStage.StageColor || defaultColor}
                      onChange={(e) => setNewStage({ ...newStage, StageColor: e.target.value })}
                    />
                  </div>
                </Field>
                <Field label="Notes">
                  <Input
                    value={newStage.Notes || ''}
                    onChange={(e) => setNewStage({ ...newStage, Notes: e.target.value })}
                  />
                </Field>
              </DialogContent>
              <DialogActions>
                <Button appearance="secondary" onClick={() => setShowAddDialog(false)}>Cancel</Button>
                <Button appearance="primary" onClick={handleAddStage}>Add Stage</Button>
              </DialogActions>
            </DialogBody>
          </DialogSurface>
        </Dialog>
      </div>

      <div className={styles.stageList}>
        {stages.map((stage) => (
          <div
            key={stage.Id}
            className={styles.stageItem}
            onClick={() => setSelectedStage(stage)}
          >
            <div
              className={styles.colorDot}
              style={{ backgroundColor: stage.StageColor || defaultColor }}
            />
            <Text>{stage.Title}</Text>
          </div>
        ))}
      </div>

      {selectedStage && (
        <div className={styles.stageDetails}>
          <Text size={400} weight="semibold">{selectedStage.Title}</Text>
          <Text>Start Date: {new Date(selectedStage.StartDate).toLocaleDateString()}</Text>
          <Text>End Date: {new Date(selectedStage.EndDate).toLocaleDateString()}</Text>
          <Text>Status: {selectedStage.Status}</Text>
          {selectedStage.Notes && <Text>Notes: {selectedStage.Notes}</Text>}
        </div>
      )}
    </div>
  );
};

export default StagesTab;
