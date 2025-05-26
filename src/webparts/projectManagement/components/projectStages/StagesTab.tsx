import * as React from "react";
import { useEffect, useState } from "react";
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
} from "@fluentui/react-components";
import {
  Add24Regular,
  Edit24Regular,
  Delete24Regular,
} from "@fluentui/react-icons";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { spfi } from "@pnp/sp";
import { SPFx } from "@pnp/sp/presets/all";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { Project } from "../../services/ProjectSelectionServices";
import { DatePicker } from "@fluentui/react-datepicker-compat";

interface StagesTabProps {
  project: Project;
  context: WebPartContext;
  selectedStageId?: number;
  stagesChanged?: () => void;
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

type NewStageItem = Omit<StageItem, "Id">;

const listName = "9719_ProjectStages";
const defaultColor = "#0078D4";

const useStyles = makeStyles({
  root: {
    display: "flex",
    flexDirection: "column",
    gap: tokens.spacingVerticalM,
    padding: tokens.spacingVerticalM,
    backgroundColor: tokens.colorNeutralBackground1,
    borderRadius: tokens.borderRadiusMedium,
    boxShadow: tokens.shadow4,
  },
  header: {
    display: "flex",
    justifyContent: "space-between",
    alignItems: "center",
  },
  stageList: {
    display: "flex",
    flexDirection: "column",
    gap: tokens.spacingVerticalS,
    maxHeight: "calc(100vh - 300px)",
    overflowY: "auto",
  },
  stageItem: {
    display: "flex",
    alignItems: "center",
    padding: tokens.spacingVerticalS,
    borderRadius: tokens.borderRadiusMedium,
    backgroundColor: tokens.colorNeutralBackground2,
    cursor: "pointer",
    "&:hover": {
      backgroundColor: tokens.colorNeutralBackground2Hover,
    },
  },
  colorDot: {
    width: "12px",
    height: "12px",
    borderRadius: "50%",
    marginRight: tokens.spacingHorizontalS,
  },
  dialogContent: {
    display: "flex",
    flexDirection: "column",
    gap: tokens.spacingVerticalM,
  },
  colorPickerWrapper: {
    position: "relative",
    width: "100%",
  },
  colorButton: {
    width: "100%",
    height: "32px",
    ...shorthands.border("1px", "solid", tokens.colorNeutralStroke1),
    ...shorthands.borderRadius(tokens.borderRadiusMedium),
    cursor: "pointer",
    padding: 0,
    "&:hover": {
      ...shorthands.borderColor(tokens.colorNeutralStroke1Hover),
    },
  },
  colorInput: {
    position: "absolute",
    top: 0,
    left: 0,
    width: "100%",
    height: "100%",
    opacity: 0,
    cursor: "pointer",
  },
  stageDetails: {
    display: "flex",
    flexDirection: "column",
    gap: tokens.spacingVerticalS,
    padding: tokens.spacingVerticalM,
    backgroundColor: tokens.colorNeutralBackground2,
    borderRadius: tokens.borderRadiusMedium,
  },
});

const StagesTab: React.FC<StagesTabProps> = ({
  project,
  context,
  selectedStageId,
  stagesChanged,
}) => {
  const sp = spfi().using(SPFx(context));
  const styles = useStyles();

  const [stages, setStages] = useState<StageItem[]>([]);
  const [selectedStage, setSelectedStage] = useState<StageItem | null>(null);
  const [showAddDialog, setShowAddDialog] = useState(false);
  const [newStage, setNewStage] = useState<Partial<NewStageItem>>({
    StageColor: defaultColor,
    ProjectNumber: project.ProjectNumber,
  });
  const [showEditDialog, setShowEditDialog] = useState(false);
  const [editStage, setEditStage] = useState<StageItem | null>(null);
  const [showDeleteDialog, setShowDeleteDialog] = useState(false);
  const [deleteStage, setDeleteStage] = useState<StageItem | null>(null);

  const ensureProjectStage = async (): Promise<void> => {
    const items = await sp.web.lists
      .getByTitle(listName)
      .items.filter(`ProjectNumber eq '${project.ProjectNumber}'`)
      .top(5000)();
    if (!items.length) {
      await sp.web.lists.getByTitle(listName).items.add({
        Title: "Initial Stage",
        StageColor: defaultColor,
        StartDate: new Date().toISOString(),
        EndDate: new Date().toISOString(),
        Status: "Not Started",
      });
    }
  };

  const fetchStages = async (): Promise<void> => {
    try {
      if (!project?.ProjectNumber) {
        console.log("No project selected");
        return;
      }
      await ensureProjectStage();
      const results = await sp.web.lists
        .getByTitle(listName)
        .items.filter(`ProjectNumber eq '${project.ProjectNumber}'`)
        .top(5000)();
      setStages(results);
    } catch (error) {
      console.error("Error fetching stages:", error);
    }
  };

  // Effect to handle stage selection
  useEffect(() => {
    if (selectedStageId && stages.length > 0) {
      setSelectedStage(null);
      const selected = stages.find((stage) => stage.Id === selectedStageId);
      if (selected) {
        setSelectedStage(selected || null);
      }
    }
  }, [selectedStageId, stages]);

  useEffect(() => {
    console.log("task has been changed", project);
    fetchStages().catch(console.error);
    console.log("task has been changed", project);
  }, [project?.ProjectNumber]);

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
          Status: "Not Started",
        });
        setShowAddDialog(false);
        setNewStage({
          StageColor: defaultColor,
          ProjectNumber: project.ProjectNumber,
        });
        await fetchStages();
        stagesChanged?.();
      } catch (error) {
        console.error("Error adding stage:", error);
      }
    }
  };

  const handleEditStage = async (): Promise<void> => {
    if (
      editStage &&
      editStage.Title &&
      editStage.StartDate &&
      editStage.EndDate
    ) {
      try {
        await sp.web.lists
          .getByTitle(listName)
          .items.getById(editStage.Id)
          .update({
            Title: editStage.Title,
            StartDate: editStage.StartDate,
            EndDate: editStage.EndDate,
            StageColor: editStage.StageColor || defaultColor,
            Notes: editStage.Notes,
            Status: editStage.Status,
          });
        setShowEditDialog(false);
        setEditStage(null);
        await fetchStages();
        stagesChanged?.();
      } catch (error) {
        console.error("Error editing stage:", error);
      }
    }
  };

  const handleDeleteStage = async (): Promise<void> => {
    if (deleteStage) {
      try {
        await sp.web.lists
          .getByTitle(listName)
          .items.getById(deleteStage.Id)
          .delete();
        setShowDeleteDialog(false);
        setDeleteStage(null);
        setSelectedStage(null);
        await fetchStages();
        stagesChanged?.();
      } catch (error) {
        console.error("Error deleting stage:", error);
      }
    }
  };

  return (
    <>
      {selectedStage && (
        <div className={styles.root}>
          <div className={styles.stageDetails}>
            <div
              style={{ display: "flex", justifyContent: "flex-end", gap: 8 }}
            >
              <Button
                size="small"
                icon={<Edit24Regular />}
                appearance="subtle"
                aria-label="Edit"
                onClick={() => {
                  setEditStage(selectedStage);
                  setShowEditDialog(true);
                }}
              />
              <Button
                size="small"
                icon={<Delete24Regular />}
                appearance="subtle"
                aria-label="Delete"
                onClick={() => {
                  setDeleteStage(selectedStage);
                  setShowDeleteDialog(true);
                }}
              />
            </div>
            <Text size={400} weight="semibold">
              {selectedStage.Title}
            </Text>
            <Text>
              Start Date:{" "}
              {new Date(selectedStage.StartDate).toLocaleDateString()}
            </Text>
            <Text>
              End Date: {new Date(selectedStage.EndDate).toLocaleDateString()}
            </Text>
            <Text>Status: {selectedStage.Status}</Text>
            {selectedStage.Notes && <Text>Notes: {selectedStage.Notes}</Text>}
          </div>
        </div>
      )}

      <div className={styles.root}>
        <div className={styles.header}>
          <Text size={500} weight="semibold">
            Project Stages
          </Text>
          <Dialog
            open={showAddDialog}
            onOpenChange={(_, { open }) => setShowAddDialog(open)}
          >
            <DialogTrigger disableButtonEnhancement>
              <Button icon={<Add24Regular />}>Add Stage</Button>
            </DialogTrigger>
            <DialogSurface>
              <DialogBody>
                <DialogContent className={styles.dialogContent}>
                  <Field label="Stage Name" required>
                    <Input
                      value={newStage.Title || ""}
                      onChange={(e) =>
                        setNewStage({ ...newStage, Title: e.target.value })
                      }
                    />
                  </Field>
                  <Field label="Start Date" required>
                    <DatePicker
                      value={
                        newStage.StartDate ? new Date(newStage.StartDate) : null
                      }
                      onSelectDate={(date) => {
                        if (date) {
                          setNewStage({
                            ...newStage,
                            StartDate: date.toISOString(),
                          });
                        }
                      }}
                    />
                  </Field>
                  <Field label="End Date" required>
                    <DatePicker
                      value={
                        newStage.EndDate ? new Date(newStage.EndDate) : null
                      }
                      onSelectDate={(date) => {
                        if (date) {
                          setNewStage({
                            ...newStage,
                            EndDate: date.toISOString(),
                          });
                        }
                      }}
                    />
                  </Field>
                  <Field label="Stage Color">
                    <div className={styles.colorPickerWrapper}>
                      <div
                        className={styles.colorButton}
                        style={{
                          backgroundColor: newStage.StageColor || defaultColor,
                        }}
                      />
                      <input
                        type="color"
                        className={styles.colorInput}
                        value={newStage.StageColor || defaultColor}
                        onChange={(e) =>
                          setNewStage({
                            ...newStage,
                            StageColor: e.target.value,
                          })
                        }
                      />
                    </div>
                  </Field>
                  <Field label="Notes">
                    <Input
                      value={newStage.Notes || ""}
                      onChange={(e) =>
                        setNewStage({ ...newStage, Notes: e.target.value })
                      }
                    />
                  </Field>
                </DialogContent>
                <DialogActions>
                  <Button
                    appearance="secondary"
                    onClick={() => setShowAddDialog(false)}
                  >
                    Cancel
                  </Button>
                  <Button appearance="primary" onClick={handleAddStage}>
                    Add Stage
                  </Button>
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
              style={{ justifyContent: "space-between", display: "flex" }}
            >
              <div style={{ display: "flex", alignItems: "center" }}>
                <div
                  className={styles.colorDot}
                  style={{ backgroundColor: stage.StageColor || defaultColor }}
                />
                <Text>{stage.Title}</Text>
              </div>
              <div
                style={{ display: "flex", gap: 8 }}
                onClick={(e) => e.stopPropagation()}
              >
                <Button
                  size="small"
                  icon={<Edit24Regular />}
                  appearance="subtle"
                  aria-label="Edit"
                  onClick={() => {
                    setEditStage(stage);
                    setShowEditDialog(true);
                  }}
                />
                <Button
                  size="small"
                  icon={<Delete24Regular />}
                  appearance="subtle"
                  aria-label="Delete"
                  onClick={() => {
                    setDeleteStage(stage);
                    setShowDeleteDialog(true);
                  }}
                />
              </div>
            </div>
          ))}
        </div>

        {/* Edit Stage Dialog */}
        <Dialog
          open={showEditDialog}
          onOpenChange={(_, { open }) => setShowEditDialog(open)}
        >
          <DialogSurface>
            <DialogBody>
              <DialogContent className={styles.dialogContent}>
                <Field label="Stage Name" required>
                  <Input
                    value={editStage?.Title || ""}
                    onChange={(e) =>
                      setEditStage(
                        editStage
                          ? { ...editStage, Title: e.target.value }
                          : null
                      )
                    }
                  />
                </Field>
                <Field label="Start Date" required>
                  <DatePicker
                    value={
                      editStage?.StartDate
                        ? new Date(editStage.StartDate)
                        : null
                    }
                    onSelectDate={(date) => {
                      if (date && editStage) {
                        setEditStage({
                          ...editStage,
                          StartDate: date.toISOString(),
                        });
                      }
                    }}
                  />
                </Field>
                <Field label="End Date" required>
                  <DatePicker
                    value={
                      editStage?.EndDate ? new Date(editStage.EndDate) : null
                    }
                    onSelectDate={(date) => {
                      if (date && editStage) {
                        setEditStage({
                          ...editStage,
                          EndDate: date.toISOString(),
                        });
                      }
                    }}
                  />
                </Field>
                <Field label="Stage Color">
                  <div className={styles.colorPickerWrapper}>
                    <div
                      className={styles.colorButton}
                      style={{
                        backgroundColor: editStage?.StageColor || defaultColor,
                      }}
                    />
                    <input
                      type="color"
                      className={styles.colorInput}
                      value={editStage?.StageColor || defaultColor}
                      onChange={(e) =>
                        setEditStage(
                          editStage
                            ? { ...editStage, StageColor: e.target.value }
                            : null
                        )
                      }
                    />
                  </div>
                </Field>
                <Field label="Notes">
                  <Input
                    value={editStage?.Notes || ""}
                    onChange={(e) =>
                      setEditStage(
                        editStage
                          ? { ...editStage, Notes: e.target.value }
                          : null
                      )
                    }
                  />
                </Field>
              </DialogContent>
              <DialogActions>
                <Button
                  appearance="secondary"
                  onClick={() => setShowEditDialog(false)}
                >
                  Cancel
                </Button>
                <Button appearance="primary" onClick={handleEditStage}>
                  Save Changes
                </Button>
              </DialogActions>
            </DialogBody>
          </DialogSurface>
        </Dialog>

        {/* Delete Confirmation Dialog */}
        <Dialog
          open={showDeleteDialog}
          onOpenChange={(_, { open }) => setShowDeleteDialog(open)}
        >
          <DialogSurface>
            <DialogBody>
              <DialogContent>
                <Text>
                  Are you sure you want to delete the stage "
                  {deleteStage?.Title}
                  "?
                </Text>
              </DialogContent>
              <DialogActions>
                <Button
                  appearance="secondary"
                  onClick={() => setShowDeleteDialog(false)}
                >
                  Cancel
                </Button>
                <Button appearance="primary" onClick={handleDeleteStage}>
                  Delete
                </Button>
              </DialogActions>
            </DialogBody>
          </DialogSurface>
        </Dialog>
      </div>
    </>
  );
};

export default StagesTab;
