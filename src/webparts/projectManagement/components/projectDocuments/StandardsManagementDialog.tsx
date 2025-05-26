import * as React from "react";
import { useState, useEffect } from "react";
import {
  Dialog,
  DialogSurface,
  //   DialogBody,
  //   DialogTitle,
  //   DialogContent,
  //   DialogActions,
  Button,
  makeStyles,
  tokens,
  Text,
  ProgressBar,
  Field,
  Combobox,
  Option,
  //   useComboboxFilter,
  //   useId,
} from "@fluentui/react-components";
// import { ComboBox } from '@fluentui/react';
// import { Option } from '@fluentui/react-components';
import {
  Dismiss24Regular,
  Document24Regular,
  Folder24Regular,
} from "@fluentui/react-icons";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { StandardsService, IStandard } from "../../services/StandardsService";
import { SPFI } from "@pnp/sp";
import { Project } from "../../services/ProjectSelectionServices";

const useStyles = makeStyles({
  stepperBanner: {
    display: "flex",
    justifyContent: "space-between",
    alignItems: "center",
    background: tokens.colorBrandBackground,
    padding: "24px 32px",
    borderRadius: "10px 10px 0 0",
  },

  stepperStep: {
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
    position: "relative",
    flex: 1,
  },

  stepCircle: {
    width: "16px",
    height: "16px",
    borderRadius: "50%",
    background: "#fff",
    border: "3px solid rgba(255,255,255,0.6)",
    marginBottom: "6px",
  },

  stepCircleActive: {
    width: "16px",
    height: "16px",
    borderRadius: "50%",
    background: "#fff",
    border: "3px solid #fff",
    marginBottom: "6px",
  },

  stepLabel: {
    fontSize: "11px",
    color: "rgba(255,255,255,0.7)",
    fontWeight: 400,
    textAlign: "center",
  },

  stepLabelActive: {
    fontSize: "11px",
    color: "#fff",
    fontWeight: 600,
    textAlign: "center",
  },

  stepLine: {
    position: "absolute",
    top: "9px",
    left: "calc(100% - 50px)",
    width: "calc(100% - 8px)",
    height: "2px",
    background: "rgba(255,255,255,0.4)",
    zIndex: 0,
  },
  card: {
    background: "#fff",
    borderRadius: "16px",
    boxShadow: "0 8px 32px rgba(0,0,0,0.1)",
    padding: "0",
    margin: "0 auto",
    maxWidth: "700px",
    minWidth: "350px",
    border: "none",
  },

  cardContent: {
    padding: "32px",
    background: "#fff",
    borderRadius: "0 0 16px 16px",
  },
  dropZone: {
    border: `2px dashed ${tokens.colorBrandStroke1}`,
    borderRadius: "10px",
    padding: "24px",
    textAlign: "center",
    background: "#f8fafc",
    color: tokens.colorNeutralForeground2,
    marginBottom: "16px",
    cursor: "pointer",
    transition: "background 0.2s ease",
  },
  dropZoneActive: {
    background: "#e0e0e0",
  },

  fileList: {
    padding: 0,
    listStyle: "none",
    marginTop: "12px",
    border: "1px solid #e0e0e0",
    borderRadius: "8px",
    background: "#f9f9f9",
  },
  fileItem: {
    padding: "12px",
    borderBottom: "1px solid #e0e0e0",
    display: "flex",
    alignItems: "center",
    justifyContent: "space-between",
  },
  removeBtn: {
    background: "none",
    border: "none",
    cursor: "pointer",
    padding: 0,
    marginLeft: "8px",
  },
  stepContent: {
    padding: "32px",
    background: "#fff",
    borderRadius: "0 0 16px 16px",
  },
  navActions: {
    display: "flex",
    justifyContent: "space-between",
    marginTop: "24px",
  },
  dialogBox: {
    padding: "0",
    borderRadius: "10px",
  },
  initialScreen: {
    display: "flex",
    gap: tokens.spacingHorizontalL,
    justifyContent: "center",
    alignItems: "center",
    padding: tokens.spacingVerticalL,
  },
  optionButton: {
    flex: 1,
    maxWidth: "300px",
    height: "200px",
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
    justifyContent: "center",
    gap: tokens.spacingVerticalM,
    padding: tokens.spacingVerticalL,
    borderRadius: "12px",
    border: `2px solid ${tokens.colorBrandStroke1}`,
    background: tokens.colorNeutralBackground1,
    cursor: "pointer",
    transition: "all 0.2s ease",
    "&:hover": {
      background: tokens.colorBrandBackground2,
      transform: "translateY(-2px)",
    },
  },
  optionIcon: {
    fontSize: "48px",
    color: tokens.colorBrandForeground1,
  },
  optionTitle: {
    fontSize: "20px",
    fontWeight: "600",
    color: tokens.colorNeutralForeground1,
  },
  optionDescription: {
    fontSize: "14px",
    color: tokens.colorNeutralForeground2,
    textAlign: "center",
  },
  clientVersionContainer: {
    display: "flex",
    flexDirection: "column",
    gap: tokens.spacingVerticalM,
    marginBottom: tokens.spacingVerticalL,
  },
  previewContainer: {
    border: `1px solid ${tokens.colorNeutralStroke1}`,
    borderRadius: "8px",
    padding: tokens.spacingVerticalM,
    marginTop: tokens.spacingVerticalM,
    maxHeight: "300px",
    overflowY: "auto",
  },
  previewItem: {
    display: "flex",
    alignItems: "center",
    gap: tokens.spacingHorizontalS,
    padding: tokens.spacingVerticalS,
    borderBottom: `1px solid ${tokens.colorNeutralStroke1}`,
    "&:last-child": {
      borderBottom: "none",
    },
  },
});

const steps = [
  "Select Documents",
  "Uniclass Codes",
  "Review",
  "Upload to Library",
  "Add to Project",
];

interface StandardsManagementDialogProps {
  isOpen: boolean;
  onDismiss: () => void;
  context: WebPartContext;
  sp: SPFI;
  selectedProject: Project | undefined;
  onUploadComplete: () => void;
}

interface IUniclassItem {
  code: string;
  title: string;
  subgroups?: IUniclassItem[];
  sections?: IUniclassItem[];
}

interface UploadedFile {
  name: string; // editable filename
  file: File;
  title?: string;
  uniclassCode?: string;
  uniclassTitle?: string;
  description?: string;
}

export const StandardsManagementDialog: React.FC<
  StandardsManagementDialogProps
> = ({ isOpen, onDismiss, context, selectedProject, sp, onUploadComplete }) => {
  const styles = useStyles();
  const standardsService = new StandardsService(context, sp);
  const [currentStep, setCurrentStep] = useState(0);
  const [workflowType, setWorkflowType] = useState<
    "new" | "existing" | "choose"
  >("choose");
  const [selectedClient, setSelectedClient] = useState<string>("");
  const [selectedVersion, setSelectedVersion] = useState<string>("");
  const [availableClients, setAvailableClients] = useState<string[]>([]);
  const [availableVersions, setAvailableVersions] = useState<string[]>([]);
  const [previewStandards, setPreviewStandards] = useState<IStandard[]>([]);
  const [uniclassData, setUniclassData] = useState<IUniclassItem[]>([]);
  const [files, setFiles] = useState<UploadedFile[]>([]);
  const [dragActive, setDragActive] = useState(false);
  const [queries, setQueries] = useState<string[]>([]);
  const [uploading, setUploading] = useState(false);
  const [uploadProgress, setUploadProgress] = useState(0);
  const [uploadComplete, setUploadComplete] = useState(false);
  const [client, setClient] = useState("");
  const [version, setVersion] = useState("");
  const [copying, setCopying] = useState(false);
  const [copyProgress, setCopyProgress] = useState(0);
  const [copySummary, setCopySummary] = useState<
    { fileName: string; targetPath: string }[]
  >([]);
  const [copyComplete, setCopyComplete] = useState(false);

  const resetDialog = (): void => {
    setWorkflowType("choose");
    setCurrentStep(-1);
    setFiles([]);
    setQueries([]);
    setUploading(false);
    setUploadProgress(0);
    setUploadComplete(false);
    setClient("");
    setVersion("");
    setCopying(false);
    setCopyProgress(0);
    setCopySummary([]);
    setCopyComplete(false);
    setSelectedClient("");
    setSelectedVersion("");
    setAvailableClients([]);
    setAvailableVersions([]);
    setPreviewStandards([]);
  };

  const handleDismiss = (): void => {
    onDismiss();
    resetDialog();
  };

  useEffect(() => {
    async function initialization(): Promise<void> {
      console.log("initialization", selectedProject);
      const next = await standardsService.getNextVersion(
        selectedProject?.Client || ""
      );
      setVersion(next);
      setClient(selectedProject?.Client || "");
    }
    if (isOpen && selectedProject?.Client) {
      initialization().catch(console.error);
    }
  }, [isOpen, selectedProject?.Client]);

  const loadAvailableClients = async (): Promise<void> => {
    try {
      const standards = await standardsService.getExistingStandards();
      const clients = Array.from(new Set(standards.map((s) => s.Client)));
      setAvailableClients(clients);
    } catch (error) {
      console.error("Error loading clients:", error);
    }
  };

  const loadAvailableVersions = async (client: string): Promise<void> => {
    try {
      const standards = await standardsService.getExistingStandards();
      const versions = Array.from(
        new Set(
          standards.filter((s) => s.Client === client).map((s) => s.Version)
        )
      );
      setAvailableVersions(versions);
    } catch (error) {
      console.error("Error loading versions:", error);
    }
  };

  const loadStandardsPreview = async (): Promise<void> => {
    try {
      const standards = await standardsService.getExistingStandards();
      const filteredStandards = standards.filter(
        (s) => s.Client === selectedClient && s.Version === selectedVersion
      );
      setPreviewStandards(filteredStandards);
    } catch (error) {
      console.error("Error loading standards preview:", error);
    }
  };

  // Load available clients when dialog opens
  useEffect(() => {
    if (isOpen && workflowType === "existing") {
      loadAvailableClients().catch(console.error);
    }
  }, [isOpen, workflowType]);

  // Load versions when client is selected
  useEffect(() => {
    if (selectedClient) {
      loadAvailableVersions(selectedClient).catch(console.error);
    }
  }, [selectedClient]);

  // Load standards preview when version is selected
  useEffect(() => {
    if (selectedClient && selectedVersion) {
      loadStandardsPreview().catch(console.error);
    }
  }, [selectedClient, selectedVersion]);

  // Add files, append to existing, avoid duplicates by name
  const addFiles = (newFiles: File[]): void => {
    setFiles((prev) => {
      const existingNames = new Set(prev.map((f) => f.name));
      const toAdd = newFiles
        .filter((f) => !existingNames.has(f.name))
        .map((f) => ({ name: f.name, file: f }));
      return [...prev, ...toAdd];
    });
  };

  // Load Uniclass data on mount
  useEffect(() => {
    const fetchUniclass = async (): Promise<void> => {
      try {
        const response = await fetch(
          `${context.pageContext.web.absoluteUrl}/SiteAssets/Uniclass2015_PM.json`
        );
        if (!response.ok) throw new Error("Failed to load Uniclass data");
        const data = await response.json();
        setUniclassData(data);
      } catch {
        setUniclassData([]);
      }
    };
    fetchUniclass().catch(console.error);
  }, []);

  // File selection handler (Step 1)
  const handleFileInput = (e: React.ChangeEvent<HTMLInputElement>): void => {
    if (e.target.files) {
      addFiles(Array.from(e.target.files));
    }
  };

  // Drag and drop handlers
  const handleDrop = (e: React.DragEvent<HTMLDivElement>): void => {
    e.preventDefault();
    setDragActive(false);
    if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
      addFiles(Array.from(e.dataTransfer.files));
    }
  };
  const handleDragOver = (e: React.DragEvent<HTMLDivElement>): void => {
    e.preventDefault();
    setDragActive(true);
  };
  const handleDragLeave = (e: React.DragEvent<HTMLDivElement>): void => {
    e.preventDefault();
    setDragActive(false);
  };

  // Remove file by index
  const handleRemoveFile = (idx: number): void => {
    setFiles((prev) => prev.filter((_, i) => i !== idx));
  };

  // Uniclass flat list for ComboBox
  const getUniclassOptions = (): {
    children: React.ReactNode;
    value: string;
    text: string;
    title: string;
  }[] => {
    const options: {
      children: React.ReactNode;
      value: string;
      text: string;
      title: string;
    }[] = [];
    uniclassData.forEach((l1) => {
      options.push({
        value: l1.code,
        children: `${l1.code} - ${l1.title}`,
        text: `${l1.code} - ${l1.title}`,
        title: l1.title,
      });
      l1.subgroups?.forEach((l2) => {
        options.push({
          value: l2.code,
          children: `${l2.code} - ${l2.title}`,
          text: `${l2.code} - ${l2.title}`,
          title: l2.title,
        });
        l2.sections?.forEach((l3) => {
          options.push({
            value: l3.code,
            children: `${l3.code} - ${l3.title}`,
            text: `${l3.code} - ${l3.title}`,
            title: l3.title,
          });
        });
      });
    });
    return options;
  };

  const goToNext = (): void =>
    setCurrentStep((s) => Math.min(s + 1, steps.length - 1));
  const goToPrev = (): void => {
    setCurrentStep((s) => Math.max(s - 1, 0));
    if (currentStep === 0) {
      setWorkflowType("choose");
    }
  };

  const renderStepContent = (): React.ReactNode => {
    if (workflowType === "choose") {
      return (
        <div className={styles.initialScreen}>
          <div
            className={styles.optionButton}
            onClick={() => setWorkflowType("new")}
          >
            <Document24Regular className={styles.optionIcon} />
            <Text className={styles.optionTitle}>New Standards</Text>
            <Text className={styles.optionDescription}>
              Upload and configure new standards documents
            </Text>
          </div>
          <div
            className={styles.optionButton}
            onClick={() => setWorkflowType("existing")}
          >
            <Folder24Regular className={styles.optionIcon} />
            <Text className={styles.optionTitle}>Existing Standards</Text>
            <Text className={styles.optionDescription}>
              Select and copy existing standards to your project
            </Text>
          </div>
        </div>
      );
    }

    if (workflowType === "existing") {
      return (
        <div>
          <div className={styles.clientVersionContainer}>
            <Field label="Select Client">
              <Combobox
                value={selectedClient}
                onOptionSelect={(_, data) => {
                  if (data.optionText) {
                    setSelectedClient(data.optionText);
                  }
                }}
              >
                {availableClients.map((client) => (
                  <Option key={client} text={client}>
                    {client}
                  </Option>
                ))}
              </Combobox>
            </Field>

            {selectedClient && (
              <Field label="Select Version">
                <Combobox
                  value={selectedVersion}
                  onOptionSelect={(_, data) => {
                    if (data.optionText) {
                      setSelectedVersion(data.optionText);
                    }
                  }}
                >
                  {availableVersions.map((version) => (
                    <Option key={version} text={version}>
                      {version}
                    </Option>
                  ))}
                </Combobox>
              </Field>
            )}
          </div>

          {previewStandards.length > 0 && (
            <div className={styles.previewContainer}>
              <Text
                weight="semibold"
                style={{ marginBottom: tokens.spacingVerticalS }}
              >
                Available Standards
              </Text>
              {previewStandards.map((standard) => (
                <div key={standard.Id} className={styles.previewItem}>
                  <Document24Regular />
                  <div>
                    <Text weight="semibold">{standard.Title}</Text>
                    <Text size={200} color="neutralForeground2">
                      {standard.UniclassCode3} - {standard.UniclassTitle3}
                    </Text>
                  </div>
                </div>
              ))}
            </div>
          )}

          {selectedClient && selectedVersion && (
            <Button
              appearance="primary"
              onClick={async () => {
                setCopying(true);
                setCopyProgress(0);
                setCopySummary([]);
                setCopyComplete(false);

                const filesToCopy = previewStandards.map((s) => ({
                  fileName: s.File?.Name || "",
                  uniclassCode: s.UniclassCode3 || "",
                  client: selectedClient,
                  version: selectedVersion,
                }));

                await standardsService.addStandardsToProjectWithProgress(
                  filesToCopy,
                  selectedProject?.ProjectNumber || "",
                  context,
                  sp,
                  (progress, summary) => {
                    setCopyProgress(progress);
                    setCopySummary(summary);
                  }
                );
                setCopying(false);
                setCopyComplete(true);
                // Wait for folder contents to be reloaded
                await new Promise((resolve) => setTimeout(resolve, 1000));
                onUploadComplete();
                onDismiss();
              }}
              disabled={copying || copyComplete}
            >
              Copy to Project
            </Button>
          )}
        </div>
      );
    }

    switch (currentStep) {
      case 0:
        return (
          <div>
            <Text>Select or upload multiple documents:</Text>
            <div
              className={
                dragActive
                  ? `${styles.dropZone} ${styles.dropZoneActive}`
                  : styles.dropZone
              }
              onDrop={handleDrop}
              onDragOver={handleDragOver}
              onDragLeave={handleDragLeave}
              onClick={() => document.getElementById("fileInput")?.click()}
            >
              Drag and drop files here, or click to select
              <input
                id="fileInput"
                type="file"
                multiple
                onChange={handleFileInput}
                style={{ display: "none" }}
              />
            </div>
            {files.length > 0 && (
              <ul className={styles.fileList}>
                {files.map((file, idx) => (
                  <li key={idx} className={styles.fileItem}>
                    {file.name}
                    <button
                      className={styles.removeBtn}
                      title="Remove"
                      onClick={(e) => {
                        e.stopPropagation();
                        handleRemoveFile(idx);
                      }}
                    >
                      <Dismiss24Regular />
                    </button>
                  </li>
                ))}
              </ul>
            )}
          </div>
        );
      case 1:
        return (
          <div>
            <Text>Assign Uniclass codes to each document:</Text>
            <ul className={styles.fileList}>
              {files.map((file, idx) => {
                const comboId = `uniclass-combo-${idx}`;
                const query = queries[idx] ?? file.uniclassCode ?? "";
                // Filter options manually for this file
                const options = getUniclassOptions().filter((opt) =>
                  String(opt.children)
                    .toLowerCase()
                    .includes(query.toLowerCase())
                );
                return (
                  <li key={idx} className={styles.fileItem}>
                    <div
                      style={{
                        display: "flex",
                        flexDirection: "column",
                        gap: 8,
                        flex: 1,
                      }}
                    >
                      <Field label="Title" style={{ minWidth: 200 }}>
                        <input
                          type="text"
                          value={file.title || ""}
                          onChange={(e) => {
                            const value = e.target.value;
                            setFiles((prev) =>
                              prev.map((f, i) =>
                                i === idx ? { ...f, title: value } : f
                              )
                            );
                          }}
                        />
                      </Field>
                      <Field label="Filename" style={{ minWidth: 200 }}>
                        <input
                          type="text"
                          value={file.name}
                          onChange={(e) => {
                            const value = e.target.value;
                            setFiles((prev) =>
                              prev.map((f, i) =>
                                i === idx ? { ...f, name: value } : f
                              )
                            );
                          }}
                        />
                      </Field>
                      <Field label="Uniclass Code" style={{ minWidth: 320 }}>
                        <Combobox
                          aria-labelledby={comboId}
                          placeholder="Select or search Uniclass"
                          value={query}
                          onChange={(ev) => {
                            const value = ev.target.value;
                            setQueries((prev) => {
                              const arr = [...prev];
                              arr[idx] = value;
                              return arr;
                            });
                          }}
                          onOptionSelect={(_e, data) => {
                            setQueries((prev) => {
                              const arr = [...prev];
                              arr[idx] = data.optionText ?? "";
                              return arr;
                            });
                            // Find the matching option to get its title
                            const selectedOption = getUniclassOptions().find(
                              (opt) => opt.value === data.optionValue
                            );
                            setFiles((prev) =>
                              prev.map((f, i) =>
                                i === idx
                                  ? {
                                      ...f,
                                      uniclassCode: data.optionValue ?? "",
                                      title: selectedOption?.title ?? f.title,
                                    }
                                  : f
                              )
                            );
                          }}
                        >
                          {options.length > 0 ? (
                            options.map((opt) => (
                              <Option
                                key={opt.value}
                                text={opt.text}
                                value={opt.value}
                              >
                                {opt.children}
                              </Option>
                            ))
                          ) : (
                            <Option disabled>
                              No Uniclass codes match your search.
                            </Option>
                          )}
                        </Combobox>
                      </Field>
                    </div>
                    <button
                      className={styles.removeBtn}
                      title="Remove"
                      onClick={(e) => {
                        e.stopPropagation();
                        handleRemoveFile(idx);
                      }}
                    >
                      <Dismiss24Regular />
                    </button>
                  </li>
                );
              })}
            </ul>
          </div>
        );
      case 2:
        return (
          <div>
            <Text>Review and validate all documents before upload:</Text>
            <table
              style={{
                width: "100%",
                borderCollapse: "collapse",
                marginTop: 16,
              }}
            >
              <thead>
                <tr>
                  <th
                    style={{
                      textAlign: "left",
                      borderBottom: "1px solid #ccc",
                      padding: 8,
                    }}
                  >
                    Filename
                  </th>
                  <th
                    style={{
                      textAlign: "left",
                      borderBottom: "1px solid #ccc",
                      padding: 8,
                    }}
                  >
                    Title
                  </th>
                  <th
                    style={{
                      textAlign: "left",
                      borderBottom: "1px solid #ccc",
                      padding: 8,
                    }}
                  >
                    Uniclass Code
                  </th>
                  <th
                    style={{
                      textAlign: "left",
                      borderBottom: "1px solid #ccc",
                      padding: 8,
                    }}
                  >
                    Actions
                  </th>
                </tr>
              </thead>
              <tbody>
                {files.map((file, idx) => {
                  const missing =
                    !file.name || !file.title || !file.uniclassCode;
                  return (
                    <tr
                      key={idx}
                      style={{ background: missing ? "#fffbe6" : undefined }}
                    >
                      <td style={{ padding: 8 }}>
                        {file.name || (
                          <span style={{ color: "red" }}>Missing filename</span>
                        )}
                      </td>
                      <td style={{ padding: 8 }}>
                        {file.title || (
                          <span style={{ color: "red" }}>Missing title</span>
                        )}
                      </td>
                      <td style={{ padding: 8 }}>
                        {file.uniclassCode || (
                          <span style={{ color: "red" }}>Missing Uniclass</span>
                        )}
                      </td>
                      <td style={{ padding: 8 }}>
                        <Button size="small" onClick={() => goToPrev()}>
                          Edit
                        </Button>
                        <Button
                          size="small"
                          appearance="subtle"
                          onClick={() => handleRemoveFile(idx)}
                        >
                          Remove
                        </Button>
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
            {files.some((f) => !f.name || !f.title || !f.uniclassCode) && (
              <div style={{ color: "red", marginTop: 12 }}>
                Please ensure all files have a filename, title, and Uniclass
                code before proceeding.
              </div>
            )}
          </div>
        );
      case 3:
        return (
          <div>
            <Text>
              Upload validated standards to the central Standards Library:
            </Text>
            <div style={{ display: "flex", gap: 16, margin: "12px 0" }}>
              <Field label="Client" style={{ minWidth: 180 }}>
                <input
                  type="text"
                  value={client}
                  onChange={(e) => setClient(e.target.value)}
                  //   placeholder="Client name"
                />
              </Field>
              <Field label="Version" style={{ minWidth: 120 }}>
                <input
                  type="text"
                  value={version}
                  onChange={(e) => setVersion(e.target.value)}
                  placeholder="Version (e.g. V1, P01)"
                />
              </Field>
            </div>
            <ul className={styles.fileList}>
              {files.map((file, idx) => (
                <li key={idx} className={styles.fileItem}>
                  <span>
                    <b>{file.name}</b> ({file.title}) - {file.uniclassCode}
                  </span>
                </li>
              ))}
            </ul>
            <div style={{ margin: "16px 0" }}>
              <ProgressBar value={uploadProgress} />
              <div style={{ marginTop: 8 }}>
                {uploading && !uploadComplete && (
                  <Text>Uploading... {Math.round(uploadProgress * 100)}%</Text>
                )}
                {uploadComplete && (
                  <Text style={{ color: "green" }}>Upload complete!</Text>
                )}
              </div>
            </div>
            <Button
              appearance="primary"
              disabled={uploading || uploadComplete || !client || !version}
              onClick={async () => {
                setUploading(true);
                setUploadProgress(0);
                setUploadComplete(false);
                // Call the real upload method for each file
                for (let i = 0; i < files.length; i++) {
                  const fileObj = files[i];
                  await standardsService.uploadNativeFile(
                    fileObj.file,
                    version,
                    client,
                    selectedProject?.ProjectNumber || "",
                    {
                      uniclassCode: fileObj.uniclassCode,
                      title: fileObj.title,
                      description: fileObj.description,
                    },
                    context
                  );
                  setUploadProgress((i + 1) / files.length);
                }
                setUploading(false);
                setUploadComplete(true);
                onUploadComplete();
              }}
            >
              Upload
            </Button>
            <div style={{ marginTop: 8, color: "#888", fontSize: 13 }}>
              Files will be uploaded to:{" "}
              <b>{selectedProject?.Client || "[Client]"}</b> /{" "}
              <b>{version || "[Version]"}</b>
            </div>
          </div>
        );
      case 4:
        return (
          <div>
            <Text>
              Copy uploaded standards to the project folder using Uniclass
              folder structure:
            </Text>
            {!copying && !copyComplete && (
              <Button
                appearance="primary"
                onClick={async () => {
                  setCopying(true);
                  setCopyProgress(0);
                  setCopySummary([]);
                  setCopyComplete(false);
                  const filesToCopy = files.map((f) => ({
                    fileName: f.name,
                    uniclassCode: f.uniclassCode || "",
                    client,
                    version,
                  }));
                  await standardsService.addStandardsToProjectWithProgress(
                    filesToCopy,
                    selectedProject?.ProjectNumber || "",
                    context,
                    sp,
                    (progress, summary) => {
                      setCopyProgress(progress);
                      setCopySummary(summary);
                    }
                  );
                  setCopying(false);
                  setCopyComplete(true);
                  // Wait for folder contents to be reloaded
                  await new Promise((resolve) => setTimeout(resolve, 1000));
                  onUploadComplete();
                  onDismiss();
                }}
              >
                Add to Project
              </Button>
            )}
            {copying && (
              <div>
                <ProgressBar value={copyProgress} />
                <Text>Copying... {Math.round(copyProgress * 100)}%</Text>
              </div>
            )}
            {copyComplete && (
              <div>
                <Text style={{ color: "green" }}>Copy complete!</Text>
                <ul>
                  {copySummary.map((s) => (
                    <li key={s.fileName}>
                      <b>{s.fileName}</b> →{" "}
                      <span style={{ fontFamily: "monospace" }}>
                        {s.targetPath}
                      </span>
                    </li>
                  ))}
                </ul>
              </div>
            )}
          </div>
        );
      default:
        return null;
    }
  };

  // Only allow Next if all files have a Uniclass code (for step 1)
  const canGoNext = (): boolean => {
    if ((currentStep === 1 || currentStep === 2) && files.length > 0) {
      return files.every((f) => !!f.uniclassCode && !!f.title && !!f.name);
    }
    return true;
  };

  return (
    <Dialog
      open={isOpen}
      onOpenChange={(_, { open }) => !open && handleDismiss()}
    >
      <DialogSurface className={styles.dialogBox}>
        <div className={styles.card}>
          {workflowType === "new" ? (
            <div className={styles.stepperBanner}>
              {steps.slice(0, 5).map((label, idx) => (
                <div key={label} className={styles.stepperStep}>
                  <div
                    className={
                      idx === currentStep
                        ? styles.stepCircleActive
                        : styles.stepCircle
                    }
                  />

                  <div
                    className={
                      idx === currentStep
                        ? styles.stepLabelActive
                        : styles.stepLabel
                    }
                  >
                    {label.split(" ").length > 3
                      ? label.replace(/ (?!.* )/, "\n")
                      : label}
                  </div>
                  {idx < 4 && <div className={styles.stepLine} />}
                </div>
              ))}
            </div>
          ) : (
            <Text className={styles.stepperBanner}>Standards Management</Text>
          )}
          <div className={styles.cardContent}>
            <div className={styles.stepContent}>{renderStepContent()}</div>
            <div className={styles.navActions}>
              <Button
                appearance="secondary"
                onClick={handleDismiss}
                disabled={uploading}
              >
                Cancel
              </Button>
              <Button
                onClick={goToPrev}
                disabled={currentStep === -1 || uploading}
              >
                Back
              </Button>
              <Button
                onClick={goToNext}
                disabled={
                  currentStep === steps.length - 1 || !canGoNext() || uploading
                }
              >
                Next
              </Button>
            </div>
          </div>
        </div>
      </DialogSurface>
    </Dialog>
  );
};

export default StandardsManagementDialog;
