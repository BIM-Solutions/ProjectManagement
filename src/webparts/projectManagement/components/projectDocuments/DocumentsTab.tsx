import * as React from "react";
import { spfi, SPFx } from "@pnp/sp";
import { useEffect, useState, useCallback, useMemo } from "react";
import {
  Text,
  Button,
  makeStyles,
  tokens,
  Toaster,
  useId,
  useToastController,
  Toast,
  ToastTitle,
  ToastBody,
  Badge,
  Tooltip,
} from "@fluentui/react-components";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { DocumentService, IDocument } from "../../services/DocumentService";
import { Project } from "../../services/ProjectSelectionServices";
import { mappedDocuments } from "./FolderStructure";
import { Eye24Regular } from "@fluentui/react-icons";
import { eventService } from "../../services/EventService";

interface DocumentsTabProps {
  project: Project;
  context: WebPartContext;
}

interface DocumentTemplate {
  DocumentType: string;
  uniclassCode1: string;
  description1: string;
  uniclassCode2: string;
  description2: string;
  uniclassCode3: string;
  description3: string;
}

export interface SharePointDocumentTemplate {
  DocumentType: string;
  uniclassCode1: string;
  description1: string;
  uniclassCode2: string;
  description2: string;
  uniclassCode3: string;
  description3: string;
}

const useStyles = makeStyles({
  container: {
    display: "flex",
    flexDirection: "column",
    gap: tokens.spacingVerticalL,
    paddingTop: tokens.spacingVerticalL,
  },
  docRow: {
    display: "flex",
    flexDirection: "row",
    alignItems: "flex-start",
    gap: tokens.spacingHorizontalM,
    padding: tokens.spacingVerticalS,
    borderBottom: `1px solid ${tokens.colorNeutralStroke1}`,
  },
  fieldLabel: {
    width: "250px",
    fontWeight: tokens.fontWeightSemibold,
    paddingTop: tokens.spacingVerticalS,
  },
  fieldValue: {
    flexGrow: 1,
    fontStyle: "italic",
    whiteSpace: "pre-wrap",
    wordBreak: "break-word",
    maxWidth: "400px",
  },
  statusBadge: {
    marginLeft: tokens.spacingHorizontalS,
    marginTop: tokens.spacingVerticalS,
  },
  actions: {
    display: "flex",
    gap: tokens.spacingHorizontalS,
    marginTop: tokens.spacingVerticalS,
  },
  descriptionTooltip: {
    maxWidth: "400px",
    whiteSpace: "pre-wrap",
    wordBreak: "break-word",
  },
  actionButton: {
    minWidth: "auto",
    padding: "4px 8px",
  },
});

const DocumentsTab: React.FC<DocumentsTabProps> = ({ context, project }) => {
  const sp = spfi().using(SPFx(context));
  const styles = useStyles();
  const toasterId = useId("doc-toast");
  const { dispatchToast } = useToastController();
  const [documents, setDocuments] = useState<IDocument[]>([]);
  const documentService = useMemo(
    () => new DocumentService(context),
    [context]
  );
  const [documentTypes, setDocumentTypes] = useState<DocumentTemplate[]>([]);

  const getDocumentTemplates = async (): Promise<void> => {
    const templates = await sp.web.lists
      .getByTitle("9719_ProjectDocumentTemplates")
      .items.top(100)();
    setDocumentTypes(
      templates.map((template: SharePointDocumentTemplate) => ({
        DocumentType: template.DocumentType,
        uniclassCode1: template.uniclassCode1,
        description1: template.description1,
        uniclassCode2: template.uniclassCode2,
        description2: template.description2,
        uniclassCode3: template.uniclassCode3,
        description3: template.description3,
      }))
    );
  };

  const fetchDocumentData = useCallback(async (): Promise<void> => {
    try {
      const docs = await documentService.getProjectDocuments(
        project.ProjectNumber
      );
      setDocuments(docs);
      await getDocumentTemplates();
    } catch (error) {
      console.error("Error fetching documents:", error);
      dispatchToast(
        <Toast>
          <ToastTitle>Error</ToastTitle>
          <ToastBody>Failed to load project documents.</ToastBody>
        </Toast>,
        { intent: "error", timeout: 3000 }
      );
    }
  }, [project.ProjectNumber, dispatchToast]);

  // Single effect to handle both initial load and document upload subscription
  useEffect(() => {
    let isSubscribed = true;

    const loadData = async (): Promise<void> => {
      if (isSubscribed) {
        try {
          await fetchDocumentData();
        } catch (error) {
          console.error("Error loading documents:", error);
        }
      }
    };

    // Initial load
    loadData().catch(console.error);

    // Subscribe to document uploads
    const handleDocumentUpload = (): void => {
      if (isSubscribed) {
        loadData().catch(console.error);
      }
    };

    eventService.subscribeToDocumentUpload(handleDocumentUpload);

    // Cleanup
    return () => {
      isSubscribed = false;
      eventService.unsubscribeFromDocumentUpload(handleDocumentUpload);
      setDocuments([]);
    };
  }, [fetchDocumentData]);

  const getDocumentForType = (docType: string): IDocument | undefined => {
    const mappedDoc = mappedDocuments.find((doc) => doc.name === docType);
    if (!mappedDoc) return undefined;
    return documents.find(
      (doc) => doc.UniclassCode3 === mappedDoc.code3 && !doc.IsFolder
    );
  };

  const isDocumentAvailable = (docType: string): boolean => {
    const mappedDoc = mappedDocuments.find((doc) => doc.name === docType);
    if (!mappedDoc) return false;
    return documents.some(
      (doc) => doc.UniclassCode3 === mappedDoc.code3 && !doc.IsFolder
    );
  };

  const getDocumentStatus = (
    docType: string
  ): {
    status: "available" | "not-available" | "incorporated";
    label: string;
  } => {
    const doc = getDocumentForType(docType);
    if (doc && !doc.IsFolder) {
      return { status: "available", label: "Available" };
    }

    if (isDocumentAvailable(docType)) {
      return { status: "available", label: "Available" };
    }

    return { status: "not-available", label: "Not Available" };
  };

  const truncateDescription = (description: string): string => {
    if (!description) return "";
    return description.length > 100
      ? `${description.substring(0, 100)}...`
      : description;
  };

  const handleViewDocument = (doc: IDocument): void => {
    if (doc.FileRef) {
      window.open(doc.FileRef, "_blank");
    }
  };

  return (
    <>
      <div className={styles.container}>
        <Text size={600} weight="semibold">
          Key Project Documents
        </Text>

        {documents ? (
          <>
            {documentTypes.map((docType) => {
              const doc = getDocumentForType(docType.DocumentType);
              const status = getDocumentStatus(docType.DocumentType);

              return (
                <div key={docType.DocumentType} className={styles.docRow}>
                  <Text className={styles.fieldLabel}>
                    {docType.DocumentType}
                  </Text>
                  <div style={{ flexGrow: 1 }}>
                    {doc && (
                      <Tooltip
                        content={
                          <div className={styles.descriptionTooltip}>
                            {doc.Title}
                          </div>
                        }
                        relationship="label"
                      >
                        <Text className={styles.fieldValue}>
                          {truncateDescription(doc.FileLeafRef || doc.Title)}
                        </Text>
                      </Tooltip>
                    )}
                  </div>
                  {status.status === "available" && (
                    <Button
                      className={styles.actionButton}
                      appearance="subtle"
                      icon={<Eye24Regular />}
                      onClick={() => {
                        const docToView = getDocumentForType(
                          docType.DocumentType
                        );
                        // console.log('the docType is: ', docType);
                        // console.log('the docToView is: ', docToView);
                        if (docToView) handleViewDocument(docToView);
                      }}
                    />
                  )}
                  <Badge
                    className={styles.statusBadge}
                    color={
                      status.status === "incorporated"
                        ? "success"
                        : status.status === "available"
                        ? "success"
                        : "danger"
                    }
                  >
                    {status.label}
                  </Badge>
                </div>
              );
            })}
          </>
        ) : (
          <Text>Loading documents...</Text>
        )}
      </div>

      <Toaster toasterId={toasterId} />
    </>
  );
};

export default DocumentsTab;
