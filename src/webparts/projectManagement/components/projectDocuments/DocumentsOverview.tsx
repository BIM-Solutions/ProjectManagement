import * as React from 'react';
import { useState, useEffect, useRef, useMemo } from 'react';
import {
  makeStyles,
  tokens,
  Button,
  Table,
  TableHeader,
  TableRow,
  TableHeaderCell,
  TableBody,
  TableCell,
  Dialog,
  DialogSurface,
  DialogBody,
  DialogTitle,
  DialogContent,
  DialogActions,
  Field,
  // Dropdown,
  // Option,
  Input,
  Text,
  // DialogTrigger,
  Tooltip,
  DialogTrigger,
} from "@fluentui/react-components";
import {
  Folder24Regular,
  Document24Regular,
  DocumentPdf24Regular,
  Image32Regular,
  TextTRegular,
  DocumentRegular,
  Document24Filled,
  Archive24Regular,
  Eye24Regular,
  Delete24Regular,
} from '@fluentui/react-icons';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { Project } from '../../services/ProjectSelectionServices';
import { DocumentService, IDocument } from '../../services/DocumentService';
import { TemplateService } from '../../services/TemplateService';
import { DocumentUpload, DocumentUploadHandle } from './DocumentUpload';
import { SPFI } from '@pnp/sp';
import { eventService } from '../../services/EventService';
import { StandardsWorkflowDialog } from './StandardsWorkflowDialog';
import { StandardsService } from '../../services/StandardsService';


export interface IDocumentsTabProps {
  context: WebPartContext;
  project: Project | undefined;
  sp: SPFI;
  documentService: DocumentService;
  templateService: TemplateService;
}

const useStyles = makeStyles({
  container: {
    display: "flex",
    flexDirection: "column",
    gap: tokens.spacingVerticalM,
  },
  header: {
    display: "flex",
    justifyContent: "space-between",
    alignItems: "center",
  },
  table: {
    width: "100%",
  },
  actions: {
    display: "flex",
    gap: tokens.spacingHorizontalS,
  },
  breadcrumb: {
    display: "flex",
    gap: tokens.spacingHorizontalS,
    alignItems: "center",
    padding: tokens.spacingVerticalS,
  },
  breadcrumbItem: {
    cursor: "pointer",
    "&:hover": {
      textDecoration: "underline",
    },
  },
  icon: {
    marginRight: tokens.spacingHorizontalS,
  },
  actionButton: {
    minWidth: "auto",
    padding: "4px 8px",
  },
});

// const documentTypes = [
//   'Drawing',
//   'Specification',
//   'Report',
//   'Contract',
//   'Correspondence',
//   'Other',
// ];

const getFileIcon = (fileName: string): JSX.Element => {
  const extension = fileName.split(".").pop()?.toLowerCase();

  switch (extension) {
    case "pdf":
      return <DocumentPdf24Regular />;
    case "png":
    case "jpg":
    case "jpeg":
    case "gif":
    case "bmp":
      return <Image32Regular />;
    case "doc":
    case "docx":
      return <Document24Filled />;
    case "xls":
    case "xlsx":
      return <DocumentRegular />;
    case "ppt":
    case "pptx":
      return <DocumentRegular />;
    case "txt":
    case "rtf":
      return <TextTRegular />;
    case "zip":
    case "rar":
    case "7z":
      return <Archive24Regular />;
    default:
      return <Document24Regular />;
  }
};

export const DocumentsOverview: React.FC<IDocumentsTabProps> = ({
  context,
  project,
  sp,
  documentService,
  templateService,
}) => {
  const styles = useStyles();
  const [documents, setDocuments] = useState<IDocument[]>([]);
  const [showNewFolderDialog, setShowNewFolderDialog] = useState(false);
  const [currentPath, setCurrentPath] = useState<string[]>([]);
  const [newFolderName, setNewFolderName] = useState("");
  // const [isUploadDialogOpen, setIsUploadDialogOpen] = useState(false);
  const [uploadDialogKey, setUploadDialogKey] = useState(0);
  const [showStandardsDialog, setShowStandardsDialog] = useState(false);
  const documentsLibrary = 'Projects';
  const uploadRef = useRef<DocumentUploadHandle>(null);
  const standardsService = useMemo(() => new StandardsService(context, sp), [context, sp]);
  

  const getCurrentFolderPath = (): string => {
    if (!project) return "";
    const basePath = `${documentsLibrary}/${project.ProjectNumber}`;
    return currentPath.length > 0
      ? `${basePath}/${currentPath.join("/")}`
      : basePath;
  };

  const loadFolderContents = async (): Promise<void> => {
    if (!project) return;
    try {
      const path = getCurrentFolderPath();
      const docs = await documentService.getFolderContents(path);
      //console.log(docs);
      setDocuments(docs);
    } catch (error) {
      console.error("Error loading documents:", error);
    }
  };

  useEffect(() => {
    loadFolderContents().catch(console.error);
  }, [project, currentPath]);

  const handleCreateFolder = async (): Promise<void> => {
    if (!project || !newFolderName) return;

    try {
      const currentFolder = getCurrentFolderPath();
      await documentService.createFolder(`${currentFolder}/${newFolderName}`);
      setShowNewFolderDialog(false);
      setNewFolderName("");
      await loadFolderContents();
    } catch (error) {
      console.error("Error creating folder:", error);
    }
  };

  // const handleClose = (): void => setIsUploadDialogOpen(false);
  const handleDelete = async (document: IDocument): Promise<void> => {
    try {
      if (document.IsFolder) {
        // Handle folder deletion if needed
      } else {
        await documentService.deleteDocument(document.Id);
      }
      await loadFolderContents();
    } catch (error) {
      console.error("Error deleting item:", error);
    }
  };

  const navigateToFolder = (folderName: string): void => {
    setCurrentPath([...currentPath, folderName]);
  };

  const navigateToBreadcrumb = (index: number): void => {
    setCurrentPath(currentPath.slice(0, index));
  };

  const handleUploadComplete = async (): Promise<void> => {
    await loadFolderContents();
    eventService.notifyDocumentUpload();
  };
  const handleViewDocument = (doc: IDocument): void => {
    console.log("the doc is: ", doc);
    if (doc.FileRef) {
      window.open(doc.FileRef, "_blank");
    }
  };

  return (
    <div className={styles.container}>
      <div className={styles.header}>
        <h2>Project Documents</h2>
        <div className={styles.actions}>
          {/* <Button appearance="primary" onClick={() => setShowStandardsDialog(true)}>
            Standards Workflow
          </Button>
          <Button appearance="primary" onClick={() => setShowNewFolderDialog(true)}>
            New Folder
          </Button>
          <Dialog modalType="modal">
            <DialogTrigger disableButtonEnhancement>
              <Button
                appearance="primary"
                onClick={() => setUploadDialogKey((prev) => prev + 1)}
              >
                Upload Document
              </Button>
            </DialogTrigger>
            <DialogSurface style={{ width: "100%" }}>
              <DialogBody>
                <DialogTitle>Upload Document</DialogTitle>
                <DialogContent id="documentUpload">
                  <DocumentUpload
                    ref={uploadRef}
                    key={uploadDialogKey}
                    sp={sp}
                    projectId={project?.ProjectNumber}
                    libraryName={documentsLibrary}
                    context={context}
                    onUploadComplete={handleUploadComplete}
                    onCancel={() => {
                      setUploadDialogKey((prev) => prev + 1);
                    }}
                  />
                </DialogContent>
                <DialogActions>
                  <Button
                    appearance="primary"
                    onClick={() => uploadRef.current?.upload()}
                  >
                    Upload Document
                  </Button>
                  <Button
                    appearance="secondary"
                    onClick={() => uploadRef.current?.cancel()}
                  >
                    Cancel
                  </Button>
                  <DialogTrigger disableButtonEnhancement>
                    <Button appearance="secondary">Close</Button>
                  </DialogTrigger>
                </DialogActions>
              </DialogBody>
            </DialogSurface>
          </Dialog>
        </div>
      </div>

      <div className={styles.breadcrumb}>
        <Text
          className={styles.breadcrumbItem}
          onClick={() => setCurrentPath([])}
        >
          Root
        </Text>
        {currentPath.map((folder, index) => (
          <React.Fragment key={index}>
            <Text>/</Text>
            <Text
              className={styles.breadcrumbItem}
              onClick={() => navigateToBreadcrumb(index + 1)}
            >
              {folder}
            </Text>
          </React.Fragment>
        ))}
      </div>

      <Table className={styles.table}>
        <TableHeader>
          <TableRow>
            <TableHeaderCell>Name</TableHeaderCell>
            {documents.some((doc) => !doc.IsFolder) && (
              <>
                <TableHeaderCell>Title</TableHeaderCell>
                <TableHeaderCell>Uniclass</TableHeaderCell>
                <TableHeaderCell>Description</TableHeaderCell>
                <TableHeaderCell>Status</TableHeaderCell>
              </>
            )}
            <TableHeaderCell>Modified</TableHeaderCell>
            <TableHeaderCell>Actions</TableHeaderCell>
          </TableRow>
        </TableHeader>
        <TableBody>
          {documents.map((doc) => (
            <TableRow key={doc.FileRef}>
              <TableCell>
                {doc.IsFolder ? (
                  <div
                    style={{
                      display: "flex",
                      alignItems: "center",
                      cursor: "pointer",
                    }}
                    onClick={() => navigateToFolder(doc.Title)}
                  >
                    <Folder24Regular className={styles.icon} />
                    {doc.Title}
                  </div>
                ) : (
                  <div style={{ display: "flex", alignItems: "center" }}>
                    {getFileIcon(doc.Title)}
                    <span style={{ marginLeft: tokens.spacingHorizontalS }}>
                      {doc.Title}
                    </span>
                  </div>
                )}
              </TableCell>
              {!doc.IsFolder && (
                <>
                  <TableCell>{doc.Title || "-"}</TableCell>
                  <TableCell>{doc.UniclassCode3 || "-"}</TableCell>
                  <TableCell>{doc.UniclassTitle3 || "-"}</TableCell>
                  <TableCell>{doc.Status || "-"}</TableCell>
                </>
              )}
              <TableCell>
                {new Date(doc.Modified).toLocaleDateString()}
              </TableCell>
              <TableCell>
                <div className={styles.actions}>
                  {!doc.IsFolder && (
                    <Tooltip content="View" relationship="label">
                      <Button
                        className={styles.actionButton}
                        appearance="subtle"
                        icon={<Eye24Regular />}
                        onClick={() => doc && handleViewDocument(doc)}
                        // onClick={() => window.open(doc.FileRef, '_blank')}
                      />
                    </Tooltip>
                  )}
                  <Tooltip content="Delete" relationship="label">
                    <Button
                      className={styles.actionButton}
                      appearance="subtle"
                      icon={<Delete24Regular />}
                      onClick={() => handleDelete(doc)}
                    />
                  </Tooltip>
                </div>
              </TableCell>
            </TableRow>
          ))}
        </TableBody>
      </Table>

      {/* New Folder Dialog */}
      <Dialog open={showNewFolderDialog}>
        <DialogSurface>
          <DialogBody>
            <DialogTitle>Create New Folder</DialogTitle>
            <DialogContent>
              <div className={styles.container}>
                <Field label="Folder Name">
                  <Input
                    value={newFolderName}
                    onChange={(e) => setNewFolderName(e.target.value)}
                  />
                </Field>
              </div>
            </DialogContent>
            <DialogActions>
              <Button
                appearance="secondary"
                onClick={() => setShowNewFolderDialog(false)}
              >
                Cancel
              </Button>
              <Button
                appearance="primary"
                onClick={handleCreateFolder}
                disabled={!newFolderName}
              >
                Create
              </Button>
            </DialogActions>
          </DialogBody>
        </DialogSurface>
      </Dialog>

      {/* Standards Workflow Dialog */}
      {/* <StandardsWorkflowDialog
        isOpen={showStandardsDialog}
        onDismiss={() => setShowStandardsDialog(false)}
        context={context}
        sp={sp}
        projectNumber={project?.ProjectNumber || ''}
        standardsService={standardsService}
      />
    </div>
  );
};

export default DocumentsOverview;
