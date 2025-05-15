import * as React from 'react';
import { useState, useEffect } from 'react';
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
//   DialogTrigger,
  DialogSurface,
  DialogBody,
  DialogTitle,
  DialogContent,
  DialogActions,
//   Input,
  Field,
  Dropdown,
  Option,
} from '@fluentui/react-components';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { IProject } from '../../services/ProjectService';
import { DocumentService, IDocument } from '../../services/DocumentService';
import { TemplateService } from '../../services/TemplateService';

export interface IDocumentsTabProps {
  context: WebPartContext;
  project: IProject | undefined;
  documentService: DocumentService;
  templateService: TemplateService;
}

const useStyles = makeStyles({
  container: {
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalM,
  },
  header: {
    display: 'flex',
    justifyContent: 'space-between',
    alignItems: 'center',
  },
  table: {
    width: '100%',
  },
  actions: {
    display: 'flex',
    gap: tokens.spacingHorizontalS,
  },
});

const documentTypes = [
  'Drawing',
  'Specification',
  'Report',
  'Contract',
  'Correspondence',
  'Other',
];

const DocumentsTab: React.FC<IDocumentsTabProps> = ({
  context,
  project,
  documentService,
  templateService,
}) => {
  const styles = useStyles();
  const [documents, setDocuments] = useState<IDocument[]>([]);
//   const [selectedDocument, setSelectedDocument] = useState<IDocument | null>(null);
  const [showUploadDialog, setShowUploadDialog] = useState(false);
  const [uploadFile, setUploadFile] = useState<File | null>(null);
  const [documentType, setDocumentType] = useState('');

  const loadDocuments = async (): Promise<void> => {
    if (project) {
      try {
        const docs = await documentService.getProjectDocuments(project.ProjectNumber);
        setDocuments(docs);
      } catch (error) {
        console.error('Error loading documents:', error);
      }
    }
  };

  useEffect(() => {
    if (project) {
      loadDocuments().catch(console.error);
    }
  }, [project]);

  const handleFileChange = (event: React.ChangeEvent<HTMLInputElement>): void => {
    const files = event.target.files;
    if (files && files.length > 0) {
      setUploadFile(files[0]);
    }
  };

  const handleUpload = async (): Promise<void> => {
    if (!project || !uploadFile) return;

    try {
      await documentService.uploadDocument(project.ProjectNumber, uploadFile, {
        DocumentType: documentType,
        Status: 'New',
      });
      setShowUploadDialog(false);
      setUploadFile(null);
      setDocumentType('');
      await loadDocuments();
    } catch (error) {
      console.error('Error uploading document:', error);
    }
  };

  const handleDelete = async (document: IDocument): Promise<void> => {
    try {
      await documentService.deleteDocument(document.Id);
      await loadDocuments();
    } catch (error) {
      console.error('Error deleting document:', error);
    }
  };

  return (
    <div className={styles.container}>
      <div className={styles.header}>
        <h2>Project Documents</h2>
        <Button appearance="primary" onClick={() => setShowUploadDialog(true)}>
          Upload Document
        </Button>
      </div>

      <Table className={styles.table}>
        <TableHeader>
          <TableRow>
            <TableHeaderCell>Name</TableHeaderCell>
            <TableHeaderCell>Type</TableHeaderCell>
            <TableHeaderCell>Status</TableHeaderCell>
            <TableHeaderCell>Modified</TableHeaderCell>
            <TableHeaderCell>Actions</TableHeaderCell>
          </TableRow>
        </TableHeader>
        <TableBody>
          {documents.map((doc) => (
            <TableRow key={doc.Id}>
              <TableCell>{doc.Title}</TableCell>
              <TableCell>{doc.DocumentType}</TableCell>
              <TableCell>{doc.Status}</TableCell>
              <TableCell>{new Date(doc.Modified).toLocaleDateString()}</TableCell>
              <TableCell>
                <div className={styles.actions}>
                  <Button
                    appearance="subtle"
                    onClick={() => window.open(doc.FileRef, '_blank')}
                  >
                    View
                  </Button>
                  <Button
                    appearance="subtle"
                    onClick={() => handleDelete(doc)}
                  >
                    Delete
                  </Button>
                </div>
              </TableCell>
            </TableRow>
          ))}
        </TableBody>
      </Table>

      <Dialog open={showUploadDialog}>
        <DialogSurface>
          <DialogBody>
            <DialogTitle>Upload Document</DialogTitle>
            <DialogContent>
              <div className={styles.container}>
                <Field label="Document Type">
                  <Dropdown
                    value={documentType}
                    onOptionSelect={(_, data) => setDocumentType(data.optionValue || '')}
                  >
                    {documentTypes.map((type) => (
                      <Option key={type} value={type}>
                        {type}
                      </Option>
                    ))}
                  </Dropdown>
                </Field>
                <Field label="File">
                  <input type="file" onChange={handleFileChange} style={{ width: '100%' }} />
                </Field>
              </div>
            </DialogContent>
            <DialogActions>
              <Button appearance="secondary" onClick={() => setShowUploadDialog(false)}>
                Cancel
              </Button>
              <Button appearance="primary" onClick={handleUpload} disabled={!uploadFile || !documentType}>
                Upload
              </Button>
            </DialogActions>
          </DialogBody>
        </DialogSurface>
      </Dialog>
    </div>
  );
};

export default DocumentsTab; 