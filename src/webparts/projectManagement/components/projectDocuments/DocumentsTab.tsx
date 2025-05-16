import * as React from 'react';
import { useEffect, useState } from 'react';
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
} from '@fluentui/react-components';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { DocumentService, IDocument } from '../../services/DocumentService';
import { Project } from '../../services/ProjectSelectionServices';

interface DocumentsTabProps {
  project: Project;
  context: WebPartContext;
}

const documentTypes = [
  "EIR", "AIR", "PIR", "SMP", "EIRAppraisal", "PreContractBEP", "BEP", "MPDT", "AIDP", "LAP_EIR", "DRM", "MIDP",
  "TIDPs", "ResponsibilityAssignmentMatrix", "IMRiskRegister", "MobilisationPlan", "FederatedModel", "QAR",
  "DataReports", "HelthCheck_WarningReport",
  "ProjectExecutionPlan", "ProjectManagementPlan", "ProjectQualityPlan", "ProjectControlPlan",
  "ProjectRiskPlan", "ProjectChangePlan", "ProjectCostPlan", "ProjectSchedulePlan", "ProjectResourcePlan",
  "ProjectCommunicationsPlan", "ProjectProcurementPlan", "ProjectStakeholderPlan"
];

const useStyles = makeStyles({
  container: { display: 'flex', flexDirection: 'column', gap: tokens.spacingVerticalL, paddingTop: tokens.spacingVerticalL },
  docRow: { display: 'flex', flexDirection: 'row', alignItems: 'center', gap: tokens.spacingHorizontalM },
  fieldLabel: { width: '250px', fontWeight: tokens.fontWeightSemibold },
  fieldValue: { flexGrow: 1, fontStyle: 'italic' },
});

const DocumentsTab: React.FC<DocumentsTabProps> = ({ context, project }) => {
  const styles = useStyles();
  const toasterId = useId('doc-toast');
  const { dispatchToast } = useToastController();
  const [documents, setDocuments] = useState<IDocument[]>([]);
  const documentService = new DocumentService(context);

  const fetchDocumentData = async (): Promise<void> => {
    try {
      const docs = await documentService.getProjectDocuments(project.ProjectNumber);
      setDocuments(docs);
    } catch (error) {
      console.error('Error fetching documents:', error);
      dispatchToast(
        <Toast>
          <ToastTitle>Error</ToastTitle>
          <ToastBody>Failed to load project documents.</ToastBody>
        </Toast>,
        { intent: 'error', timeout: 3000 }
      );
    }
  };

  const copyToClipboard = async (value: string, label: string): Promise<void> => {
    try {
      await navigator.clipboard.writeText(value);
      dispatchToast(
        <Toast>
          <ToastTitle>Copied</ToastTitle>
          <ToastBody>Path for <strong>{label}</strong> copied to clipboard.</ToastBody>
        </Toast>,
        { intent: 'success', timeout: 3000 }
      );
    } catch (err) {
      console.error('Failed to copy path:', err);
    }
  };

  useEffect(() => {
    fetchDocumentData().catch(console.error);
  }, [project]);

  const getDocumentForType = (docType: string): IDocument | undefined => {
    return documents.find(doc => doc.DocumentType === docType);
  };

  return (
    <>
      <div className={styles.container}>
        <Text size={600} weight="semibold">Project Documents</Text>

        {documents ? (
          <>
            {documentTypes.map((docType) => {
              const doc = getDocumentForType(docType);
              return (
                <div key={docType} className={styles.docRow}>
                  <Text className={styles.fieldLabel}>{docType}</Text>
                  <Text className={styles.fieldValue}>
                    {doc ? doc.FileRef : 'Not available'}
                  </Text>
                  {doc && (
                    <Button appearance="secondary" onClick={() => copyToClipboard(doc.FileRef, docType)}>
                      Copy Path
                    </Button>
                  )}
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
