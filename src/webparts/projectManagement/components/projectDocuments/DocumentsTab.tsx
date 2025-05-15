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
import { spfi } from '@pnp/sp';
import { SPFx } from '@pnp/sp/presets/all';
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { Project } from '../../services/ProjectSelectionServices';
import { FieldTypes } from '@pnp/sp/fields/types';

interface DocumentsTabProps {
  project: Project;
  context: WebPartContext;
}

interface ProjectDocumentItem {
  Id: number;
  [key: string]: FieldTypes;
}

const listName = '9719_ProjectDocuments';

const documentFields = [
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
  const sp = spfi().using(SPFx(context));
  const styles = useStyles();
  const toasterId = useId('doc-toast');
  const { dispatchToast } = useToastController();
  const [docItem, setDocItem] = useState<ProjectDocumentItem | null>(null);

  const ensureDocumentEntry = async (): Promise<void> => {
    const items = await sp.web.lists.getByTitle(listName).items
      .filter(`ProjectNumber eq '${project.ProjectNumber}'`).top(1)();
    if (!items.length) {
      const addResult = await sp.web.lists.getByTitle(listName).items.add({
        Title: project.ProjectName,
        ProjectNumber: project.ProjectNumber
      });
      setDocItem(addResult.data);
    } else {
      setDocItem(items[0]);
    }
  };

  const fetchDocumentData = async (): Promise<void> => {
    await ensureDocumentEntry();
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

  return (
    <>
      <div className={styles.container}>
        <Text size={600} weight="semibold">Project Documents</Text>

        {docItem ? (
          <>
            {documentFields.map((field) => {
              const path = docItem[field];
              return (
                <div key={field} className={styles.docRow}>
                  <Text className={styles.fieldLabel}>{field}</Text>
                  <Text className={styles.fieldValue}>
                    {path ? String(path) : 'Not available'}
                  </Text>
                  {path && (
                    <Button appearance="secondary" onClick={() => copyToClipboard(String(path), field)}>
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
