// DocumentsTab.tsx
import * as React from 'react';
import { useEffect, useState } from 'react';
import { Stack, Text, DefaultButton, MessageBar, MessageBarType } from '@fluentui/react';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { spfi } from '@pnp/sp';
import { SPFx } from '@pnp/sp/presets/all';
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { Project } from '../../../common/services/ProjectSelectionServices';
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

const DocumentsTab: React.FC<DocumentsTabProps> = ({ context, project }) => {
  const sp = spfi().using(SPFx(context));
  const [docItem, setDocItem] = useState<ProjectDocumentItem | null>(null);
  const [copiedField, setCopiedField] = useState<string | null>(null);

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

  const copyToClipboard: (value: string, label: string) => Promise<void> = async (value: string, label: string) => {
    try {
      await navigator.clipboard.writeText(value);
      setCopiedField(label);
      setTimeout(() => setCopiedField(null), 3000);
    } catch (err) {
      console.error('Failed to copy path:', err);
    }
  };

  useEffect(() => {
    fetchDocumentData().catch(console.error);
  }, [project]);

  return (
    <Stack tokens={{ childrenGap: 20 }} styles={{ root: { paddingTop: 20 } }}>
      <Text variant="xLarge">Project Documents</Text>
      {docItem ? (
        <Stack tokens={{ childrenGap: 10 }}>
          {documentFields.map((field) => {
            const path = docItem[field];
            return (
              <Stack horizontal key={field} tokens={{ childrenGap: 10 }} verticalAlign="center">
                <Text styles={{ root: { width: 250, fontWeight: 600 } }}>{field}</Text>
                <Text styles={{ root: { flexGrow: 1 } }}>
                  {path ? path : <em>Not available</em>}
                </Text>
                {path && (
                  <DefaultButton
                    text="Copy Path"
                    onClick={() => copyToClipboard(String(path), field)}
                  />
                )}
              </Stack>
            );
          })}
        </Stack>
      ) : (
        <Text>Loading documents...</Text>
      )}
      {copiedField && (
        <MessageBar messageBarType={MessageBarType.success}>
          Path for <strong>{copiedField}</strong> copied to clipboard!
        </MessageBar>
      )}
    </Stack>
  );
};

export default DocumentsTab;
