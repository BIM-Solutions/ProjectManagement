import * as React from 'react';
import { useState } from 'react';
import {
  Dialog,
  DialogSurface,
  DialogBody,
  DialogTitle,
  DialogContent,
  DialogActions,
  Button,
  makeStyles,
  tokens,
  Dropdown,
  Option,
  Text,
  Spinner,
} from '@fluentui/react-components';
import { FilePicker, IFilePickerResult } from '@pnp/spfx-controls-react';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { StandardsService, IStandard } from '../../services/StandardsService';
import { SPFI } from '@pnp/sp';

interface StandardsWorkflowDialogProps {
  isOpen: boolean;
  onDismiss: () => void;
  context: WebPartContext;
  sp: SPFI;
  projectNumber: string;
  standardsService: StandardsService;
}

const useStyles = makeStyles({
  container: {
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalM,
  },
  loadingContainer: {
    display: 'flex',
    justifyContent: 'center',
    alignItems: 'center',
    padding: tokens.spacingVerticalL,
  },
  errorText: {
    color: tokens.colorPaletteRedForeground1,
    marginTop: tokens.spacingVerticalS,
  },
});

export const StandardsWorkflowDialog: React.FC<StandardsWorkflowDialogProps> = ({
  isOpen,
  onDismiss,
  context,
  sp,
  projectNumber,
  standardsService,
}) => {
  const styles = useStyles();
  const [isLoading, setIsLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [standards, setStandards] = useState<IStandard[]>([]);
  const [selectedStandard, setSelectedStandard] = useState<IStandard | null>(null);
  const [workflowType, setWorkflowType] = useState<'new' | 'existing' | null>(null);

  const handleNewStandard = async (files: IFilePickerResult[]): Promise<void> => {
    if (!files.length) return;
    const file = files[0];
    
    try {
      setIsLoading(true);
      setError(null);
      await standardsService.uploadNewStandard(file, projectNumber);
      if (file.fileName) {
        await standardsService.copyTemplates(projectNumber, file.fileName);
      }
      onDismiss();
    } catch (error) {
      setError('Error uploading new standard. Please try again.');
      console.error('Error:', error);
    } finally {
      setIsLoading(false);
    }
  };

  const handleExistingStandard = async (): Promise<void> => {
    if (!selectedStandard) return;

    try {
      setIsLoading(true);
      setError(null);
      await standardsService.copyExistingStandard(selectedStandard, projectNumber);
      if (selectedStandard.UniclassCode) {
        await standardsService.copyTemplates(projectNumber, selectedStandard.UniclassCode);
      }
      onDismiss();
    } catch (error) {
      setError('Error copying existing standard. Please try again.');
      console.error('Error:', error);
    } finally {
      setIsLoading(false);
    }
  };

  const loadStandards = async (): Promise<void> => {
    try {
      setIsLoading(true);
      const standardsList = await standardsService.getExistingStandards();
      setStandards(standardsList);
    } catch (error) {
      setError('Error loading standards. Please try again.');
      console.error('Error:', error);
    } finally {
      setIsLoading(false);
    }
  };

  React.useEffect(() => {
    if (isOpen && workflowType === 'existing') {
      loadStandards().catch(console.error);
    }
  }, [isOpen, workflowType]);

  return (
    <Dialog open={isOpen} onOpenChange={(_, { open }) => !open && onDismiss()}>
      <DialogSurface>
        <DialogBody>
          <DialogTitle>Standards Workflow</DialogTitle>
          <DialogContent>
            <div className={styles.container}>
              {!workflowType ? (
                <>
                  <Button onClick={() => setWorkflowType('new')}>New Standards</Button>
                  <Button onClick={() => setWorkflowType('existing')}>Existing Standards</Button>
                </>
              ) : workflowType === 'new' ? (
                <FilePicker
                  context={context}
                  buttonLabel="Upload Standard"
                  onSave={handleNewStandard}
                  onCancel={() => setWorkflowType(null)}
                  hideLocalUploadTab={false}
                  hideLinkUploadTab={true}
                  hideOneDriveTab={true}
                  hideWebSearchTab={true}
                  hideStockImages={true}
                />
              ) : (
                <>
                  <Dropdown
                    placeholder="Select a standard"
                    onOptionSelect={(_, data) => {
                      const optionValue = data.optionValue as string;
                      const standard = standards.find(s => s.Id === parseInt(optionValue));
                      setSelectedStandard(standard || null);
                    }}
                  >
                    {standards.map(standard => (
                      <Option 
                        key={standard.Id} 
                        value={standard.Id.toString()}
                        text={`${standard.Title} (v${standard.Version})`}
                      />
                    ))}
                  </Dropdown>
                  <Button 
                    onClick={handleExistingStandard}
                    disabled={!selectedStandard}
                  >
                    Copy Selected Standard
                  </Button>
                </>
              )}

              {isLoading && (
                <div className={styles.loadingContainer}>
                  <Spinner />
                </div>
              )}

              {error && (
                <Text className={styles.errorText}>{error}</Text>
              )}
            </div>
          </DialogContent>
          <DialogActions>
            <Button appearance="secondary" onClick={onDismiss}>
              Cancel
            </Button>
          </DialogActions>
        </DialogBody>
      </DialogSurface>
    </Dialog>
  );
}; 