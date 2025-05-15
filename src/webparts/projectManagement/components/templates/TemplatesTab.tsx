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
  DialogSurface,
  DialogBody,
  DialogTitle,
  DialogContent,
  DialogActions,
  Input,
  Field,
  Textarea,
  Dropdown,
  Option,
} from '@fluentui/react-components';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { IProject } from '../../services/ProjectService';
import { TemplateService, ITemplate } from '../../services/TemplateService';

export interface ITemplatesTabProps {
  context: WebPartContext;
  project: IProject | undefined;
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
  form: {
    display: 'grid',
    gridTemplateColumns: '1fr 1fr',
    gap: tokens.spacingHorizontalM,
  },
  fullWidth: {
    gridColumn: '1 / -1',
  },
});

const templateCategories = [
  'Project Setup',
  'Quality Control',
  'Documentation',
  'Reports',
  'Standards',
];

const TemplatesTab: React.FC<ITemplatesTabProps> = ({
  context,
  project,
  templateService,
}) => {
  const styles = useStyles();
  const [templates, setTemplates] = useState<ITemplate[]>([]);
  const [showTemplateDialog, setShowTemplateDialog] = useState(false);
//   const [selectedTemplate, setSelectedTemplate] = useState<ITemplate | null>(null);
  const [formData, setFormData] = useState<Partial<ITemplate>>({});

  const loadTemplates = async (): Promise<void> => {
    try {
      const loadedTemplates = await templateService.getTemplates();
      setTemplates(loadedTemplates);
    } catch (error) {
      console.error('Error loading templates:', error);
    }
  };

  useEffect(() => {
    loadTemplates().catch(console.error);
  }, []);

  const handleInputChange = (field: keyof ITemplate, value: string | boolean): void => {
    setFormData(prev => ({
      ...prev,
      [field]: value,
    }));
  };

  const handleSave = async (): Promise<void> => {
    try {
      const newTemplate = await templateService.createTemplate({
        ...formData,
        CreatedDate: new Date().toISOString(),
        ModifiedDate: new Date().toISOString(),
      });
      setTemplates([...templates, newTemplate]);
      setShowTemplateDialog(false);
      setFormData({});
    } catch (error) {
      console.error('Error saving template:', error);
    }
  };

  const handleApplyTemplate = async (template: ITemplate): Promise<void> => {
    if (!project) return;

    try {
      await templateService.copyTemplateToProject(template.Id, project.ProjectNumber);
      // You might want to refresh the documents list or show a success message
    } catch (error) {
      console.error('Error applying template:', error);
    }
  };

  return (
    <div className={styles.container}>
      <div className={styles.header}>
        <h2>Templates & Standards</h2>
        <Button appearance="primary" onClick={() => setShowTemplateDialog(true)}>
          New Template
        </Button>
      </div>

      <Table className={styles.table}>
        <TableHeader>
          <TableRow>
            <TableHeaderCell>Title</TableHeaderCell>
            <TableHeaderCell>Category</TableHeaderCell>
            <TableHeaderCell>Version</TableHeaderCell>
            <TableHeaderCell>Status</TableHeaderCell>
            <TableHeaderCell>Modified</TableHeaderCell>
            <TableHeaderCell>Actions</TableHeaderCell>
          </TableRow>
        </TableHeader>
        <TableBody>
          {templates.map((template) => (
            <TableRow key={template.Id}>
              <TableCell>{template.Title}</TableCell>
              <TableCell>{template.Category}</TableCell>
              <TableCell>{template.Version}</TableCell>
              <TableCell>{template.IsActive ? 'Active' : 'Inactive'}</TableCell>
              <TableCell>
                {new Date(template.ModifiedDate).toLocaleDateString()}
              </TableCell>
              <TableCell>
                <Button
                  appearance="subtle"
                  onClick={() => handleApplyTemplate(template)}
                  disabled={!project}
                >
                  Apply
                </Button>
              </TableCell>
            </TableRow>
          ))}
        </TableBody>
      </Table>

      <Dialog open={showTemplateDialog}>
        <DialogSurface>
          <DialogBody>
            <DialogTitle>New Template</DialogTitle>
            <DialogContent>
              <div className={styles.form}>
                <Field label="Title" required>
                  <Input
                    value={formData.Title || ''}
                    onChange={(_, data) => handleInputChange('Title', data.value)}
                  />
                </Field>

                <Field label="Category">
                  <Dropdown
                    value={formData.Category || ''}
                    onOptionSelect={(_, data) =>
                      handleInputChange('Category', data.optionValue || '')
                    }
                  >
                    {templateCategories.map((category) => (
                      <Option key={category} value={category}>
                        {category}
                      </Option>
                    ))}
                  </Dropdown>
                </Field>

                <Field label="Version" required>
                  <Input
                    value={formData.Version || ''}
                    onChange={(_, data) => handleInputChange('Version', data.value)}
                  />
                </Field>

                <Field label="Status">
                  <Dropdown
                    value={formData.IsActive ? 'Active' : 'Inactive'}
                    onOptionSelect={(_, data) =>
                      handleInputChange('IsActive', data.optionValue === 'Active')
                    }
                  >
                    <Option value="Active">Active</Option>
                    <Option value="Inactive">Inactive</Option>
                  </Dropdown>
                </Field>

                <Field label="Description" className={styles.fullWidth}>
                  <Textarea
                    value={formData.Description || ''}
                    onChange={(_, data) =>
                      handleInputChange('Description', data.value)
                    }
                  />
                </Field>
              </div>
            </DialogContent>
            <DialogActions>
              <Button
                appearance="secondary"
                onClick={() => setShowTemplateDialog(false)}
              >
                Cancel
              </Button>
              <Button appearance="primary" onClick={handleSave}>
                Save
              </Button>
            </DialogActions>
          </DialogBody>
        </DialogSurface>
      </Dialog>
    </div>
  );
};

export default TemplatesTab; 