import * as React from 'react';
import { useState, useEffect } from 'react';
import { Dropdown, IDropdownOption } from '@fluentui/react/lib/Dropdown';
import { PrimaryButton } from '@fluentui/react/lib/Button';
import { TextField } from '@fluentui/react/lib/TextField';
import { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/files';
import '@pnp/sp/folders';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import { WebPartContext } from '@microsoft/sp-webpart-base';


interface IDocumentUploadProps {
    sp: SPFI;
    libraryName: string;
    onUploadComplete: () => void;
    context: WebPartContext;
    projectId: string | undefined;
}

interface IUniclassItem {
    code: string;
    title: string;
    subgroups?: IUniclassItem[];
    sections?: IUniclassItem[];
}

export const DocumentUpload: React.FC<IDocumentUploadProps> = ({ 
    sp, 
    libraryName, 
    onUploadComplete,
    context,
    projectId
}) => {
    const [uniclassData, setUniclassData] = useState<IUniclassItem[]>([]);
    const [selectedLevel1, setSelectedLevel1] = useState<string>('');
    const [selectedLevel2, setSelectedLevel2] = useState<string>('');
    const [selectedLevel3, setSelectedLevel3] = useState<string>('');
    const [file, setFile] = useState<File | null>(null);
    const [description, setDescription] = useState('');

    

    // Load Uniclass data and verify library setup
    useEffect(() => {
        const initializeComponent = async (): Promise<void> => {
            try {
                // Load Uniclass data
                const response = await fetch(`${context.pageContext.web.absoluteUrl}/SiteAssets/Uniclass2015_PM.json`);
                if (!response.ok) {
                    throw new Error(`HTTP error! status: ${response.status}`);
                }
                const data = await response.json();
                setUniclassData(data);
            } catch (error) {
                console.error('Error initializing component:', error);
            }
        };

        initializeComponent().catch(console.error);
    }, [context.pageContext.web.absoluteUrl, libraryName, sp]);

    // Generate dropdown options
    const getLevel1Options = (): IDropdownOption[] => {
        return uniclassData.map(item => ({
            key: item.code,
            text: `${item.code} - ${item.title}`
        }));
    };

    const getLevel2Options = (): IDropdownOption[] => {
        if (!selectedLevel1) return [];
        const level1Item = uniclassData.find(item => item.code === selectedLevel1);
        return (level1Item?.subgroups || []).map(item => ({
            key: item.code,
            text: `${item.code} - ${item.title}`
        }));
    };

    const getLevel3Options = (): IDropdownOption[] => {
        if (!selectedLevel1 || !selectedLevel2) return [];
        const level1Item = uniclassData.find(item => item.code === selectedLevel1);
        const level2Item = level1Item?.subgroups?.find(item => item.code === selectedLevel2);
        return (level2Item?.sections || []).map(item => ({
            key: item.code,
            text: `${item.code} - ${item.title}`
        }));
    };

    const handleFileChange = (event: React.ChangeEvent<HTMLInputElement>): void => {
        if (event.target.files && event.target.files[0]) {
            setFile(event.target.files[0]);
        }
    };

    const getFolderPath = (): string => {
        const parts: string[] = [];
        const level1Item = uniclassData.find(item => item.code === selectedLevel1);
        if (level1Item) {
            parts.push(level1Item.title);
            const level2Item = level1Item.subgroups?.find(item => item.code === selectedLevel2);
            if (level2Item) {
                parts.push(level2Item.title);
                
            }
        }
        return `${libraryName}/${parts.join('/')}`;
    };
    const getSelectedTitle = (level: 1 | 2 | 3): string => {
        if (level === 3 && selectedLevel3) {
            const level1Item = uniclassData.find(item => item.code === selectedLevel1);
            const level2Item = level1Item?.subgroups?.find(item => item.code === selectedLevel2);
            const level3Item = level2Item?.sections?.find(item => item.code === selectedLevel3);
            return level3Item?.title || '';
        } else if (level === 2 && selectedLevel2) {
            const level1Item = uniclassData.find(item => item.code === selectedLevel1);
            const level2Item = level1Item?.subgroups?.find(item => item.code === selectedLevel2);
            return level2Item?.title || '';
        } else if (level === 1 && selectedLevel1) {
            const level1Item = uniclassData.find(item => item.code === selectedLevel1);
            return level1Item?.title || '';
        }
        return '';
    };



    const handleUpload = async (): Promise<void> => {
        if (!file || !selectedLevel1) return;

        try {
            const folderPath = getFolderPath();
            console.log('Upload folderPath:', folderPath);
            
            // Create folder structure
            const docLib = await sp.web.lists.ensure(libraryName);
            console.log('docLib', docLib);
            const relativeFolderPath = folderPath.replace(`${libraryName}/`, '');
            console.log('relativeFolderPath:', relativeFolderPath);

            // Create each folder level if it doesn't exist
            const folderParts = relativeFolderPath.split('/');
            let currentPath = '';
            
            for (const part of folderParts) {
                if (currentPath === '') {
                    currentPath = part;
                } else {
                    currentPath += '/' + part;
                }
                try {
                    await sp.web.folders.addUsingPath(`${libraryName}/${projectId}/${currentPath}`);
                    console.log('Created/verified folder:', currentPath);
                } catch (error) {
                    console.log('Folder might exist:', currentPath, error);
                }
            }

            // Upload file
            const fileContent = await file.arrayBuffer();
            const result = await sp.web.getFolderByServerRelativePath(`${libraryName}/${projectId}/${relativeFolderPath}`)
                .files.addUsingPath(file.name, fileContent);
            console.log('File uploaded to:', result.ServerRelativeUrl);

            // Update metadata
            const item = await sp.web.getFileByServerRelativePath(result.ServerRelativeUrl).getItem();
            await item.update({
                ProjectId: projectId,
                Title: description,
                UniclassCode1: selectedLevel1,
                UniclassTitle1: getSelectedTitle(1),
                UniclassCode2: selectedLevel2,
                UniclassTitle2: getSelectedTitle(2),
                UniclassCode3: selectedLevel3,
                UniclassTitle3: getSelectedTitle(3)
            });
            console.log('metadata updated');

            // Clear form
            setFile(null);
            setSelectedLevel1('');
            setSelectedLevel2('');
            setSelectedLevel3('');
            setDescription('');
            onUploadComplete();

        } catch (error) {
            console.error('Error uploading file:', error);
        }
    };



    return (
        <div className="document-upload">
            <input
                type="file"
                onChange={handleFileChange}
                style={{ marginBottom: '1rem' }}
            />
            
            <Dropdown
                label="Level 1 Classification"
                options={getLevel1Options()}
                selectedKey={selectedLevel1}
                onChange={(_, option) => {
                    setSelectedLevel1(option?.key as string);
                    setSelectedLevel2('');
                    setSelectedLevel3('');
                }}
            />

            {selectedLevel1 && (
                <Dropdown
                    label="Level 2 Classification"
                    options={getLevel2Options()}
                    selectedKey={selectedLevel2}
                    onChange={(_, option) => {
                        setSelectedLevel2(option?.key as string);
                        setSelectedLevel3('');
                    }}
                />
            )}

            {selectedLevel2 && (
                <Dropdown
                    label="Level 3 Classification"
                    options={getLevel3Options()}
                    selectedKey={selectedLevel3}
                    onChange={(_, option) => setSelectedLevel3(option?.key as string)}
                />
            )}

            <TextField
                label="Description"
                multiline
                rows={3}
                value={description}
                onChange={(_, newValue) => setDescription(newValue || '')}
            />

            <PrimaryButton
                text="Upload Document"
                onClick={handleUpload}
                disabled={!file || !selectedLevel1}
                style={{ marginTop: '1rem' }}
            />
        </div>
    );
}; 