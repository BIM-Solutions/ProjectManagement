import * as React from 'react';

// Define interfaces for the folder structure
interface UniclassFolder {
    code: string;
    name: string;
    description: string;
    children?: UniclassFolder[];
}

// Define the root level folders based on Uniclass PM codes
export const ProjectFolders: UniclassFolder[] = [
    {
        code: 'PM_10',
        name: 'Project information',
        description: 'Project Information',
        children: [
            {
                code: 'PM_10_20',
                name: 'Client requirements',
                description: 'Client requirements',
                children: [
                    {
                        code: 'PM_10_20_28',
                        name: 'Exchange information requirements',
                        description: 'Exchange information requirements'
                    }
                ]
            },
            {
                code: 'PM_10_21',
                name: 'Client requirements',
                description: 'Client requirements'
            },
            {
                code: 'PM_10_22',
                name: 'Client requirements',
                description: 'Client requirements'
            }
        ]
    },
    {
        code: 'PM_40',
        name: 'Design and approvals information',
        description: 'Design and approvals information',
        children: [
            {
                code: 'PM_40_60',
                name: 'Project information management information',
                description: 'Project information management information',
                children: [
                    {
                        code: 'PM_40_60_64',
                        name: 'Pre-appointment BIM execution plan',
                        description: 'Pre-appointment BIM execution plan'
                    },
                    {
                        code: 'PM_40_60_66',
                        name: 'Project information production methods and procedures',
                        description: 'Project information production methods and procedures'
                    },
                    {
                        code: 'PM_40_60_68',
                        name: 'Project information standards',
                        description: 'Project information standards'
                    }
                ]
            }
        ]
    },
    {
        code: 'PM_80',
        name: 'Asset management information',
        description: 'Asset management information',
        children: [
            {
                code: 'PM_80_10',
                name: 'Asset strategy, planning and implementation',
                description: 'Asset strategy, planning and implementation',
                children: [
                    {
                        code: 'PM_80_10_03',
                        name: 'Asset information protocol',
                        description: 'Asset information protocol'
                    }
                ]
            }
        ]
    }
];

// Recursive component to render the folder structure
const FolderNode: React.FC<{ folder: UniclassFolder, level?: number }> = ({ folder, level = 0 }) => {
    return (
        <div style={{ marginLeft: `${level * 20}px` }}>
            <div className="folder-item">
                <span className="folder-code">{folder.code}</span>
                <span className="folder-name">{folder.name}</span>
            </div>
            {folder.children && folder.children.length > 0 && (
                <div className="folder-children">
                    {folder.children.map((child, index) => (
                        <FolderNode key={child.code} folder={child} level={level + 1} />
                    ))}
                </div>
            )}
        </div>
    );
};

// Main component to render the complete folder structure
export const FolderStructure: React.FC = () => {
    return (
        <div className="folder-structure">
            {ProjectFolders.map((folder) => (
                <FolderNode key={folder.code} folder={folder} />
            ))}
        </div>
    );
};

