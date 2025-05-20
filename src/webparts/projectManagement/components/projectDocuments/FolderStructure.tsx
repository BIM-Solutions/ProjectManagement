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
                name: '',
                description: '',
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
                name: '',
                description: ''
            },
            {
                code: 'PM_10_22',
                name: '',
                description: ''
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

interface MappedDocument {
    name: string;
    code1: string;
    description1: string;
    code2: string;
    description2: string;
    code3: string;
    description3: string;
}

export const mappedDocuments: MappedDocument[] = [

    {
        name: 'BIM execution plan',
        code1: 'PM_40',
        description1: 'Design and approvals information',
        code2: 'PM_40_60',
        description2: 'Project information management information',
        code3: 'PM_40_60_08',
        description3: 'BIM execution plan'
    },
    {
        name: 'Capability and capacity assessment information',
        code1: 'PM_40',
        description1: 'Design and approvals information',
        code2: 'PM_40_60',
        description2: 'Project information management information',
        code3: 'PM_40_60_11',
        description3: 'Capability and capacity assessment information'
    },
    {
        name: 'Clash detection resolution report',
        code1: 'PM_40',
        description1: 'Design and approvals information',
        code2: 'PM_40_60',
        description2: 'Project information management information',
        code3: 'PM_40_60_12',
        description3: 'Clash detection resolution report'
    },
    {
        name: 'COBie data set',
        code1: 'PM_40',
        description1: 'Design and approvals information',
        code2: 'PM_40_60',
        description2: 'Project information management information',     
        code3: 'PM_40_60_16',
        description3: 'COBie data set'
    },
    {
        name: 'Detailed responsibility matrix',
        code1: 'PM_40',
        description1: 'Design and approvals information',
        code2: 'PM_40_60',
        description2: 'Project information management information',
        code3: 'PM_40_60_24',
        description3: 'Detailed responsibility matrix'
    },
    {
        name: 'Design programme',
        code1: 'PM_40',
        description1: 'Design and approvals information',
        code2: 'PM_40_60',
        description2: 'Project information management information',
        code3: 'PM_40_60_26',
        description3: 'Design programme'
    },
    {
        name: 'Federated coordination models',
        code1: 'PM_40',
        description1: 'Design and approvals information',
        code2: 'PM_40_60',
        description2: 'Project information management information',
        code3: 'PM_40_60_31',
        description3: 'Federated coordination models'
    },
    {
        name: 'Federated information models',
        code1: 'PM_40',
        description1: 'Design and approvals information',
        code2: 'PM_40_60',
        description2: 'Project information management information',
        code3: 'PM_40_60_33',
        description3: 'Federated information models'
    },
    {
        name: 'Information assurance',
        code1: 'PM_40',
        description1: 'Design and approvals information',
        code2: 'PM_40_60',
        description2: 'Project information management information',
        code3: 'PM_40_60_38',
        description3: 'Information assurance'
    },
    {
        name: 'Information validation',
        code1: 'PM_40',
        description1: 'Design and approvals information',
        code2: 'PM_40_60',
        description2: 'Project information management information',
        code3: 'PM_40_60_39',
        description3: 'Information validation'
    },
    {
        name: 'Information verification',
        code1: 'PM_40',
        description1: 'Design and approvals information',
        code2: 'PM_40_60',
        description2: 'Project information management information',
        code3: 'PM_40_60_40',
        description3: 'Information verification'
    },
    {
        name: 'Information delivery plan',
        code1: 'PM_40',
        description1: 'Design and approvals information',
        code2: 'PM_40_60',
        description2: 'Project information management information',
        code3: 'PM_40_60_41',
        description3: 'Information delivery plan'
    },
    {
        name: 'Key performance indicator (KPI) information',
        code1: 'PM_40',
        description1: 'Design and approvals information',
        code2: 'PM_40_60',
        description2: 'Project information management information',
        code3: 'PM_40_60_45',
        description3: 'Key performance indicator (KPI) information'
    },
    {
        name: 'Master information delivery plan',
        code1: 'PM_40',
        description1: 'Design and approvals information',
        code2: 'PM_40_60',
        description2: 'Project information management information',
        code3: 'PM_40_60_50',
        description3: 'Master information delivery plan'
    },
    {
        name: 'Mobilization plan',
        code1: 'PM_40',
        description1: 'Design and approvals information',
        code2: 'PM_40_60',
        description2: 'Project information management information',
        code3: 'PM_40_60_53',
        description3: 'Mobilization plan'
    },
    {
        name: 'Pre-appointment BIM execution plan',
        code1: 'PM_40',
        description1: 'Design and approvals information',
        code2: 'PM_40_60',
        description2: 'Project information management information',
        code3: 'PM_40_60_64',
        description3: 'Pre-appointment BIM execution plan'
    },
    {
        name: 'Project execution plan',
        code1: 'PM_40',
        description1: 'Design and approvals information',
        code2: 'PM_40_60',
        description2: 'Project information management information',
        code3: 'PM_40_60_65',
        description3: 'Project execution plan'
    },
    {
        name: 'Project information production methods and procedures',
        code1: 'PM_40',
        description1: 'Design and approvals information',
        code2: 'PM_40_60',
        description2: 'Project information management information',
        code3: 'PM_40_60_66',
        description3: 'Project information production methods and procedures'
    },
    {
        name: 'Project information protocol',
        code1: 'PM_40',
        description1: 'Design and approvals information',
        code2: 'PM_40_60',
        description2: 'Project information management information',
        code3: 'PM_40_60_67',
        description3: 'Project information protocol'
    },
    {
        name: 'Project information standards',
        code1: 'PM_40',
        description1: 'Design and approvals information',
        code2: 'PM_40_60',
        description2: 'Project information management information',
        code3: 'PM_40_60_68',
        description3: 'Project information standards'
    },
    {
        name: 'Risk schedule',
        code1: 'PM_40',
        description1: 'Design and approvals information',
        code2: 'PM_40_60',
        description2: 'Project information management information',
        code3: 'PM_40_60_70',
        description3: 'Risk schedule'
    },
    {
        name: 'Risk and opportunities management plan',
        code1: 'PM_40',
        description1: 'Design and approvals information',
        code2: 'PM_40_60',
        description2: 'Project information management information',
        code3: 'PM_40_60_72',
        description3: 'Risk and opportunities management plan'
    },
    {
        name: 'Common data environment information',
        code1: 'PM_10',
        description1: 'Project information',
        code2: 'PM_10_20',
        description2: 'Client requirements',
        code3: 'PM_10_20_15',
        description3: 'Common data environment information'
    },
    {
        name: 'Exchange information requirements',
        code1: 'PM_10',
        description1: 'Project information',
        code2: 'PM_10_20',
        description2: 'Client requirements',
        code3: 'PM_10_20_28',
        description3: 'Exchange information requirements'
    }
    

]

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

