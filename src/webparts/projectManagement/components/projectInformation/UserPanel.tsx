import * as React from 'react';
import { 
    Text, 
    Persona, 
    Card,
} from '@fluentui/react-components';
import { IUserField } from '../../services/ProjectSelectionServices';

import { WebPartContext } from '@microsoft/sp-webpart-base';

interface UserPanelProps {
    title: string;
    user: IUserField;
    context: WebPartContext;
    onClick?: (user: IPersonProperties) => void;
}

export interface IPersonProperties {
    displayName: string;
    jobTitle: string;
    department: string;
    presence: string;
    email: string;
    mail: string;
    mobilePhone: string;
    officeLocation: string;
    title: string;
  }


 const getUserPresence = async (email: string, context: WebPartContext): Promise<{ presence: string, user: IPersonProperties }> => {
    const client = await context.msGraphClientFactory.getClient("3");
    const user = await client.api(`/users/${email}`).get();
    const presence = await client.api(`/users/${user.id}/presence`).get();
    return { presence, user };
 }

const UserPanel: React.FC<UserPanelProps> = ({ title, user, context, onClick }) => {
    const [presence, setPresence] = React.useState<string | null>(null);
    const [extendedUser, setExtendedUser] = React.useState<IPersonProperties | null>(null);

    React.useEffect(() => {
        const fetchPresence = async (): Promise<void> => {
            try {
                if (!user?.Email) {
                    return;
                }
                const userPresence = await getUserPresence(user.Email, context);
                setPresence(userPresence.presence);
                setExtendedUser(userPresence.user);
            } catch {
                // Do nothing
            }
        };
        fetchPresence().catch((error) => {
            console.error("Error occurred while fetching presence:", error);
        });
    }, [user.Email, context]);

    return (
        <Card
            style={{
                margin: '12px 0',
                boxShadow: '0 2px 8px rgba(0,0,0,0.08)',
                borderRadius: 12,
                cursor: onClick && extendedUser ? 'pointer' : 'default',
                transition: 'box-shadow 0.2s',
                padding: 16,
                minWidth: 260,
                maxWidth: 340,
                background: '#fff',
            }}
            onClick={() => onClick && extendedUser && onClick(extendedUser)}
        >
            <Text size={400} weight='bold'>{title}</Text>
            <Persona
                    size="extra-large"
                    name={user.Title}
                    presence={presence}
                    secondaryText={user.JobTitle}
                    tertiaryText={user.Department}
                    avatar={{ image: {
                        src:`/_layouts/15/userphoto.aspx?accountname=${encodeURIComponent(user.Email)}`,
                        alt: user.Title,
                    }}}
                />
        </Card>
    );
};

export default UserPanel;