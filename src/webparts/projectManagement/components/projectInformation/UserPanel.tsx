import * as React from 'react';
import { Text, Divider, Persona, makeStyles} from '@fluentui/react-components';
import { IUserField } from '../../services/ProjectSelectionServices';
import { WebPartContext } from '@microsoft/sp-webpart-base';
interface UserPanelProps {
    title: string;
    user: IUserField;
    context: WebPartContext
}

const useStyles = makeStyles({
    userPanel: {
        gap: '20',
        width: '100%'
    }
})
 const getUserPresence = async (email: string, context: WebPartContext): Promise<string> => {
    const client = await context.msGraphClientFactory.getClient("3");
    const presence = await client.api(`/users/${email}/presence`).get();
    return presence.availability;
 }

const UserPanel: React.FC<UserPanelProps> = ({ title, user, context }) => {
    const styles = useStyles();
    const [presence, setPresence] = React.useState<string | null>(null);

    React.useEffect(() => {
        const fetchPresence = async (): Promise<void> => {
            try {
                if (!user?.Email) {
                    // console.warn("User email is missing, cannot fetch presence.");
                    return;
                }
    
                // console.log("Requesting presence for:", user.Email);
                const userPresence = await getUserPresence(user.Email, context);
                // console.log("Fetched presence:", userPresence);
                setPresence(userPresence);
            } catch {
                // Do nothing
            }
        };
    
        fetchPresence().catch((error) => {
            console.error("Error occurred while fetching presence:", error);
        });
    }, [user.Email, context]);
    
    // console.log("the presence is2:", presence);
    return (
        <div className={styles.userPanel}>
        <Divider />
        {user && (
            <>
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

            </>
        )}
        </div>
    );
};

export default UserPanel;