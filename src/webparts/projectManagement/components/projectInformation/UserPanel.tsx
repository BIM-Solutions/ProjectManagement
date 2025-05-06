import * as React from 'react';
import { Text, Stack, Separator, Persona, PersonaSize } from '@fluentui/react';
import { IUserField } from '../../services/ProjectSelectionServices';
interface UserPanelProps {
    title: string;
    user: IUserField;
}

const UserPanel: React.FC<UserPanelProps> = ({ title, user }) => (
    <Stack tokens={{ childrenGap: 20 }} styles={{ root: { width: "100%" } }}>
        <Separator />
        {user && (
            <>
                <Text variant="mediumPlus"><strong>{title}</strong></Text>
                <Persona
                    text={user.Title}
                    secondaryText={user.JobTitle}
                    tertiaryText={user.Department}
                    imageUrl={`/_layouts/15/userphoto.aspx?accountname=${encodeURIComponent(user.Email)}`}
                    size={PersonaSize.size48}
                />

            </>
        )}
    </Stack>
);

export default UserPanel;