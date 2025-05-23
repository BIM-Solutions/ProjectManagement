import * as React from 'react';
import {
  Dialog, DialogSurface, DialogBody, Button, Text, Link,
  Persona,
  makeStyles
} from '@fluentui/react-components';
// import { Stack, FontIcon } from '@fluentui/react';
import { IPersonProperties } from './UserPanel';
import { CallRegular, MailRegular, ChatRegular } from '@fluentui/react-icons';
const pink = '#E3008C';

interface UserDetailsModalProps {
  open: boolean;
  user: IPersonProperties; // Graph API user object
  manager?: IPersonProperties; // Optionally pass manager object if available
  onClose: () => void;
}

const useStyles = makeStyles({
    innerWrapper: {
      alignItems: "start",
      columnGap: "15px",
      display: "flex",
    },
    outerWrapper: {
      display: "flex",
      flexDirection: "column",
      rowGap: "15px",
      minWidth: "min-content",
    },
  });

// const stackFieldTokens = { childrenGap: 15 };

const getUserPhotoUrl = (email?: string): string | undefined =>
  email ? `/_layouts/15/userphoto.aspx?accountname=${encodeURIComponent(email)}` : undefined;

const UserDetailsModal: React.FC<UserDetailsModalProps> = ({ open, user, manager, onClose }) => {
  if (!user) return null;
  const styles = useStyles();

  return (
    <Dialog open={open} onOpenChange={onClose}>
      <DialogSurface>
        <DialogBody>
          <div className={styles.outerWrapper}>

            {/* Profile Section */}
            <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', padding: 24, borderBottom: '1px solid #f3f3f3' }}>
            <Persona
                    size="huge"
                    name={user.displayName}
                    secondaryText={user.jobTitle}
                    tertiaryText={user.department}
                    avatar={{ image: {
                        src:getUserPhotoUrl(user.email || user.mail),
                        alt: user.title,
                    }}}
                />
            </div>

            {/* Tabs (Contact/Phone) */}
            <div className={styles.innerWrapper}>
                <Button icon={<CallRegular />} appearance="transparent" />
                <Button icon={<MailRegular />} appearance="transparent" />
                <Button icon={<ChatRegular />} appearance="transparent" />
            </div>

            {/* Contact Section */}
            <div style={{ padding: 24, paddingBottom: 8 }}>
              {user.mail && (
                <div style={{ display: 'flex', alignItems: 'center', gap: 12, marginBottom: 12 }}>
                  <span style={{ color: pink, fontSize: 18 }}>✉️</span>
                  <Link href={`mailto:${user.mail}`} target="_blank" style={{ color: pink, fontWeight: 500 }}>
                    <Text size={300}>{user.mail}</Text>
                  </Link>
                </div>
              )}
              {user.mobilePhone && (
                <div style={{ display: 'flex', alignItems: 'center', gap: 12, marginBottom: 12 }}>
                  <span style={{ color: pink, fontSize: 18 }}>📱</span>
                  <Link href={`tel:${user.mobilePhone}`} target="_blank" style={{ color: pink, fontWeight: 500 }}>
                    <Text size={300}>{user.mobilePhone}</Text>
                  </Link>
                </div>
              )}
              {user.officeLocation && (
                <div style={{ display: 'flex', alignItems: 'center', gap: 12, marginBottom: 12 }}>
                  <span style={{ color: pink, fontSize: 18 }}>📍</span>
                  <Link
                    href={`https://www.bing.com/maps?q=${encodeURIComponent(user.officeLocation)}`}
                    target="_blank"
                    style={{ color: pink, fontWeight: 500 }}
                  >
                    <Text size={300}>{user.officeLocation}</Text>
                  </Link>
                </div>
              )}
            </div>

            {/* Reports To Section */}
            {manager && (
              <div style={{ background: '#faf9f8', padding: 20, borderTop: '1px solid #f3f3f3', marginTop: 8 }}>
                <Text size={400} weight="semibold" style={{ color: '#888', marginBottom: 10, display: 'block' }}>
                  Reports to
                </Text>
                <div style={{ display: 'flex', alignItems: 'center', gap: 14 }}>
                  <img
                    src={getUserPhotoUrl(manager.email || manager.mail)}
                    alt={manager.displayName}
                    style={{ width: 40, height: 40, borderRadius: '50%', objectFit: 'cover', border: `2px solid ${pink}` }}
                  />
                  <div>
                    <Text size={300} weight="semibold">{manager.displayName}</Text>
                    <Text size={200} style={{ color: '#888', display: 'block' }}>{manager.title || manager.jobTitle}</Text>
                  </div>
                </div>
              </div>
            )}

            {/* Close Button */}
            <div style={{ display: 'flex', justifyContent: 'center', margin: 20 }}>
              <Button onClick={onClose} appearance="secondary">Close</Button>
            </div>
          </div>
        </DialogBody>
      </DialogSurface>
    </Dialog>
  );
};

export default UserDetailsModal; 