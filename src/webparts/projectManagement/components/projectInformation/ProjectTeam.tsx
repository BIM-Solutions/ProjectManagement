import * as React from 'react';
import { makeStyles, Text} from '@fluentui/react-components';
import { Project } from '../../services/ProjectSelectionServices';
import UserPanel from './UserPanel';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import UserDetailsModal from './UserDetailsModal';
// import { IUserField } from '../../services/ProjectSelectionServices';
import { IPersonProperties } from './UserPanel';

interface ProjectTeamProps {
  project: Project;
  context: WebPartContext;
}

const useStyles = makeStyles({
  container: {
    gap: '10',
    width: '100%'
  }

})
const ProjectTeam: React.FC<ProjectTeamProps> = ({ project, context }) => {
  const styles =useStyles();
  const [selectedUser, setSelectedUser] = React.useState<IPersonProperties | null>(null);
  const [modalOpen, setModalOpen] = React.useState(false);

  const handleUserClick = (user: IPersonProperties): void => {
    setSelectedUser(user);
    setModalOpen(true);
  };

  const handleModalClose = (): void => {
    setModalOpen(false);
    setSelectedUser(null);
  };

  return(
  <div className={styles.container}>
    <Text size={600}>Project Team</Text>
    {project.PM && (
      <UserPanel title="Project Manager" user={project.PM} context={context} onClick={handleUserClick} />
    )}
    {project.Manager && (
      <UserPanel title="Information Manager" user={project.Manager} context={context} onClick={handleUserClick} />
    )}
    {project.Checker && (
      <UserPanel title="Project Checker" user={project.Checker} context={context} onClick={handleUserClick} />
    )}
    {project.Approver && (
      <UserPanel title="Project Approver" user={project.Approver} context={context} onClick={handleUserClick} />
    )}
    {selectedUser && (
      <UserDetailsModal open={modalOpen} user={selectedUser} onClose={handleModalClose} />
    )}
  </div>
  );
};

export default ProjectTeam;