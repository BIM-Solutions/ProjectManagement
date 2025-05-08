import * as React from 'react';
import { makeStyles, Text} from '@fluentui/react-components';
import { Project } from '../../services/ProjectSelectionServices';
import UserPanel from './UserPanel';
import { WebPartContext } from '@microsoft/sp-webpart-base';
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
  return(
  <div className={styles.container}>
    <Text size={600}>Project Team</Text>
    {project.PM && (
      <>
      <UserPanel title="Project Manager" user={project.PM} context={context} />
      </>
    )}
    {project.Manager && (
      <>
      <UserPanel title="Information Manager" user={project.Manager} context={context} />
      </>
    )}
    {project.Checker && (
      <>
      <UserPanel title="Project Checker" user={project.Checker} context={context} />
      </>
    )}
    {project.Approver && (
      <>
      <UserPanel title="Project Approver" user={project.Approver} context={context} />
      </>
    )} 
  </div>
  );
};

export default ProjectTeam;