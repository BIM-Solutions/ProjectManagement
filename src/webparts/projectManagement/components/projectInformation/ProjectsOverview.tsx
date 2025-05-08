import * as React from 'react';
import {
  Divider,
  makeStyles,
  Body1,
  Body1Strong
} from '@fluentui/react-components';
import { Image } from '@fluentui/react/lib/Image';
import { Project } from '../../services/ProjectSelectionServices';

interface ProjectOverviewProps {
  project: Project;
}

const useStyles = makeStyles({
  container: {
    display: 'flex',
    flexDirection: 'column',
    rowGap: '16px',
    width: '100%',
  },
  imageContainer: {
    display: 'flex',
    justifyContent: 'center',
    marginBottom: '16px',
  },
  grid: {
    display: 'grid',
    gridTemplateColumns: '1fr 1fr',
    columnGap: '24px',
    rowGap: '12px',
  },
  fullRow: {
    gridColumn: '1 / -1',
  },
  fieldBlock: {
    display: 'flex',
    flexDirection: 'column',
  }
});

const ProjectOverview: React.FC<ProjectOverviewProps> = ({ project }) => {
  const styles = useStyles();

  let projectImageElement = null;
  if (project.ProjectImage) {
    try {
      const image = JSON.parse(project.ProjectImage);
      projectImageElement = (
        <div className={styles.imageContainer}>
          <Image
            src={image.serverRelativeUrl}
            alt={image.fileName}
            style={{ height: 200, width: 'auto', objectFit: 'contain' }}
          />
        </div>
      );
    } catch {
      // ignore malformed image JSON
    }
  }

  return (
    <div className={styles.container}>
      {projectImageElement}
      <div className={styles.grid}>
        <div className={styles.fieldBlock}>
          <Body1Strong>Status</Body1Strong>
          <Body1>{project.Status}</Body1>
        </div>
        <div className={styles.fieldBlock}>
          <Body1Strong>Sector</Body1Strong>
          <Body1>{project.Sector}</Body1>
        </div>
        <div className={styles.fieldBlock}>
          <Body1Strong>Project Number</Body1Strong>
          <Body1>{project.ProjectNumber}</Body1>
        </div>
        <div className={styles.fieldBlock}>
          <Body1Strong>Project Name</Body1Strong>
          <Body1>{project.ProjectName}</Body1>
        </div>
        <div className={styles.fieldBlock}>
          <Body1Strong>Deltek Code</Body1Strong>
          <Body1>{project.DeltekSubCodes}</Body1>
        </div>
        <div className={styles.fieldBlock}>
          <Body1Strong>Sub Codes</Body1Strong>
          <Body1>{project.SubCodes || '-'}</Body1>
        </div>
        <div className={styles.fieldBlock}>
          <Body1Strong>Client</Body1Strong>
          <Body1>{project.Client}</Body1>
        </div>
        <div className={styles.fieldBlock}>
          <Body1Strong>Client Contact</Body1Strong>
          <Body1>{project.ClientContact}</Body1>
        </div>
        <div className={styles.fullRow}>
          <Body1Strong>Project Description</Body1Strong>
          <Body1>{project.ProjectDescription}</Body1>
        </div>
      </div>
      <Divider />
    </div>
  );
};

export default ProjectOverview;
