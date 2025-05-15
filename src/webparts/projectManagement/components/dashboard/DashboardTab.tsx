import * as React from 'react';
import { useState, useEffect } from 'react';
import {
  makeStyles,
  tokens,
  Card,
  CardHeader,
  Text,
  Button,
  Spinner,
} from '@fluentui/react-components';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { IProject } from '../../services/ProjectService';
import { IntegrationService, ISyncData } from '../../services/IntegrationService';

export interface IDashboardTabProps {
  context: WebPartContext;
  project: IProject | undefined;
  integrationService: IntegrationService;
}

interface IProjectMetrics {
  totalTasks: number;
  completedTasks: number;
  upcomingDeadlines: number;
  documentsCount: number;
  budget: {
    allocated: number;
    spent: number;
    remaining: number;
  };
  timeTracking: {
    planned: number;
    actual: number;
    variance: number;
  };
}

const useStyles = makeStyles({
  container: {
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalM,
  },
  header: {
    display: 'flex',
    justifyContent: 'space-between',
    alignItems: 'center',
  },
  metricsGrid: {
    display: 'grid',
    gridTemplateColumns: 'repeat(auto-fit, minmax(300px, 1fr))',
    gap: tokens.spacingHorizontalL,
    marginBottom: tokens.spacingVerticalL,
  },
  card: {
    padding: tokens.spacingHorizontalM,
  },
  metric: {
    display: 'flex',
    flexDirection: 'column',
    alignItems: 'center',
    textAlign: 'center',
  },
  value: {
    fontSize: '2rem',
    fontWeight: tokens.fontWeightSemibold,
    color: tokens.colorBrandForeground1,
  },
  label: {
    fontSize: tokens.fontSizeBase300,
    color: tokens.colorNeutralForeground2,
  },
  powerBiContainer: {
    width: '100%',
    height: '600px',
    border: `1px solid ${tokens.colorNeutralStroke1}`,
    borderRadius: tokens.borderRadiusMedium,
  },
});

const DashboardTab: React.FC<IDashboardTabProps> = ({
  context,
  project,
  integrationService,
}) => {
  const styles = useStyles();
  const [metrics, setMetrics] = useState<IProjectMetrics | null>(null);
  const [isLoading, setIsLoading] = useState(true);
  const [viewPointData, setViewPointData] = useState<ISyncData | undefined>(undefined);
  const [deltekData, setDeltekData] = useState<ISyncData | undefined>(undefined);

  const loadMetrics = async (): Promise<void> => {
    try {
      // This would be replaced with actual API calls
      const mockMetrics: IProjectMetrics = {
        totalTasks: 25,
        completedTasks: 15,
        upcomingDeadlines: 3,
        documentsCount: 42,
        budget: {
          allocated: 100000,
          spent: 45000,
          remaining: 55000,
        },
        timeTracking: {
          planned: 1000,
          actual: 850,
          variance: -150,
        },
      };
      setMetrics(mockMetrics);
    } catch (error) {
      console.error('Error loading metrics:', error);
    }
  };

  const loadIntegrationData = async (): Promise<void> => {
    if (!project) return;

    try {
      const [viewPoint, deltek] = await Promise.all([
        integrationService.getLastSyncData(project.ProjectNumber, 'ViewPoint'),
        integrationService.getLastSyncData(project.ProjectNumber, 'Deltek'),
      ]);

      setViewPointData(viewPoint);
      setDeltekData(deltek);
    } catch (error) {
      console.error('Error loading integration data:', error);
    }
  };

  const refreshData = async (): Promise<void> => {
    setIsLoading(true);
    await Promise.all([loadMetrics(), loadIntegrationData()]);
    setIsLoading(false);
  };

  useEffect(() => {
    if (project) {
      refreshData().catch(console.error);
    }
  }, [project]);

  if (isLoading) {
    return <Spinner />;
  }

  if (!project || !metrics) {
    return (
      <div className={styles.container}>
        <Text>Please select a project to view the dashboard.</Text>
      </div>
    );
  }

  return (
    <div className={styles.container}>
      <div className={styles.header}>
        <Text size={800} weight="semibold">
          Project Dashboard
        </Text>
        <Button
          appearance="primary"
          onClick={() => {
            refreshData().catch(console.error);
          }}
        >
          Refresh Data
        </Button>
      </div>

      <div className={styles.metricsGrid}>
        <Card className={styles.card}>
          <CardHeader header="Task Progress" />
          <div className={styles.metric}>
            <Text className={styles.value}>
              {metrics.completedTasks}/{metrics.totalTasks}
            </Text>
            <Text className={styles.label}>Tasks Completed</Text>
          </div>
        </Card>

        <Card className={styles.card}>
          <CardHeader header="Budget Overview" />
          <div className={styles.metric}>
            <Text className={styles.value}>
              £{metrics.budget.remaining.toLocaleString()}
            </Text>
            <Text className={styles.label}>Remaining Budget</Text>
          </div>
        </Card>

        <Card className={styles.card}>
          <CardHeader header="Time Tracking" />
          <div className={styles.metric}>
            <Text className={styles.value}>
              {metrics.timeTracking.variance > 0 ? '+' : ''}
              {metrics.timeTracking.variance}h
            </Text>
            <Text className={styles.label}>Time Variance</Text>
          </div>
        </Card>

        <Card className={styles.card}>
          <CardHeader header="Documents" />
          <div className={styles.metric}>
            <Text className={styles.value}>{metrics.documentsCount}</Text>
            <Text className={styles.label}>Total Documents</Text>
          </div>
        </Card>
      </div>

      <Card>
        <CardHeader header="Power BI Dashboard" />
        <div className={styles.powerBiContainer}>
          {/* Power BI Embed component would go here */}
          <iframe
            title="Power BI Dashboard"
            width="100%"
            height="100%"
            src={`https://app.powerbi.com/reportEmbed?reportId=${project.ProjectNumber}`}
            
            allowFullScreen
          />
        </div>
      </Card>

      {viewPointData && (
        <Card>
          <CardHeader header="ViewPoint Integration" />
          <pre>{JSON.stringify(viewPointData, null, 2)}</pre>
        </Card>
      )}

      {deltekData && (
        <Card>
          <CardHeader header="Deltek Integration" />
          <pre>{JSON.stringify(deltekData, null, 2)}</pre>
        </Card>
      )}
    </div>
  );
};

export default DashboardTab; 