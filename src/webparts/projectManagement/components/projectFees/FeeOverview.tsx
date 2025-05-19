import * as React from 'react';
import {
  makeStyles,
  tokens,
  Card,
  Text,
} from '@fluentui/react-components';
import {
  BarChart,
  LineChart,
  PieChart,
  ResponsiveContainer,
  XAxis,
  YAxis,
  CartesianGrid,
  Tooltip,
  Legend,
  Line,
  Bar,
  Pie,
  Cell,
} from 'recharts';

const useStyles = makeStyles({
  root: {
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalM,
    padding: tokens.spacingVerticalL,
  },
  summaryCards: {
    display: 'grid',
    gridTemplateColumns: 'repeat(auto-fit, minmax(200px, 1fr))',
    gap: tokens.spacingHorizontalM,
    marginBottom: tokens.spacingVerticalL,
  },
  card: {
    padding: tokens.spacingVerticalM,
    textAlign: 'center',
  },
  chartContainer: {
    display: 'grid',
    gridTemplateColumns: '1fr 1fr',
    gap: tokens.spacingHorizontalL,
    marginBottom: tokens.spacingVerticalL,
  },
  chart: {
    height: '300px',
    backgroundColor: tokens.colorNeutralBackground2,
    padding: tokens.spacingVerticalM,
    borderRadius: tokens.borderRadiusMedium,
  },
  fullWidthChart: {
    gridColumn: '1 / -1',
    height: '300px',
    backgroundColor: tokens.colorNeutralBackground2,
    padding: tokens.spacingVerticalM,
    borderRadius: tokens.borderRadiusMedium,
  },
});

// Dummy data for charts
const monthlySpendData = [
  { month: 'Jan', spend: 12000 },
  { month: 'Feb', spend: 15000 },
  { month: 'Mar', spend: 18000 },
  { month: 'Apr', spend: 16000 },
  { month: 'May', spend: 21000 },
  { month: 'Jun', spend: 19000 },
];

const feeDistributionData = [
  { name: 'Design', value: 35000 },
  { name: 'Development', value: 45000 },
  { name: 'Testing', value: 20000 },
  { name: 'Management', value: 25000 },
];

const budgetVsActualData = [
  { category: 'Design', budget: 40000, actual: 35000 },
  { category: 'Development', budget: 50000, actual: 45000 },
  { category: 'Testing', budget: 25000, actual: 20000 },
  { category: 'Management', budget: 30000, actual: 25000 },
];

const COLORS = ['#0088FE', '#00C49F', '#FFBB28', '#FF8042'];

interface PieLabelProps {
  name: string;
  percent: number;
}

const FeeOverview: React.FC = () => {
  const styles = useStyles();

  return (
    <div className={styles.root}>
      <div className={styles.summaryCards}>
        <Card className={styles.card}>
          <Text size={800} weight="semibold">£125,000</Text>
          <Text>Total Budget</Text>
        </Card>
        <Card className={styles.card}>
          <Text size={800} weight="semibold">£101,000</Text>
          <Text>Total Spend</Text>
        </Card>
        <Card className={styles.card}>
          <Text size={800} weight="semibold">£24,000</Text>
          <Text>Remaining Budget</Text>
        </Card>
        <Card className={styles.card}>
          <Text size={800} weight="semibold">81%</Text>
          <Text>Budget Utilized</Text>
        </Card>
      </div>

      <div className={styles.chartContainer}>
        <div className={styles.chart}>
          <Text weight="semibold" size={400}>Monthly Spend Trend</Text>
          <ResponsiveContainer width="100%" height="90%">
            <LineChart data={monthlySpendData}>
              <CartesianGrid strokeDasharray="3 3" />
              <XAxis dataKey="month" />
              <YAxis />
              <Tooltip />
              <Legend />
              <Line 
                type="monotone" 
                dataKey="spend" 
                stroke="#8884d8" 
                name="Monthly Spend"
              />
            </LineChart>
          </ResponsiveContainer>
        </div>

        <div className={styles.chart}>
          <Text weight="semibold" size={400}>Fee Distribution</Text>
          <ResponsiveContainer width="100%" height="90%">
            <PieChart>
              <Pie
                data={feeDistributionData}
                cx="50%"
                cy="50%"
                labelLine={false}
                outerRadius={80}
                fill="#8884d8"
                dataKey="value"
                label={({ name, percent }: PieLabelProps) => `${name} ${(percent * 100).toFixed(0)}%`}
              >
                {feeDistributionData.map((entry, index) => (
                  <Cell key={`cell-${index}`} fill={COLORS[index % COLORS.length]} />
                ))}
              </Pie>
              <Tooltip />
              <Legend />
            </PieChart>
          </ResponsiveContainer>
        </div>

        <div className={styles.fullWidthChart}>
          <Text weight="semibold" size={400}>Budget vs Actual by Category</Text>
          <ResponsiveContainer width="100%" height="90%">
            <BarChart data={budgetVsActualData}>
              <CartesianGrid strokeDasharray="3 3" />
              <XAxis dataKey="category" />
              <YAxis />
              <Tooltip />
              <Legend />
              <Bar dataKey="budget" fill="#8884d8" name="Budget" />
              <Bar dataKey="actual" fill="#82ca9d" name="Actual" />
            </BarChart>
          </ResponsiveContainer>
        </div>
      </div>
    </div>
  );
};

export default FeeOverview; 