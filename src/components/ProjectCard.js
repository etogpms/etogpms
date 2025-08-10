import React from 'react';
import {
  Card,
  CardContent,
  Typography,
  Grid,
  Box,
  IconButton,
  Collapse,
  LinearProgress,
  Tooltip,
} from '@mui/material';
import {
  ExpandMore as ExpandMoreIcon,
  Warning as WarningIcon,
  CheckCircle as CheckCircleIcon,
} from '@mui/icons-material';
import { Line } from 'react-chartjs-2';
import {
  Chart as ChartJS,
  CategoryScale,
  LinearScale,
  PointElement,
  LineElement,
  Title,
  Tooltip as ChartTooltip,
  Legend,
} from 'chart.js';

ChartJS.register(
  CategoryScale,
  LinearScale,
  PointElement,
  LineElement,
  Title,
  ChartTooltip,
  Legend
);

const ProjectCard = ({ project }) => {
  const [expanded, setExpanded] = React.useState(false);

  const toggleExpand = () => {
    setExpanded(!expanded);
  };

  const sCurveData = {
    labels: Array.from({ length: 10 }, (_, i) => `Week ${i + 1}`),
    datasets: [
      {
        label: 'Planned',
        data: project.plannedProgress,
        borderColor: '#4CAF50',
        tension: 0.4,
      },
      {
        label: 'Actual',
        data: project.actualProgress,
        borderColor: '#FF9800',
        tension: 0.4,
      },
    ],
  };

  return (
    <Card sx={{ mb: 3 }}>
      <CardContent>
        <Grid container spacing={2}>
          <Grid item xs={12}>
            <Typography variant="h6" gutterBottom>
              {project.name}
            </Typography>
          </Grid>
          <Grid item xs={12} sm={6}>
            <Typography variant="body2" color="text.secondary">
              Contractor: {project.contractor}
            </Typography>
          </Grid>
          <Grid item xs={12} sm={6}>
            <Typography variant="body2" color="text.secondary">
              Contract Amount: {project.currency}{project.amount.toLocaleString()}
            </Typography>
          </Grid>
          <Grid item xs={12}>
            <Box sx={{ display: 'flex', alignItems: 'center' }}>
              <Typography variant="h6" sx={{ mr: 1 }}>
                {project.progress}% Complete
              </Typography>
              <LinearProgress
                variant="determinate"
                value={project.progress}
                sx={{ width: '70%', ml: 1 }}
              />
            </Box>
          </Grid>
          <Grid item xs={12}>
            <IconButton
              onClick={toggleExpand}
              aria-expanded={expanded}
              aria-label="show more"
            >
              <ExpandMoreIcon />
            </IconButton>
          </Grid>
        </Grid>
      </CardContent>
      <Collapse in={expanded} timeout="auto" unmountOnExit>
        <CardContent>
          <Grid container spacing={2}>
            <Grid item xs={12} md={6}>
              <Typography variant="subtitle1" gutterBottom>
                Progress Visualization
              </Typography>
              <Box sx={{ height: 200 }}>
                <Line data={sCurveData} />
              </Box>
            </Grid>
            <Grid item xs={12} md={6}>
              <Typography variant="subtitle1" gutterBottom>
                Key Dates
              </Typography>
              <Box sx={{ p: 1 }}>
                <Typography variant="body2">
                  NTP: {project.ntp}
                </Typography>
                <Typography variant="body2">
                  Original Duration: {project.duration} days
                </Typography>
                <Typography variant="body2">
                  Target Completion: {project.targetCompletion}
                </Typography>
                {project.revisedCompletion && (
                  <Typography variant="body2">
                    Revised Completion: {project.revisedCompletion}
                  </Typography>
                )}
              </Box>
            </Grid>
            <Grid item xs={12}>
              <Typography variant="subtitle1" gutterBottom>
                Current Status
              </Typography>
              <Box sx={{ p: 1 }}>
                <Typography variant="body2">
                  <strong>Issues:</strong>
                  {project.issues.map((issue, index) => (
                    <React.Fragment key={index}>
                      <Box
                        sx={{
                          display: 'flex',
                          alignItems: 'center',
                          mb: 1,
                        }}
                      >
                        <Tooltip title={issue.status === 'resolved' ? 'Resolved' : 'Pending'}>
                          <Box
                            sx={{
                              mr: 1,
                              color: issue.status === 'resolved' ? 'success.main' : 'warning.main',
                            }}
                          >
                            {issue.status === 'resolved' ? <CheckCircleIcon /> : <WarningIcon />}
                          </Box>
                        </Tooltip>
                        {issue.description}
                      </Box>
                    </React.Fragment>
                  ))}
                </Typography>
                <Typography variant="body2">
                  <strong>Remarks:</strong> {project.remarks}
                </Typography>
              </Box>
            </Grid>
          </Grid>
        </CardContent>
      </Collapse>
    </Card>
  );
};

export default ProjectCard;
