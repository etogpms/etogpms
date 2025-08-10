import React from 'react';
import {
  CssBaseline,
  Container,
  AppBar,
  Toolbar,
  Typography,
  Box,
  Grid,
  Paper,
  TextField,
  Select,
  MenuItem,
  FormControl,
  InputLabel,
} from '@mui/material';
import ProjectCard from './components/ProjectCard';

const App = () => {
  const [filter, setFilter] = React.useState({
    contractor: '',
    status: '',
    region: '',
  });

  const [searchTerm, setSearchTerm] = React.useState('');

  // Sample project data (in a real app, this would come from an API)
  const projects = [
    {
      id: 1,
      name: 'Highway Expansion Project',
      contractor: 'ABC Construction Co.',
      amount: 50000000,
      currency: '₱',
      ntp: '2024-01-15',
      duration: 365,
      targetCompletion: '2025-01-15',
      progress: 52,
      plannedProgress: [0, 10, 20, 30, 40, 50, 60, 70, 80, 90],
      actualProgress: [0, 5, 15, 25, 35, 45, 52, 65, 75, 85],
      issues: [
        {
          description: 'Material delivery delay',
          status: 'pending',
        },
        {
          description: 'Weather-related delays',
          status: 'resolved',
        },
      ],
      remarks: 'Project progressing well despite minor delays',
    },
    // Add more sample projects as needed
  ];

  const filteredProjects = projects.filter((project) => {
    const matchesContractor = !filter.contractor || project.contractor === filter.contractor;
    const matchesStatus = !filter.status || project.status === filter.status;
    const matchesRegion = !filter.region || project.region === filter.region;
    const matchesSearch = project.name.toLowerCase().includes(searchTerm.toLowerCase());

    return matchesContractor && matchesStatus && matchesRegion && matchesSearch;
  });

  return (
    <Box sx={{ flexGrow: 1 }}>
      <CssBaseline />
      <AppBar position="static">
        <Toolbar>
          <Typography variant="h6" component="div" sx={{ flexGrow: 1 }}>
            Ongoing Construction Project Tracker
          </Typography>
          <Typography variant="subtitle2" color="text.secondary">
            {new Date().toLocaleDateString()}
          </Typography>
        </Toolbar>
      </AppBar>

      <Container maxWidth="lg" sx={{ mt: 4 }}>
        <Grid container spacing={3}>
          <Grid item xs={12} md={4}>
            <Paper sx={{ p: 2 }}>
              <Typography variant="h6" gutterBottom>
                Filter Projects
              </Typography>
              <FormControl fullWidth sx={{ mb: 2 }}>
                <InputLabel>Contractor</InputLabel>
                <Select
                  value={filter.contractor}
                  onChange={(e) => setFilter({ ...filter, contractor: e.target.value })}
                >
                  <MenuItem value="">All Contractors</MenuItem>
                  <MenuItem value="ABC Construction Co.">ABC Construction Co.</MenuItem>
                  {/* Add more contractor options as needed */}
                </Select>
              </FormControl>
              <FormControl fullWidth sx={{ mb: 2 }}>
                <InputLabel>Status</InputLabel>
                <Select
                  value={filter.status}
                  onChange={(e) => setFilter({ ...filter, status: e.target.value })}
                >
                  <MenuItem value="">All Statuses</MenuItem>
                  <MenuItem value="ongoing">Ongoing</MenuItem>
                  <MenuItem value="delayed">Delayed</MenuItem>
                  <MenuItem value="completed">Completed</MenuItem>
                </Select>
              </FormControl>
              <FormControl fullWidth>
                <InputLabel>Region</InputLabel>
                <Select
                  value={filter.region}
                  onChange={(e) => setFilter({ ...filter, region: e.target.value })}
                >
                  <MenuItem value="">All Regions</MenuItem>
                  {/* Add region options as needed */}
                </Select>
              </FormControl>
            </Paper>
          </Grid>

          <Grid item xs={12} md={8}>
            <TextField
              fullWidth
              label="Search Projects"
              value={searchTerm}
              onChange={(e) => setSearchTerm(e.target.value)}
              sx={{ mb: 2 }}
            />
          </Grid>

          <Grid item xs={12}>
            <Grid container spacing={3}>
              {filteredProjects.map((project) => (
                <Grid item xs={12} md={6} lg={4} key={project.id}>
                  <ProjectCard project={project} />
                </Grid>
              ))}
            </Grid>
          </Grid>
        </Grid>
      </Container>
    </Box>
  );
};

export default App;
