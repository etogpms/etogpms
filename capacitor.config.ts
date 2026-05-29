import type { CapacitorConfig } from '@capacitor/cli';

const config: CapacitorConfig = {
  appId: 'com.mwss.projectmonitoring',
  appName: 'Project Monitoring',
  webDir: 'www',
  // Comment out or remove the 'server' block below for production/release builds.
  // With this enabled, the app loads assets directly from your local dev server.
  /*
  server: {
    url: 'http://localhost:8080',
    cleartext: true
  }
  */
};

export default config;
