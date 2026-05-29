// Centralized constants used across dashboard scripts.
// Exposed on window.AppConstants for reuse without duplicating literals.
(function (window) {
  if (window.AppConstants) {
    return;
  }

  const constants = {
    // Legacy localStorage key
    LS_KEY: 'construction_projects',
    // Firestore collections
    PROJECTS_COL: 'projects',
    DEEPWELLS_COL: 'deepwells',
    REFORESTATION_COL: 'reforestations',
    SERVICE_UPDATES_COL: 'service_updates',
    PERSONAL_TASKS_COL: 'personal_tasks',
    PRESENTATIONS_COL: 'product_presentations',
    OPCR_COL: 'opcr_entries',
    IPCR_COL: 'ipcr_entries',
    CALENDAR_COL: 'calendar_activities',
    NOTIFICATIONS_COL: 'notifications',
    MAIL_COL: 'mail',
    // Admin identity
    ADMIN_EMAIL: 'johnlowel.fradejas@mwss.gov.ph',
    ADMIN_FULL_NAME: 'Fradejas, John Lowel',
    ADMIN_DESIGNATION: 'Principal Engineer C',
    ADMIN_DEPARTMENT: 'FOMD',
    // Pagination defaults
    PERSONAL_TASKS_PER_PAGE: 30,
    PROJECTS_PER_PAGE: 30,
    DEEPWELLS_PER_PAGE: 30,
    SERVICE_UPDATES_PER_PAGE: 50,
    // SUR plant lists
    MWCI_PLANTS: [
      'Balara WTP 1',
      'Balara WTP 2',
      'East LMTP',
      'Luzon WTP',
      'Cardona WTP',
      'Calawis TP',
      'East Bay TP',
    ],
    MWSI_PLANTS: [
      'La Mesa WTP 1',
      'La Mesa WTP 2',
      'Putatan WTP 1',
      'Putatan WTP 2',
      'PQE New Water',
      'Anabu WTP',
      'Poblacion WTP',
      'Nanostone',
      'Cross-Border Flow',
    ],
  };

  constants.ADMIN_EMAIL_LOWER = constants.ADMIN_EMAIL.toLowerCase();

  window.AppConstants = Object.freeze(constants);
})(window);
