// Generic collection services for various dashboard modules.
// Depends on AppFirebase (db) and AppConstants being loaded first.
(function (window) {
  if (window.DeepwellService) return; // assume all services created together
  if (!window.AppFirebase || !window.AppConstants) {
    throw new Error('Collection services require AppFirebase and AppConstants loaded first');
  }
  const { db } = window.AppFirebase;
  const {
    DEEPWELLS_COL,
    REFORESTATION_COL,
    SERVICE_UPDATES_COL,
    PERSONAL_TASKS_COL,
    PRESENTATIONS_COL,
    OPCR_COL,
    IPCR_COL,
    CALENDAR_COL,
    NOTIFICATIONS_COL,
    MAIL_COL,
  } = window.AppConstants;

  function makeCollectionService(collectionName) {
    const collection = () => db.collection(collectionName);
    return Object.freeze({
      collection,
      batch: () => db.batch(),
      save: (data) => collection().doc(data.id).set(data),
      update: (id, data) => collection().doc(id).update(data),
      remove: (id) => collection().doc(id).delete(),
      subscribe: (handler, options, onError) => {
        return collection().onSnapshot(options || {}, handler, onError);
      },
    });
  }

  window.DeepwellService = makeCollectionService(DEEPWELLS_COL);
  window.ReforestationService = makeCollectionService(REFORESTATION_COL);
  window.ServiceUpdateService = makeCollectionService(SERVICE_UPDATES_COL);
  window.PersonalTaskService = makeCollectionService(PERSONAL_TASKS_COL);
  window.PresentationService = makeCollectionService(PRESENTATIONS_COL);
  window.OPCRService = makeCollectionService(OPCR_COL);
  window.IPCRService = makeCollectionService(IPCR_COL);
  window.CalendarService = makeCollectionService(CALENDAR_COL);
  window.NotificationService = makeCollectionService(NOTIFICATIONS_COL);
  window.EmailService = makeCollectionService(MAIL_COL);
})(window);
