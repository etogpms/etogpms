// Projects data service: wraps Firestore access for construction projects.
// Depends on AppFirebase (db) and AppConstants (PROJECTS_COL) being loaded first.
(function (window) {
  if (window.ProjectService) return;
  if (!window.AppFirebase || !window.AppConstants) {
    throw new Error('ProjectService requires AppFirebase and AppConstants loaded first');
  }
  const { db } = window.AppFirebase;
  const { PROJECTS_COL } = window.AppConstants;

  const collection = () => db.collection(PROJECTS_COL);
  const batch = () => db.batch();

  async function save(project) {
    return collection().doc(project.id).set(project);
  }

  async function remove(id) {
    return collection().doc(id).delete();
  }

  function subscribe(handler) {
    return collection().onSnapshot(handler);
  }

  window.ProjectService = Object.freeze({
    collection,
    batch,
    save,
    remove,
    subscribe,
  });
})(window);
