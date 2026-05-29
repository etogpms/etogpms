// Personal Tasks feature: thin state holder + render bridge to existing renderPersonalTasks.
(function (window) {
  if (window.PersonalTasksFeature) return;

  function init(opts) {
    const { elements, auth, personalTaskService, utils } = opts || {};
    if (!elements || !auth || !personalTaskService || !utils) {
      throw new Error('PersonalTasksFeature.init missing required dependencies');
    }
    let personalTasks = [];

    function setPersonalTasks(list) {
      personalTasks = Array.isArray(list) ? list.slice() : [];
    }
    function getPersonalTasks() {
      return personalTasks.slice();
    }
    function render() {
      if (typeof window.renderPersonalTasks === 'function') {
        try { window.renderPersonalTasks(); } catch (_) { }
      }
    }

    return { setPersonalTasks, getPersonalTasks, render };
  }

  window.PersonalTasksFeature = { init };
})(window);
