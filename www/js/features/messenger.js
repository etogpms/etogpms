// Messenger feature: holds chat messages/unread state and exposes simple setters/getters for reuse.
(function (window) {
  if (window.MessengerFeature) return;

  function init() {
    let messages = [];
    let unread = {};

    function setMessages(list) {
      messages = Array.isArray(list) ? list.slice() : [];
    }
    function getMessages() {
      return messages.slice();
    }
    function setUnreadCounts(map) {
      unread = map ? { ...map } : {};
    }
    function getUnreadCounts() {
      return { ...unread };
    }

    return { setMessages, getMessages, setUnreadCounts, getUnreadCounts };
  }

  window.MessengerFeature = { init };
})(window);
