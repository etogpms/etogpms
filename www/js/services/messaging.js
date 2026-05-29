// Messaging service for chat threads and messages.
// Depends on AppFirebase (db) being loaded first.
(function (window) {
  if (window.MessagingService) return;
  if (!window.AppFirebase) throw new Error('MessagingService requires AppFirebase');
  const { db } = window.AppFirebase;

  const messagesCol = () => db.collection('messages');

  function sendMessage(data) {
    return messagesCol().add(data);
  }

  function deleteMessage(id) {
    return messagesCol().doc(id).delete();
  }

  function updateMessage(id, payload) {
    return messagesCol().doc(id).update(payload);
  }

  function fetchLatest(limit = 50) {
    return messagesCol().orderBy('timestamp', 'desc').limit(limit).get();
  }

  function fetchOlder(beforeTs, limit = 50) {
    return messagesCol()
      .orderBy('timestamp', 'desc')
      .startAfter(beforeTs)
      .limit(limit)
      .get();
  }

  function subscribe(handler, limit = 50) {
    return messagesCol()
      .orderBy('timestamp', 'desc')
      .limit(limit)
      .onSnapshot(handler);
  }

  window.MessagingService = Object.freeze({
    sendMessage,
    deleteMessage,
    updateMessage,
    fetchLatest,
    fetchOlder,
    subscribe,
  });
})(window);
