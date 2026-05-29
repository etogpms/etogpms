// User management service for pending/approved users and access updates.
// Depends on AppFirebase (db, auth) and AppConstants (ADMIN_EMAIL) being loaded first.
(function (window) {
  if (window.UserService) return;
  if (!window.AppFirebase) {
    throw new Error('UserService requires AppFirebase loaded first');
  }
  const { db } = window.AppFirebase;

  const usersCol = () => db.collection('users');

  function getPending() {
    return usersCol().where('approved', '==', false).get();
  }

  function getApproved() {
    return usersCol().where('approved', '==', true).get();
  }

  function setAccessLevel(userId, level) {
    return usersCol().doc(userId).update({ accessLevel: level });
  }

  function setApproved(userId, approved, extras = {}) {
    return usersCol().doc(userId).update({ approved, ...extras });
  }

  function deleteUser(userId) {
    return usersCol().doc(userId).delete();
  }

  function getById(id) {
    return usersCol().doc(id).get();
  }

  function docRef(id) {
    return usersCol().doc(id);
  }

  function findByEmail(email, limit = 1) {
    return usersCol().where('email', '==', email).limit(limit).get();
  }

  function setUserDoc(id, data) {
    return usersCol().doc(id).set(data);
  }

  function migrateLegacyEmail(emailExact, emailLower) {
    return usersCol().where('email', 'in', [emailExact, emailLower]).get();
  }

  window.UserService = Object.freeze({
    getPending,
    getApproved,
    setAccessLevel,
    setApproved,
    deleteUser,
    getById,
    docRef,
    findByEmail,
    setUserDoc,
    migrateLegacyEmail,
  });
})(window);
