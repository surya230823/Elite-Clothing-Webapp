/**
 * Session storage + demo login/logout.
 * Used by: index.html, pages/login.html
 * Logout sends users to pages/login.html (path relative to site root).
 */
(function () {
  var SESSION_KEY = "elite_ops_session";

  /** Demo credentials — replace with real auth when you add a backend. */
  var DEMO_USERNAME = "elite";
  var DEMO_PASSWORD = "elite";

  function isLoggedIn() {
    return sessionStorage.getItem(SESSION_KEY) === "1";
  }

  function setLoggedIn() {
    sessionStorage.setItem(SESSION_KEY, "1");
  }

  function clearSession() {
    sessionStorage.removeItem(SESSION_KEY);
  }

  function validateCredentials(username, password) {
    return username === DEMO_USERNAME && password === DEMO_PASSWORD;
  }

  function logout() {
    clearSession();
    window.location.href = "pages/login.html";
  }

  window.EliteAuth = {
    SESSION_KEY: SESSION_KEY,
    DEMO_USERNAME: DEMO_USERNAME,
    DEMO_PASSWORD: DEMO_PASSWORD,
    isLoggedIn: isLoggedIn,
    setLoggedIn: setLoggedIn,
    validateCredentials: validateCredentials,
    logout: logout,
  };
})();
