// Shared access gate for every GT page except the landing page. Visitors log
// in once on index.html (roster ID, Master ID or master password); the result
// is kept for the browser session as `gtUser` = { name, id, master }.
(function(){
  var AUTH_KEY = 'gtGoalsAuth';
  var USER_KEY = 'gtUser';
  var user = null;
  try { user = JSON.parse(sessionStorage.getItem(USER_KEY) || 'null'); } catch (e) {}

  if (sessionStorage.getItem(AUTH_KEY) !== '1' || !user){
    document.documentElement.style.display = 'none';
    location.replace('index.html');
    return;
  }
  window.gtUser = user;

  document.addEventListener('DOMContentLoaded', function(){
    var el = document.getElementById('recentNavLink');
    if (el) el.style.display = '';
  });
})();

function gtLogout(){
  sessionStorage.removeItem('gtGoalsAuth');
  sessionStorage.removeItem('gtUser');
  location.href = 'index.html';
}
