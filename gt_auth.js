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

  // Roster IDs only get their own Dashboard; Master IDs get every page.
  var page = location.pathname.split('/').pop() || 'index.html';
  if (!user.master && page !== 'gt_dashboard.html'){
    document.documentElement.style.display = 'none';
    location.replace('gt_dashboard.html');
    return;
  }

  document.addEventListener('DOMContentLoaded', function(){
    var links = document.querySelectorAll('.nav-links a');
    for (var i = 0; i < links.length; i++){
      var href = links[i].getAttribute('href');
      if (!user.master && href !== 'gt_dashboard.html') links[i].style.display = 'none';
      else if (links[i].id === 'recentNavLink') links[i].style.display = '';
    }
  });
})();

function gtLogout(){
  sessionStorage.removeItem('gtGoalsAuth');
  sessionStorage.removeItem('gtUser');
  location.href = 'index.html';
}
