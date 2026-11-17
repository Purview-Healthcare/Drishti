(function () {
  // Microsoft sends its sign-in answer back to this page inside the pop-up. Hand it to the Drishti window
  // that opened the pop-up, and do not start the app in here.
  var looks = function (s) { return /[?#&]state=/.test(s) && /[?#&](code|error)=/.test(s); };
  if (!(looks(location.search || "") || looks(location.hash || ""))) return;
  window.__DRISHTI_MS_CALLBACK = true;
  var say = function (ok) {
    setTimeout(function () {
      var r = document.getElementById("root");
      if (r) r.innerHTML = '<div style="font:15px system-ui,sans-serif;padding:40px;color:#16233b;line-height:1.6">' + (ok ? "Microsoft sign in is complete. You can close this window and go back to Drishti." : "Microsoft sign in could not finish in this window. Close it and try again from Drishti.") + ' <a href="' + location.pathname + '">Open Drishti</a></div>';
    }, 1500);
  };
  try { Promise.resolve(window.msalRedirectBridge.broadcastResponseToMainFrame()).then(function () { say(true); }, function () { say(false); }); }
  catch (e) { say(false); }
})();