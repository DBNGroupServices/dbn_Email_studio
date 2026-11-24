/* Start-up watchdog, written in plain ES5 so it runs even in an older Outlook
   browser engine. If the pane hasn't started within 12 seconds, it shows why
   on screen (script errors, Office.js status, browser version) instead of
   leaving "Connecting..." forever. */
(function(){
  var errors=[];
  window.__dbnErrors=errors;
  window.addEventListener('error',function(e){
    var where=e.filename?(' ('+String(e.filename).split('/').pop()+':'+e.lineno+')'):'';
    errors.push(String(e.message||'Script error')+where);
  },true);
  window.addEventListener('unhandledrejection',function(e){
    var r=e.reason;errors.push('Unhandled: '+String(r&&r.message||r));
  });
  function report(){
    if(window.__dbnStarted)return;
    var el=document.getElementById('status');
    if(!el)return;
    var lines=['Email Studio could not finish starting.'];
    if(typeof window.Office==='undefined')lines.push('Office.js did not load (network or firewall blocking appsforoffice.microsoft.com?).');
    else if(!window.__dbnOfficeReady)lines.push('Outlook did not confirm the add-in was ready (Office.onReady).');
    if(!window.DBNClaude)lines.push('Claude library did not load.');
    for(var i=0;i<errors.length&&i<5;i++)lines.push('Error: '+errors[i]);
    lines.push('Browser: '+navigator.userAgent);
    lines.push('Close the pane and reopen it. If this stays, send a screenshot of this message.');
    el.className='status error';
    el.textContent=lines.join('\n');
  }
  window.__dbnReport=report;
  setTimeout(report,12000);
})();
