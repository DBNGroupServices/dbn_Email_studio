/* Shows the right message for ?result=success|cancelled|updated (Stripe return pages). */
(function(){
  var r=new URLSearchParams(location.search).get('result');
  var set=function(id,t){document.getElementById(id).textContent=t;};
  if(r==='cancelled'){set('eyebrow','Checkout cancelled');set('title','No charge made.');set('text','You can choose a plan any time from the Assistant pane in Outlook.');}
  else if(r==='updated'){set('eyebrow','Billing');set('title','All set.');set('text','Your billing changes are saved. Go back to Outlook: the Assistant pane shows your plan within a minute.');}
})();
