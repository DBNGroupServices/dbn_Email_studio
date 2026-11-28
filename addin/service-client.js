/* Bulwark HQ Assistant — subscription service client.
   Signs in with the user's Microsoft account (Office single sign-on) and calls the
   Bulwark HQ service, which holds the AI key and the subscription. Email content is
   sent for drafting only; the service does not store it. */
(function(root,factory){
  const api=factory();
  if(typeof module==='object'&&module.exports)module.exports=api;else root.DBNService=api;
})(typeof self!=='undefined'?self:this,function(){
  'use strict';
  class ServiceError extends Error{
    constructor(message,code,status,account){super(message);this.name='ServiceError';this.code=code||'';this.status=status||0;this.account=account||null;}
  }
  const SSO_HELP={
    13001:'Sign in to Outlook with your Microsoft work or personal account, then reopen the pane.',
    13002:'Sign-in was cancelled. Choose Sign in to try again.',
    13003:'This account type is not supported. Sign in to Outlook with a Microsoft 365 work or personal account.',
    13006:'Outlook could not sign you in. Restart Outlook and try again.',
    13007:'Outlook could not get a sign-in token. Check your connection and try again.',
    13012:'Sign-in is not available in this version of Outlook. Update Outlook, or use Outlook on the web.',
  };
  function create(opts){
    const o=opts||{};
    const base=String(o.baseUrl||'').replace(/\/$/,'');
    const fetchFn=o.fetch||((...a)=>fetch(...a));
    let token=null,tokenAt=0;
    async function getToken(force,interactive){
      if(!force&&token&&Date.now()-tokenAt<45*60000)return token;
      if(o.getToken){token=await o.getToken();tokenAt=Date.now();return token;}
      const auth=typeof Office!=='undefined'&&Office.auth;
      if(!auth||typeof auth.getAccessToken!=='function')throw new ServiceError(SSO_HELP[13012],'sso');
      try{
        token=await auth.getAccessToken({allowSignInPrompt:interactive!==false,allowConsentPrompt:interactive!==false,forMSGraphAccess:false});
        tokenAt=Date.now();return token;
      }catch(e){
        const code=e&&e.code;
        throw new ServiceError(SSO_HELP[code]||('Sign-in failed'+(code?' ('+code+')':'')+'. Close the pane and try again.'),'sso');
      }
    }
    async function call(path,body,retried){
      const t=await getToken(!!retried);
      let res;
      try{
        res=await fetchFn(base+path,{method:body===undefined?'GET':'POST',
          headers:{authorization:'Bearer '+t,...(body===undefined?{}:{'content-type':'application/json'})},
          body:body===undefined?undefined:JSON.stringify(body)});
      }catch(e){throw new ServiceError('Could not reach Bulwark HQ. Check your internet connection and try again.','network');}
      let data={};try{data=await res.json();}catch(e){/* empty */}
      if(res.status===401&&!retried){token=null;return call(path,body,true);}
      if(!res.ok)throw new ServiceError(data.error||('Request failed ('+res.status+').'),data.code||'',res.status,data.account||null);
      return data;
    }
    return{
      signIn:()=>getToken(true,true),
      me:()=>call('/api/me'),
      suggest:(email,settings)=>call('/api/suggest',{email,settings}),
      generate:(task,settings,profile)=>call('/api/generate',{task,settings,profile:profile||null}),
      checkout:(plan,interval,seats)=>call('/api/billing/checkout',{plan,interval,seats}),
      portal:()=>call('/api/billing/portal',{}),
    };
  }
  /* "N working days" left until an end time, counting Mon–Fri in Melbourne. */
  function workingDaysLeft(endsAt,now,tz){
    const zone=tz||'Australia/Melbourne',n=now||Date.now();if(!endsAt||n>=endsAt)return 0;
    const fmt=new Intl.DateTimeFormat('en-AU',{timeZone:zone,weekday:'short',year:'numeric',month:'2-digit',day:'2-digit'});
    let count=0,last='';
    for(let t=n;t<endsAt;t+=3600000){
      const parts=fmt.formatToParts(new Date(t)),day=parts.filter(p=>p.type!=='weekday').map(p=>p.value).join('');
      if(day===last)continue;last=day;
      const wd=(parts.find(p=>p.type==='weekday')||{}).value;if(wd!=='Sat'&&wd!=='Sun')count++;
    }
    return count;
  }
  function planLine(acc,now){
    if(!acc)return'';
    if(acc.plan==='trial'){const d=workingDaysLeft(acc.trial&&acc.trial.ends_at,now),r=acc.trial?acc.trial.replies_left:0;
      return `Free trial · ${d} working day${d===1?'':'s'} · ${r} repl${r===1?'y':'ies'} left`;}
    if(acc.plan==='expired')return'Trial ended · choose a plan to keep going';
    const u=acc.usage||{};
    return `${acc.label}${acc.interval?' · '+(acc.interval==='year'?'annual':'monthly'):''} · ${u.replies||0} of ${u.replies_limit||0} replies this month`;
  }
  const PRICING=[
    {plan:'essentials',name:'Essentials',month:15,year:150,blurb:'300 replies a month · Claude Sonnet',seats:false},
    {plan:'professional',name:'Professional',month:29,year:290,blurb:'600 premium Opus replies, then Sonnet · fair use',seats:false,best:true},
    {plan:'team',name:'Team',month:25,year:250,blurb:'Professional for each person · from 3 seats',seats:true},
  ];
  return{create,ServiceError,workingDaysLeft,planLine,PRICING};
});
