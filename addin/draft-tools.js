/* Pure helpers for the Outlook task pane: Outlook-friendly HTML, pre-send
   safety checks and draft version history. No DOM or Office.js access here,
   so everything is unit-testable in Node. */
((root)=>{
  'use strict';
  const escape=text=>String(text).replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
  const BULLET=/^\s*(?:[-*•–]|●)\s+(.*)$/;
  const NUMBER=/^\s*(\d{1,2})[.)]\s+(.*)$/;
  const P='<p style="margin:0 0 11pt 0">';

  /* Plain text -> simple HTML that Outlook renders natively: paragraphs, line
     breaks inside a paragraph, and real bulleted / numbered lists. */
  function toHtml(text){
    const blocks=String(text).replace(/\r\n?/g,'\n').split(/\n{2,}/);
    const out=[];
    for(const block of blocks){
      const lines=block.split('\n').filter(l=>l.trim()!=='');
      if(!lines.length)continue;
      let i=0,para=[];
      const flush=()=>{if(para.length){out.push(P+para.map(escape).join('<br>')+'</p>');para=[];}};
      while(i<lines.length){
        const bullet=BULLET.exec(lines[i]),number=NUMBER.exec(lines[i]);
        if(bullet||number){
          flush();
          const tag=bullet?'ul':'ol',items=[],re=bullet?BULLET:NUMBER;
          while(i<lines.length&&re.test(lines[i])){const m=re.exec(lines[i]);items.push(escape(bullet?m[1]:m[2]));i++;}
          out.push('<'+tag+' style="margin:0 0 11pt 0">'+items.map(item=>'<li>'+item+'</li>').join('')+'</'+tag+'>');
        }else{para.push(lines[i]);i++;}
      }
      flush();
    }
    return out.join('');
  }

  const WEEKDAY='(?:mon|tues|wednes|thurs|fri|satur|sun)day';
  const RULES=[
    {id:'placeholder',level:'stop',test:t=>(t.match(/\[(?:confirm|insert|tbc|tba|add|name|date|time|xx+)[^\]\n]{0,80}\]/gi)||[]).length,
     message:n=>`${n} placeholder${n>1?'s':''} still to fill in (e.g. [confirm …]).`},
    {id:'security-code',level:'stop',guard:true,
     test:t=>/\b(?:door|alarm|gate|safe|lock ?box|key ?safe|keypad|access|entry|building|padlock)\s*(?:code|pin|combination)s?\b[^.\n]{0,30}?\b\d{3,8}\b/i.test(t)||/\b(?:pin|passcode|combination)\b\s*(?:is|:|=|-)?\s*\d{3,8}\b/i.test(t),
     message:()=>'Looks like it contains an access/alarm code or PIN. Send codes through a secure channel instead.'},
    {id:'security-password',level:'stop',guard:true,test:t=>/\b(?:password|passwd|pwd|login|username)\s*(?:is|:|=)\s*\S{4,}/i.test(t),
     message:()=>'Looks like it contains a password or login. Remove it before sending.'},
    {id:'security-card',level:'warn',guard:true,test:t=>/\b(?:card|fob|tag|credential)\s*(?:no\.?|number|#|id)?\s*[:#]?\s*\d{5,}\b/i.test(t)||/\b(?:\d[ -]?){13,19}\b/.test(t),
     message:()=>'Contains what looks like an access card or payment card number. Check it should be shared.'},
    {id:'security-ops',level:'warn',guard:true,test:t=>/\b(?:blind spot|not covered by cctv|cctv (?:is )?(?:down|offline|not working)|unmanned|no guard|patrol (?:times|schedule)|roster)\b/i.test(t),
     message:()=>'Mentions security coverage or rostering detail. Confirm the recipient needs it.'},
    {id:'attachment',level:'warn',test:(t,ctx)=>!ctx.hasAttachments&&/\b(?:attached|attachment|enclosed|please find)\b/i.test(t),
     message:()=>'Mentions an attachment — remember to attach it.'},
    {id:'commitment',level:'info',test:t=>/\b(?:guarantee|promise|refund|compensat\w*|liab\w*|at no (?:cost|charge)|free of charge|waive\w*)\b/i.test(t),
     message:()=>'Contains a commitment (refund, guarantee, liability or cost). Make sure it is approved.'},
    {id:'deadline',level:'info',test:t=>new RegExp('\\b(?:by|before|no later than|until)\\s+(?:'+WEEKDAY+'|tomorrow|today|cob|eod|close of business|\\d{1,2}(?:st|nd|rd|th)?\\b)','i').test(t),
     message:()=>'Contains a deadline — double-check the date/time is right.'},
  ];

  /* Pre-send review. Returns [{id, level: 'stop'|'warn'|'info', message}]. */
  function review(text,ctx={}){
    const body=String(text||''),found=[];
    for(const rule of RULES){
      if(rule.guard&&ctx.securityGuard===false)continue;
      const hit=rule.test(body,ctx);
      if(hit)found.push({id:rule.id,level:rule.level,message:rule.message(typeof hit==='number'?hit:1)});
    }
    return found;
  }

  /* Linear version history with back/forward, like undo for whole drafts. */
  function history(limit=20){
    let items=[],index=-1;
    return {
      push(version){items=items.slice(0,index+1);items.push(version);if(items.length>limit)items.shift();index=items.length-1;return version;},
      current(){return items[index]||null;},
      back(){if(index>0)index--;return items[index];},
      forward(){if(index<items.length-1)index++;return items[index];},
      update(patch){if(items[index])items[index]={...items[index],...patch};},
      get position(){return index+1;},get size(){return items.length;},
      reset(){items=[];index=-1;}
    };
  }

  function firstName(name){
    const clean=String(name||'').replace(/\(.*?\)|<.*?>/g,'').trim();
    if(!clean||/@/.test(clean))return '';
    if(clean.includes(','))return clean.split(',')[1].trim().split(/\s+/)[0]||'';
    return clean.split(/\s+/)[0];
  }

  const api={escape,toHtml,review,history,firstName};
  if(typeof module!=='undefined'&&module.exports)module.exports=api;else root.DBNDraft=api;
})(typeof window==='undefined'?{}:window);
