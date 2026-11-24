/* Settings for the hosted add-in, kept in Outlook's roaming settings: they are
   stored in the user's own mailbox, follow them to any PC where the add-in is
   installed, and are readable only by this add-in. Roaming settings are limited
   to 32 KB per add-in, so sizes are checked before saving. Outside Outlook (a
   plain browser preview) localStorage is used instead. */
((root)=>{
  'use strict';
  const LIMIT=31*1024;
  const DEFAULT_SETTINGS={model:'claude-opus-5-5',name:'',role:'',organisation:'',signature:'',default_profile:'builtin-professional',security_guard:true,
    style:'Use Australian English. Be concise, professional and natural. Preserve names, dates, times, amounts, locations and instructions. Avoid unnecessary apologies, repetition and corporate jargon.'};
  const DEFAULT_PROFILES=[
    {id:'builtin-professional',name:'Professional (default)',purpose:'Everyday business correspondence.',guidance:'Clear, courteous and efficient. Lead with the answer or next step. Short paragraphs.',greeting:'Hi {first name},',signoff:'Kind regards,',formality:3,warmth:3,directness:4,length:'Standard',format:'Paragraphs',avoid:'Corporate jargon, over-apologising, exclamation marks.'},
    {id:'builtin-tenant',name:'Tenant & occupier relations',purpose:'Replies to tenants, occupiers and visitors.',guidance:'Service-focused and reassuring. Acknowledge the concern, explain what will happen and when, and make it easy to reply. Avoid internal jargon and avoid committing to outcomes the notes do not confirm.',greeting:'Hi {first name},',signoff:'Kind regards,',formality:3,warmth:4,directness:3,length:'Standard',format:'Paragraphs',avoid:'Blame, speculation about causes, internal procedure detail.'},
    {id:'builtin-contractor',name:'Contractors & suppliers — firm',purpose:'Instructions, compliance and performance matters.',guidance:'Firm, specific and unambiguous. State requirements, deadlines and consequences only where supplied. Use numbered steps for instructions. Polite but not soft.',greeting:'Hi {first name},',signoff:'Regards,',formality:4,warmth:2,directness:5,length:'Standard',format:'Paragraphs',avoid:'Hedging language, vague timeframes.'},
    {id:'builtin-executive',name:'Executive / client briefing',purpose:'Updates to senior stakeholders, owners and clients.',guidance:'Bottom line first, then key facts, risk and recommended decision. Concise, confident and neutral. Use short bullet points for multiple facts.',greeting:'Hi {first name},',signoff:'Kind regards,',formality:4,warmth:2,directness:5,length:'Brief',format:'Paragraphs',avoid:'Operational minutiae, emotive language.'},
    {id:'builtin-team',name:'Internal team — direct',purpose:'Quick notes to staff and colleagues.',guidance:'Friendly, plain and direct. Get to the point; clear actions and owners.',greeting:'Hi {first name},',signoff:'Thanks,',formality:2,warmth:4,directness:4,length:'Brief',format:'Paragraphs',avoid:''},
  ].map(p=>({...p,signature:'',include_signature:true,examples:[]}));
  const TEXT_LIMITS={name:80,purpose:300,guidance:2000,greeting:200,signoff:200,signature:1000,avoid:1000};

  function backend(){
    const rs=root.Office&&root.Office.context&&root.Office.context.roamingSettings;
    if(rs)return {kind:'mailbox',get:k=>rs.get(k),set:(k,v)=>rs.set(k,v),remove:k=>rs.remove(k),
      save:()=>new Promise((res,rej)=>rs.saveAsync(r=>r.status==='succeeded'?res():rej(new Error((r.error&&r.error.message)||'Outlook could not save your settings.'))))};
    const mem={};
    const ls={get(k){try{const v=localStorage.getItem('dbn.store.'+k);return v===null?undefined:JSON.parse(v);}catch(e){return mem[k];}},
      set(k,v){mem[k]=v;try{localStorage.setItem('dbn.store.'+k,JSON.stringify(v));}catch(e){}},
      remove(k){delete mem[k];try{localStorage.removeItem('dbn.store.'+k);}catch(e){}}};
    return {kind:'browser',get:ls.get,set:ls.set,remove:ls.remove,save:()=>Promise.resolve()};
  }

  function level(v,d=3){if(v===''||v==null)return d;const n=Number(v);if(typeof v==='boolean'||!Number.isFinite(n)||n<1||n>5)throw new Error('Tone controls must be between 1 and 5.');return Math.round(n);}
  function validateProfile(data){
    if(!data||typeof data!=='object')throw new Error('Invalid profile.');
    const p={};
    for(const [k,max] of Object.entries(TEXT_LIMITS)){const v=data[k]==null?'':data[k];if(typeof v!=='string'||v.length>max)throw new Error(`Profile ${k} must be under ${max.toLocaleString()} characters.`);p[k]=v.trim();}
    if(!p.name)throw new Error('Give the profile a name.');
    const id=data.id||'';if(typeof id!=='string'||id.length>64||(id&&!/^[A-Za-z0-9_-]+$/.test(id)))throw new Error('Invalid profile.');
    p.id=id||('p'+Date.now().toString(36)+Math.random().toString(36).slice(2,6));
    for(const k of ['formality','warmth','directness'])p[k]=level(data[k]);
    p.length=data.length||'Standard';p.format=data.format||'Paragraphs';
    if(!['Brief','Standard','Detailed'].includes(p.length)||!['Paragraphs','Bullet points'].includes(p.format))throw new Error('Choose a valid length and layout.');
    const ex=data.examples||[];
    if(!Array.isArray(ex)||ex.length>3||!ex.every(e=>typeof e==='string'&&e.length<=2500))throw new Error('Add up to three example emails of up to 2,500 characters each.');
    p.examples=ex.map(e=>e.trim()).filter(Boolean);
    p.include_signature=data.include_signature!==false;
    return p;
  }

  function create(){
    const b=backend();
    const clone=v=>v===undefined?undefined:JSON.parse(JSON.stringify(v));
    const read=(k,d)=>{const v=b.get(k);return v===undefined||v===null?d:clone(v);};
    function size(){return JSON.stringify({k:read('apiKey',''),s:read('settings',{}),p:read('profiles',null),t:read('templates',[])}).length;}
    async function commit(key,value){
      const before=clone(b.get(key));b.set(key,value);
      if(b.kind==='mailbox'&&size()>LIMIT){if(before===undefined)b.remove(key);else b.set(key,before);
        throw new Error('Outlook add-in storage is full (32 KB). Shorten profile examples or delete unused templates/profiles.');}
      await b.save();
    }
    return {
      kind:b.kind,
      usage(){return {used:size(),limit:LIMIT};},
      apiKey(){return read('apiKey','');},
      async setApiKey(key){
        const k=String(key||'').trim();
        if(k&&(!/^sk-ant-[A-Za-z0-9_-]{20,}$/.test(k)))throw new Error('Paste a Claude API key that starts with sk-ant-.');
        if(k)await commit('apiKey',k);else{b.remove('apiKey');await b.save();}
      },
      settings(){return {...DEFAULT_SETTINGS,...read('settings',{})};},
      async saveSettings(patch){
        const next={...this.settings(),...patch};
        for(const k of ['name','role','organisation']){if(typeof next[k]!=='string'||next[k].length>200)throw new Error(k+' must be under 200 characters.');}
        if(typeof next.signature!=='string'||next.signature.length>1000)throw new Error('Signature must be under 1,000 characters.');
        if(typeof next.style!=='string'||next.style.length>2000)throw new Error('Writing style must be under 2,000 characters.');
        if(typeof next.security_guard!=='boolean')throw new Error('Security guard must be on or off.');
        await commit('settings',next);return next;
      },
      profiles(){return read('profiles',null)||DEFAULT_PROFILES.map(p=>({...p}));},
      async saveProfile(data){
        const p=validateProfile(data),list=this.profiles(),i=list.findIndex(x=>x.id===p.id);
        if(i>=0)list[i]=p;else{if(list.length>=15)throw new Error('You can keep up to 15 profiles.');list.push(p);}
        await commit('profiles',list);return {profiles:list,saved:p};
      },
      async deleteProfile(id){const list=this.profiles().filter(p=>p.id!==id);await commit('profiles',list);return list;},
      templates(){return read('templates',[]);},
      async saveTemplate(name,body){
        name=String(name||'').trim();body=String(body||'').trim();
        if(!name||!body)throw new Error('Give the template a name and some text.');
        if(name.length>100||body.length>4000)throw new Error('Templates are limited to 4,000 characters.');
        const list=this.templates();if(list.length>=30)throw new Error('You can keep up to 30 templates.');
        list.push({id:'t'+Date.now().toString(36),name,body});await commit('templates',list);return list;
      },
      async deleteTemplate(id){const list=this.templates().filter(t=>t.id!==id);await commit('templates',list);return list;},
    };
  }
  const api={create,validateProfile,DEFAULT_PROFILES,DEFAULT_SETTINGS,LIMIT};
  if(typeof module!=='undefined'&&module.exports)module.exports=api;else root.DBNStore=api;
})(typeof window==='undefined'?globalThis:window);
