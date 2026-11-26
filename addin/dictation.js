/* Dictation for the task pane.
   Uses the browser's speech recognition (Web Speech API) where Outlook's
   browser engine provides it and allows the microphone. Where it doesn't, the
   mic button focuses the box and explains Windows' built-in dictation
   (Win + H), which types into whatever box has focus. No audio is stored. */
((root)=>{
  'use strict';
  const FALLBACK='Dictation isn’t available inside Outlook here. Press Win + H to use Windows dictation; it types into this box.';
  function create(opts){
    const {button,field,onState,onMessage,lang}=opts;
    const SR=root.SpeechRecognition||root.webkitSpeechRecognition;
    let rec=null,listening=false,base='',blocked=!SR;
    function set(on){listening=on;button.classList.toggle('listening',on);button.setAttribute('aria-pressed',String(on));
      button.setAttribute('aria-label',on?'Stop dictation':'Start dictation');if(onState)onState(on);}
    function fallback(text){blocked=true;set(false);field.focus();if(onMessage)onMessage(text||FALLBACK);}
    function start(){
      if(blocked)return fallback();
      try{
        rec=new SR();rec.lang=lang||'en-AU';rec.continuous=true;rec.interimResults=true;
        base=field.value&&!/\s$/.test(field.value)?field.value+' ':field.value;
        rec.onresult=e=>{
          let finalText='',interim='';
          for(let i=0;i<e.results.length;i++){const r=e.results[i];(r.isFinal?(t=>finalText+=t):(t=>interim+=t))(r[0].transcript);}
          field.value=base+finalText+interim;
          field.dispatchEvent(new Event('input',{bubbles:true}));
        };
        rec.onerror=e=>{
          const code=e&&e.error;
          if(code==='not-allowed'||code==='service-not-allowed'||code==='network'||code==='audio-capture')fallback(code==='audio-capture'?'No microphone was found. Check your headset or mic, or press Win + H for Windows dictation.':FALLBACK);
          else if(code!=='no-speech'&&code!=='aborted'&&onMessage)onMessage('Dictation stopped ('+code+'). Click the mic to try again.');
        };
        rec.onend=()=>{if(listening)set(false);};
        rec.start();set(true);
        if(onMessage)onMessage('');
      }catch(error){fallback();}
    }
    function stop(){try{rec&&rec.stop();}catch(e){}set(false);}
    button.addEventListener('click',()=>{listening?stop():start();});
    return {start,stop,get listening(){return listening;},get supported(){return !blocked;}};
  }
  const api={create,FALLBACK};
  if(typeof module!=='undefined'&&module.exports)module.exports=api;else root.DBNDictation=api;
})(typeof window==='undefined'?globalThis:window);
