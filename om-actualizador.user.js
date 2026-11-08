// ==UserScript==
// @name         O&M TowerNex · actualizador
// @namespace    lps-om-towernex
// @version      1.0.0
// @description  Actualiza la página O&M LPS/TowerNex con Sitetracker y el SharePoint de TowerNex usando tu propia sesión. SOLO LECTURA: no crea, edita ni borra nada en TowerNex.
// @author       LPS Grupo
// @match        https://krlos8586.github.io/-om-towernex/*
// @match        https://towernex.my.site.com/scp/*
// @run-at       document-idle
// @noframes
// @grant        GM_xmlhttpRequest
// @grant        GM_setValue
// @grant        GM_getValue
// @grant        GM_deleteValue
// @grant        GM_addValueChangeListener
// @grant        GM_removeValueChangeListener
// @grant        GM_openInTab
// @connect      towernexcolombiasas.sharepoint.com
// @updateURL    https://krlos8586.github.io/-om-towernex/om-actualizador.user.js
// @downloadURL  https://krlos8586.github.io/-om-towernex/om-actualizador.user.js
// ==/UserScript==
(function(){
'use strict';
const VERSION='1.0.0';
const SP_HOST='https://towernexcolombiasas.sharepoint.com';
const SP_BASE='/sites/TowerNex-ContratistasOM/';
const ST_URL='https://towernex.my.site.com/scp/s/';

/* ---------- 1. En la página O&M (GitHub): puente para que la página lea las fuentes ---------- */
if(location.hostname==='krlos8586.github.io'){
  document.documentElement.setAttribute('data-omgm',VERSION);
  const reply=(src,id,ok,data,transfer)=>{try{src.postMessage({om:'gm-res',id,ok,data},'*',transfer||[])}catch(e){}};
  window.addEventListener('message',ev=>{
    const m=ev.data;if(!m||m.om!=='gm'||!ev.source)return;
    // solo la propia página (mismo origen) puede pedir
    if(ev.origin!==location.origin&&ev.origin!=='null')return;
    if(m.op==='ping')return reply(ev.source,m.id,true,{version:VERSION});
    if(m.op==='sp')return spGet(m).then(d=>reply(ev.source,m.id,true,d,d.buf?[d.buf]:[])).catch(e=>reply(ev.source,m.id,false,String(e&&e.message||e)));
    if(m.op==='st')return stRun().then(d=>reply(ev.source,m.id,true,d)).catch(e=>reply(ev.source,m.id,false,String(e&&e.message||e)));
  });
  // SharePoint: solo GET y solo dentro del sitio de contratistas O&M
  function spGet(m){return new Promise((res,rej)=>{
    const path=String(m.path||'');
    if(!path.startsWith(SP_BASE))return rej(new Error('ruta no permitida'));
    const headers={Accept:'application/json;odata=nometadata'};if(m.range)headers.Range='bytes='+m.range;
    GM_xmlhttpRequest({method:'GET',url:SP_HOST+path,headers,responseType:m.bin?'arraybuffer':'text',timeout:60000,
      onload:r=>{if(r.status===401||r.status===403||/login\.microsoftonline/.test(r.finalUrl||''))return rej(new Error('login-sp'));
        if(r.status>=400)return rej(new Error('SP '+r.status));
        if(m.bin)res({status:r.status,buf:r.response});else res({status:r.status,text:r.responseText})},
      onerror:()=>rej(new Error('red-sp')),ontimeout:()=>rej(new Error('timeout-sp'))})})}
  // Sitetracker: abre una pestaña en segundo plano con tu sesión, lee y se cierra
  function stRun(){return new Promise((res,rej)=>{
    const id=Date.now()+'-'+Math.random().toString(36).slice(2,7);let tab=null,done=false;
    const fin=(f,v)=>{if(done)return;done=true;try{GM_removeValueChangeListener(lid)}catch(e){}try{tab&&tab.close()}catch(e){}GM_deleteValue('om_req');f(v)};
    const lid=GM_addValueChangeListener('om_res',(n,o,v,remote)=>{if(!v||v.id!==id)return;GM_deleteValue('om_res');v.error?fin(rej,new Error(v.error)):fin(res,v.data)});
    GM_setValue('om_req',{id,ts:Date.now()});
    tab=GM_openInTab(ST_URL+'?om='+id,{active:false,insert:true,setParent:true});
    setTimeout(()=>fin(rej,new Error('timeout-st')),120000);
  })}
  return;
}

/* ---------- 2. En Sitetracker: solo actúa si la página O&M lo pidió hace menos de 2 minutos ---------- */
if(location.hostname==='towernex.my.site.com'){
  const req=GM_getValue('om_req');if(!req||Date.now()-req.ts>120000)return;
  if(window.top!==window)return;
  const send=(o)=>GM_setValue('om_res',Object.assign({id:req.id},o));
  if(/\/login/i.test(location.pathname)||document.querySelector('input[type=password]'))return send({error:'login-st'});
  window.addEventListener('message',ev=>{const m=ev.data;if(!m||m.omst!==req.id)return;m.error?send({error:m.error}):send({data:m.data})});
  // código que corre dentro de la página (necesita la sesión de Sitetracker); solo usa la acción de LECTURA getListRecordsById
  const code=`(async()=>{const ID=${JSON.stringify(req.id)};const post=o=>window.postMessage(Object.assign({omst:ID},o),'*');
  try{let t=0;while(!(window.$A&&$A.clientService)&&t<60){await new Promise(r=>setTimeout(r,500));t++}
   if(!window.$A)throw new Error('sin-aura');
   let tok=null;for(let i=0;i<40&&!tok;i++){for(const k of Object.keys($A.clientService)){const v=$A.clientService[k];if(typeof v==='string'&&v.length>40&&/^eyJ/.test(v)){tok=v;break}}if(!tok)await new Promise(r=>setTimeout(r,500))}
   if(!tok)throw new Error('sin-token');
   const h=await (await fetch('/scp/s/',{credentials:'include'})).text();
   const fw=(h.match(/"fwuid":"([^"]+)"/)||[])[1],app=(h.match(/"app":"([^"]+)"/)||[])[1],ld=(h.match(/"loaded":(\\{[^}]*\\})/)||[])[1];
   const ctx=JSON.stringify({mode:'PROD',fwuid:fw,app:app,loaded:ld?JSON.parse(ld):{},dn:[],globals:{},uad:true});
   let seq=100;async function call(p){const b=new URLSearchParams();b.set('message',JSON.stringify({actions:[{id:(++seq)+';a',descriptor:'aura://ListUiController/ACTION$getListRecordsById',callingDescriptor:'UNKNOWN',params:p}]}));b.set('aura.context',ctx);b.set('aura.pageURI','/scp/s/');b.set('aura.token',tok);
    const r=await fetch('/scp/s/sfsites/aura?r='+seq+'&aura.ListUiController.getListRecordsById=1',{method:'POST',headers:{'Content-Type':'application/x-www-form-urlencoded; charset=UTF-8'},body:b.toString(),credentials:'include'});
    const j=JSON.parse((await r.text()).replace(/^while\\(1\\);/,''));if(!j.actions||j.actions[0].state!=='SUCCESS')throw new Error('consulta-rechazada');return j.actions[0].returnValue}
   const P='sitetracker__',J='sitetracker__Job__c.',Sv='sitetracker__Site__c.';
   const jf=['Name',P+'Job_Type__c',P+'Category__c',P+'Sub_Category__c',P+'Description__c',P+'Customer_Comments__c',P+'Completion_Notes__c',P+'Preferred_Start__c',P+'Due_Date__c',P+'Scheduled_Start__c',P+'Scheduled_End__c',P+'Actual_Start__c',P+'Actual_End__c',P+'Is_Job_Delayed__c','Report_Date__c',P+'Job_Status__c',P+'Priority__c',P+'Activity__c',P+'Related_Job_Reason__c',P+'Site__r.Name',P+'Vendor__r.Name',P+'Job_Template__r.Name',P+'Ticket__r.Name',P+'Customer__r.Name',P+'Related_Job__r.Name'];
   const sf=['Name','Site_ID__c',P+'City__c',P+'State__c',P+'Full_Address__c',P+'Site_Type__c',P+'Site_Description__c',P+'Site_Status__c',P+'Lat__c',P+'Long__c','Structure_type_I__c','Structure_type_II__c','Structure_Height__c','Lifeline__c','LifelineVendor__c','Lifeline_Manufacture_Date__c','Lifeline_Certification_Date__c','AccessType__c','st_telecom__Access_Instructions__c','st_telecom__Visiting_Hours__c','st_telecom__Site_Restrictions__c','Environment_tipology__c','Preventive_Contractor__r.Name','Corrective_Contractor__r.Name','OYM__c','Power_supply_type__c','st_telecom__Site_Name__c','Claro_Site_Code__c','On_Air_Year__c','st_telecom__Power_Provider__c','Structure_Factory__c'];
   const val=(rec,f)=>{const ps=f.split('.');let o=rec.fields[ps[0]];for(const p of ps.slice(1)){o=o&&o.value&&o.value.fields&&o.value.fields[p]}return o?(o.displayValue??o.value):null};
   const raw=(rec,f)=>{const o=rec.fields[f];return o?o.value:null};
   const key=f=>f.replace(/sitetracker__|st_telecom__/g,'').replace(/__c|__r/g,'');
   const rj=await call({listViewId:'00Bbn0000099k01EAA',pageSize:2000,optionalFields:jf.map(x=>J+x)});
   const rs=await call({listViewId:'00Bbn0000099Vc5EAE',pageSize:2000,optionalFields:sf.map(x=>Sv+x)});
   const jobs=rj.records.map(rc=>{const o={};jf.forEach(f=>{o[key(f)]=/Start|End|Due|Date/.test(f)&&!/__r/.test(f)?raw(rc,f):val(rc,f)});return o});
   const sitios=rs.records.map(rc=>{const o={};sf.forEach(f=>{o[key(f).replace(/\\.Name$/,'')]=val(rc,f)});return o});
   post({data:{generado:new Date().toISOString(),jobs,sitios}});
  }catch(e){post({error:String(e&&e.message||e)})}})();`;
  const s=document.createElement('script');s.src=URL.createObjectURL(new Blob([code],{type:'text/javascript'}));(document.head||document.documentElement).appendChild(s);
  setTimeout(()=>send({error:'timeout-st-pagina'}),100000);
}
})();
