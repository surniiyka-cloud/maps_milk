/* TIAN data-entry prototype. All values stay in this browser; no remote sync. */
(function(){
"use strict";
const STORE="tian-metrics-draft-v1";
const $=id=>document.getElementById(id);
const num=v=>{const n=Number(String(v??"").replace(/\s|\u00a0/g,"").replace(",","."));return Number.isFinite(n)&&n>=0?n:NaN};
const esc=v=>String(v??"").replace(/[&<>"']/g,ch=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[ch]));
const key=(city,period)=>city+"|"+period;
const form=$("entryForm"), period=$("entryPeriod"), msg=$("entryMessage"), paste=$("bulkPaste"), preview=$("pastePreview"), history=$("savedHistory");
let parsed=[], all={};
function loadStore(){try{const o=JSON.parse(localStorage.getItem(STORE)||"{}");return o&&typeof o==="object"&&!Array.isArray(o)?o:{}}catch{return {}}}
all=loadStore();
function getSaved(city=currentOffice,p=period.value){return all[key(city,p)]||null}
function setNotice(message,error=false){msg.textContent=message;msg.style.color=error?"#cf5344":"#26864a";}
function currentPeriod(){return period.value}
function normalizedDebtor(parts){let [name,amount,days,region,manager]=parts.map(x=>String(x??"").trim());const value=num(amount),delay=num(days||0);
if(!name||!Number.isFinite(value)||!Number.isInteger(delay))return null;
return {name,amount:value,days:delay,region:region||currentOffice,manager:manager||"Не указан",status:delay>0?"danger":"ok"}}
function parsePaste(raw){const lines=raw.split(/\r?\n/).map(s=>s.trim()).filter(Boolean);let result=[],errors=[];lines.forEach((line,i)=>{let parts=line.includes("\t")?line.split("\t"):line.split(";");let v=normalizedDebtor(parts);if(v)result.push(v);else if(i!==0||!/(контрагент|клиент|наименование|сумма)/i.test(line))errors.push(i+1)});return {result,errors}}
function showPreview(){const p=parsePaste(paste.value);parsed=p.result;preview.innerHTML='<div class="entry-preview-head"><strong>Распознано строк: '+p.result.length+'</strong><span>'+ (p.errors.length?'Ошибки в строках: '+p.errors.join(", "):'Ошибок нет')+'</span></div>'+ (p.result.length?'<div class="entry-preview-list">'+p.result.slice(0,6).map(d=>'<div><b>'+esc(d.name)+'</b><span>'+money(d.amount)+' · '+d.days+' дн.</span></div>').join('')+(p.result.length>6?'<small>И ещё '+(p.result.length-6)+' записей</small>':'')+'</div>':'<p class="help">Вставьте данные из 1С или Excel, разделённые табуляцией либо точкой с запятой.</p>');return p}
function formValues(){const fields=["revenue","revenuePlan","debt","orders","clients"];let result={};for(const field of fields){const el=form.elements.namedItem(field);if(!el||el.value.trim()==="")throw Error("Заполните поле «"+(el?.dataset.label||field)+"»");let n=num(el.value);if(!Number.isFinite(n))throw Error("В поле «"+el.dataset.label+"» должно быть неотрицательное число");if(["orders","clients"].includes(field)&&!Number.isInteger(n))throw Error("Количество должно быть целым");result[field]=n}return result}
function applySaved(){const saved=getSaved();if(saved){data={...sample,...saved.data,debtors:saved.data.debtors||[]};render();}else if(currentOffice==="Киров"){data=structuredClone(sample);render()}else{data={revenue:0,revenuePlan:1,debt:0,orders:0,clients:0,debtors:[]};render()}}
function fillForm(){const saved=getSaved();for(const field of ["revenue","revenuePlan","debt","orders","clients"]){const el=form.elements.namedItem(field);el.value=saved?.data?.[field]??""}paste.value="";parsed=[];showPreview();$("entryOffice").textContent=currentOffice;$("entryRecordState").textContent=saved?"Локально сохранено: "+new Date(saved.savedAt).toLocaleString("ru-RU"):"Запись за этот период ещё не создана";setNotice("");renderHistory()}
function renderHistory(){const rows=Object.entries(all).filter(([k])=>k.startsWith(currentOffice+"|")).sort(([a],[b])=>b.localeCompare(a));history.innerHTML=rows.length?rows.map(([k,v])=>'<div class="entry-history-row"><div><strong>'+esc(k.split("|")[1])+'</strong><small>'+new Date(v.savedAt).toLocaleString("ru-RU")+' · '+(v.data.debtors?.length||0)+' контрагентов</small></div><button type="button" class="ghost" data-load-period="'+esc(k.split("|")[1])+'">Открыть</button></div>').join(""):'<p class="help">Пока нет сохранённых периодов в этом браузере.</p>'}
function save(){try{const values=formValues();if(values.revenuePlan<=0)throw Error("План выручки должен быть больше нуля");const p=showPreview();if(p.errors.length)throw Error("Исправьте строки с ошибками во вставке: "+p.errors.join(", "));if(p.result.length>500)throw Error("Для демо предусмотрено не более 500 строк за раз");
const old=getSaved();const debtors=paste.value.trim()?p.result:(old?.data.debtors||[]);
if(paste.value.trim()&&debtors.length===0)throw Error("Не найдено ни одного контрагента");
const record={savedAt:new Date().toISOString(),data:{...values,debtors}};const next={...all,[key(currentOffice,currentPeriod())]:record};localStorage.setItem(STORE,JSON.stringify(next));all=next;applySaved();fillForm();setNotice("✓ Сохранено в этом браузере. Для других сотрудников эти данные пока не доступны.");updateOffice();}catch(e){setNotice(e.name==="QuotaExceededError"?"Недостаточно места в браузере. Выгрузите резервную копию.":e.message,true)}}
function downloadBackup(){const obj={version:1,createdAt:new Date().toISOString(),office:currentOffice,records:Object.fromEntries(Object.entries(all).filter(([k])=>k.startsWith(currentOffice+"|")))};const b=new Blob([JSON.stringify(obj,null,2)],{type:"application/json"}),u=URL.createObjectURL(b),a=document.createElement("a");a.href=u;a.download="tian-"+currentOffice.toLowerCase().replace(/\s/g,"-")+"-backup.json";a.click();setTimeout(()=>URL.revokeObjectURL(u),200)}
function showOffice(){fillForm();applySaved();const hasData=!!getSaved();const isKirov=currentOffice==="Киров";const originalEmpty=!isKirov&&!hasData;$("officeEmpty").hidden=!originalEmpty||section==="data-entry"||section==="cabinet"||section==="sources";document.querySelectorAll(".view").forEach(v=>{const active=v.id==="view-"+section;v.style.display=(originalEmpty&&section!=="data-entry"&&section!=="cabinet"&&section!=="sources")?"none":"";v.classList.toggle("active",active)});if(!isKirov&&hasData){$("entryDynamicNotice").hidden=false;$("entryDynamicNotice").textContent="Для "+currentOffice+" показатели загружены из локального черновика. Графики динамики и маркетинговая воронка пока демонстрационные, не рассчитываются из этих данных."}else $("entryDynamicNotice").hidden=true;}
$("entrySave").addEventListener("click",save);
$("entryPreview").addEventListener("click",showPreview);
paste.addEventListener("input",()=>{parsed=[];preview.innerHTML='<span class="help">Нажмите «Проверить вставку», чтобы увидеть предварительный результат.</span>'});
period.addEventListener("change",()=>{fillForm();applySaved()});
$("entryExport").addEventListener("click",downloadBackup);
$("entryClear").addEventListener("click",()=>{if(!getSaved())return;if(!confirm("Удалить локальный черновик за "+period.value+" для "+currentOffice+"?"))return;delete all[key(currentOffice,currentPeriod())];localStorage.setItem(STORE,JSON.stringify(all));fillForm();applySaved();updateOffice();setNotice("Локальный черновик удалён.")});
history.addEventListener("click",e=>{const b=e.target.closest("[data-load-period]");if(!b)return;period.value=b.dataset.loadPeriod;fillForm();applySaved();window.scrollTo({top:0,behavior:"smooth"})});
$("entryTemplate").addEventListener("click",()=>{paste.value="Контрагент\tСумма ДЗ\tДней просрочки\tТерритория\tОтветственный\nООО «Пример»\t125000\t12\t"+currentOffice+"\tМенеджер";showPreview()});
window.dataEntryOfficeChanged=showOffice;
window.dataEntryHasSaved=city=>Object.keys(all).some(k=>k.startsWith(city+"|"));
function showPeriodDataWhenOfficeChanged(){const saved=getSaved();if(saved){data={...sample,...saved.data,debtors:saved.data.debtors||[]};render()}else if(currentOffice==="Киров"){data=structuredClone(sample);render()}else{data={revenue:0,revenuePlan:1,debt:0,orders:0,clients:0,debtors:[]};render()}}
window.dataEntryLoadOffice=showPeriodDataWhenOfficeChanged;
showOffice();
})();