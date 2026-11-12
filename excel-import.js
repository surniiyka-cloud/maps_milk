/* Browser-only Excel importer. Never sends workbook contents to a server. */
(function(){
"use strict";
const $=id=>document.getElementById(id), STORE="tian-metrics-draft-v1";
const esc=v=>String(v??"").replace(/[&<>"']/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c]));
const moneyFmt=v=>new Intl.NumberFormat("ru-RU",{maximumFractionDigits:2}).format(v);
const number=v=>{if(typeof v==="number")return Number.isFinite(v)?v:null;let s=String(v??"").replace(/[\s\u00a0₽]/g,"").replace(/,/g,".");if(!s||!/^-?\d+(\.\d+)?$/.test(s))return null;let n=Number(s);return Number.isFinite(n)?n:null};
const normalize=s=>String(s??"").toLowerCase().replace(/ё/g,"е").replace(/[\s\u00a0_.:—–-]+/g," ").trim();
const fields={revenue:{title:"Выручка",terms:[/\bвыручк/,/\bпродаж[аи]\b/,/\bтоварооборот/]},revenuePlan:{title:"План выручки",terms:[/план.*(выруч|продаж|оборот)/,/(выруч|продаж|оборот).*план/]},debt:{title:"Дебиторская задолженность",terms:[/дебиторск/,/\bдз\b/,/задолженность покупател/]},orders:{title:"Количество заказов",terms:[/количество заказ/,/\bзаказов\b/,/число заказ/]},clients:{title:"Активные клиенты",terms:[/активн.*клиент/,/количество клиент/,/\bклиентов\b/]}};
let parsed=null, selected={}, fileInfo=null, store={};
function readStore(){try{const o=JSON.parse(localStorage.getItem(STORE)||"{}");return o&&typeof o==="object"&&!Array.isArray(o)?o:{}}catch{return {}}}
store=readStore();
const recordKey=(city,period)=>city+"|"+period;
const saved=(city=currentOffice,period=$("importPeriod").value)=>store[recordKey(city,period)]||null;
function msg(text,isError=false){$("importMessage").textContent=text;$("importMessage").style.color=isError?"#c75140":"#238c4f"}
function scan(sheet, sheetName){
let matches={},debtors=[],tableCandidates=[];
const rows=sheet;
for(let r=0;r<Math.min(rows.length,1500);r++){
const row=rows[r]||[];
for(let c=0;c<Math.min(row.length,35);c++){
const label=normalize(row[c]);if(!label||number(row[c])!==null||label.length>100)continue;
for(const [field,def] of Object.entries(fields)){
if(!def.terms.some(re=>re.test(label)))continue;
if(field==="revenue"&&/(план|отклонен|процент|доля|рост|динамик)/.test(label))continue;
if(field==="debt"&&/(погашен|просроч|доля|срок)/.test(label))continue;
if(field==="clients"&&/(новых|неактивн)/.test(label))continue;
const picks=[{v:row[c+1],at:"справа",weight:4},{v:row[c+2],at:"через ячейку",weight:2},{v:rows[r+1]?.[c],at:"ниже",weight:3},{v:rows[r+1]?.[c+1],at:"по диагонали",weight:1}];
for(const p of picks){const val=number(p.v);if(val===null||val<0||((field==="orders"||field==="clients")&&!Number.isInteger(val)))continue;
const item={field,value:val,sheet:sheetName,cell:(window.XLSX?.utils.encode_cell({r,c})||""),label:String(row[c]),placement:p.at,score:p.weight+(label.length<45?1:0)};
(matches[field]??=[]).push(item);
}
}
}
const headers=row.map(normalize);
const nameAt=headers.findIndex(h=>/^(контрагент|покупатель|клиент|наименование контрагента|наименование покупателя)$/.test(h));
const amountAt=headers.findIndex(h=>/(сумма дз|сумма задолженности|дебиторская задолженность|долг|итого задолженность|задолженность покупателя)/.test(h));
if(nameAt>=0&&amountAt>=0&&nameAt!==amountAt){
const daysAt=headers.findIndex(h=>/(дней просрочки|просрочка дней|дни просрочки)/.test(h));
const regionAt=headers.findIndex(h=>/(территория|город|регион)/.test(h));
const managerAt=headers.findIndex(h=>/(менеджер|ответственный)/.test(h));
let items=[];for(let i=r+1;i<Math.min(rows.length,r+2001);i++){const rr=rows[i]||[],name=String(rr[nameAt]??"").trim(),val=number(rr[amountAt]);if(!name||/^(итого|всего|результат)/i.test(name))continue;if(val===null||val<0)continue;
const days=daysAt>=0?number(rr[daysAt]):0;if(days!==null&&(!Number.isInteger(days)||days<0))continue;
items.push({name,amount:val,days:days??0,status:(days??0)>0?"danger":"ok",region:regionAt>=0?String(rr[regionAt]||currentOffice):currentOffice,manager:managerAt>=0?String(rr[managerAt]||"Не указан"):"Не указан"});
}
if(items.length)tableCandidates.push({sheet:sheetName,headers:r+1,items});
}
}
return {matches,tableCandidates};
}
function parseWorkbook(wb){let all={},tables=[];for(const name of wb.SheetNames){const sheet=wb.Sheets[name];const rows=XLSX.utils.sheet_to_json(sheet,{header:1,defval:"",raw:false,blankrows:false});const result=scan(rows,name);for(const [k,v] of Object.entries(result.matches))(all[k]??=[]).push(...v);tables.push(...result.tableCandidates)}
for(const key of Object.keys(all))all[key].sort((a,b)=>b.score-a.score);
tables.sort((a,b)=>b.items.length-a.items.length);
return {matches:all,tables,sheets:wb.SheetNames}}
function metricSelect(field){const options=parsed.matches[field]||[];return '<div class="import-metric"><div><b>'+fields[field].title+'</b><small>'+ (options.length?'Найдены кандидаты: '+options.length:'Автоматически не найдено')+'</small></div><select class="select" data-metric="'+field+'"><option value="">Не импортировать</option>'+options.slice(0,40).map((x,i)=>'<option value="'+i+'" '+(i===0?'selected':'')+'>'+esc(moneyFmt(x.value)+' — '+x.sheet+' ('+x.cell+')')+'</option>').join('')+'</select></div>'}
function showPreview(){const target=$("importPreview");target.hidden=false;$("importFileName").textContent=fileInfo.name;$("importSheets").textContent=parsed.sheets.join(" · ");$("importMetricList").innerHTML=Object.keys(fields).map(metricSelect).join("");
const tables=parsed.tables;$("importDebtSelect").innerHTML='<option value="">Не импортировать реестр задолженности</option>'+tables.map((t,i)=>'<option value="'+i+'">'+esc(t.sheet+' — '+t.items.length+' строк')+'</option>').join("");
if(tables.length)$("importDebtSelect").value="0";
renderDebtPreview();$("importConfirm").disabled=!Object.values(parsed.matches).some(x=>x.length)&&!tables.length;
$("importStatus").textContent="Проверьте, правильно ли сайт нашёл показатели. Неподтверждённые значения не сохраняются.";
}
function renderDebtPreview(){const idx=$("importDebtSelect").value;const table=idx===""?null:parsed.tables[Number(idx)];$("importDebtPreview").innerHTML=table?'<p class="help">Найдено '+table.items.length+' строк. Показаны первые пять:</p><div class="entry-preview-list">'+table.items.slice(0,5).map(d=>'<div><b>'+esc(d.name)+'</b><span>'+moneyFmt(d.amount)+' ₽</span></div>').join("")+'</div>':'<p class="help">Реестр контрагентов не выбран.</p>'}
async function fileChosen(file){parsed=null;$("importPreview").hidden=true;$("importConfirm").disabled=true;if(!file)return;
if(!/\.xlsx?$/i.test(file.name))return msg("Поддерживаются файлы .xlsx и .xls",true);
if(file.size>15*1024*1024)return msg("Максимальный размер файла для прототипа — 15 МБ",true);
if(!window.XLSX)return msg("Библиотека чтения Excel не загрузилась. Проверьте подключение к интернету.",true);
try{msg("Читаем книгу и ищем показатели…");const bytes=await file.arrayBuffer();const workbook=XLSX.read(bytes,{type:"array",cellDates:false});fileInfo={name:file.name,size:file.size};parsed=parseWorkbook(workbook);showPreview();msg("Файл прочитан в браузере. Сверьте найденные значения перед импортом.")}catch(e){msg("Не удалось разобрать Excel: "+e.message,true)}}
function applyForOffice(){const rec=saved();if(rec){data={...sample,...rec.data,debtors:rec.data.debtors||[]};render()}else if(currentOffice==="Киров"){data=structuredClone(sample);render()}else{data={revenue:0,revenuePlan:1,debt:0,orders:0,clients:0,debtors:[]};render()}}
function showHistory(){const list=Object.entries(store).filter(([k])=>k.startsWith(currentOffice+"|")).sort(([a],[b])=>b.localeCompare(a));$("importHistory").innerHTML=list.length?list.map(([k,rec])=>'<div class="entry-history-row"><div><strong>'+esc(k.split("|")[1])+'</strong><small>'+new Date(rec.savedAt).toLocaleString("ru-RU")+' · '+esc(rec.sourceFile||"Ранее сохранённый черновик")+'</small></div><button class="ghost" data-period="'+esc(k.split("|")[1])+'">Открыть</button></div>').join(""):'<p class="help">Для этого города нет сохранённых импортов на данном устройстве.</p>'}
function syncOffice(){applyForOffice();$("importOffice").textContent=currentOffice;showHistory();const rec=saved();$("importSavedState").textContent=rec?"Последний импорт: "+new Date(rec.savedAt).toLocaleString("ru-RU"):"Нет локального импорта за этот месяц";let empty=currentOffice!=="Киров"&&!window.dataEntryHasSaved(currentOffice);$("officeEmpty").hidden=!empty||["data-entry","cabinet","sources"].includes(section);document.querySelectorAll(".view").forEach(v=>v.style.display=empty&&!["data-entry","cabinet","sources"].includes(section)?"none":"");}
function confirm(){if(!parsed)return msg("Сначала загрузите Excel.",true);
try{const changes={},details=[];document.querySelectorAll("[data-metric]").forEach(el=>{const field=el.dataset.metric;if(el.value==="")return;const match=parsed.matches[field][Number(el.value)];if(!match)return;changes[field]=match.value;details.push(fields[field].title+" = "+moneyFmt(match.value))});
const choice=$("importDebtSelect").value;const debtors=choice===""?null:parsed.tables[Number(choice)]?.items;
if(!Object.keys(changes).length&&!debtors?.length)throw Error("Выберите хотя бы один показатель или реестр задолженности.");
const period=$("importPeriod").value,old=saved(),base=old?.data||{};
const nextData={...base,...changes};if(debtors)nextData.debtors=debtors;
const next={...store,[recordKey(currentOffice,period)]:{data:nextData,savedAt:new Date().toISOString(),sourceFile:fileInfo.name,sourceSheets:parsed.sheets}};
localStorage.setItem(STORE,JSON.stringify(next));store=next;syncOffice();updateOffice();
msg("✓ Импорт сохранён локально: "+Object.keys(changes).length+" показателей"+(debtors?" и "+debtors.length+" контрагентов":"")+". Другие сотрудники пока не увидят данные.");
$("importStatus").textContent="Сохранено. Не забудьте сделать резервную копию. Сведения не отправлялись на сервер.";
}catch(e){msg(e.name==="QuotaExceededError"?"В браузере недостаточно места. Данные не сохранены.":e.message,true)}}
function backup(){const obj={version:2,office:currentOffice,createdAt:new Date().toISOString(),records:Object.fromEntries(Object.entries(store).filter(([k])=>k.startsWith(currentOffice+"|")))};const blob=new Blob([JSON.stringify(obj,null,2)],{type:"application/json"}),url=URL.createObjectURL(blob),a=document.createElement("a");a.href=url;a.download="tian-"+currentOffice.replace(/\s/g,"-")+"-imports.json";a.click();setTimeout(()=>URL.revokeObjectURL(url),1000)}
$("excelInput").addEventListener("change",e=>{fileChosen(e.target.files?.[0]);e.target.value=""});
const zone=$("excelDrop");zone.addEventListener("click",()=>$("excelInput").click());zone.addEventListener("keydown",e=>{if(e.key==="Enter"||e.key===" "){e.preventDefault();$("excelInput").click()}});zone.addEventListener("dragover",e=>{e.preventDefault();zone.classList.add("dragging")});zone.addEventListener("dragleave",()=>zone.classList.remove("dragging"));zone.addEventListener("drop",e=>{e.preventDefault();zone.classList.remove("dragging");fileChosen(e.dataTransfer.files?.[0])});
$("importDebtSelect").addEventListener("change",renderDebtPreview);$("importConfirm").addEventListener("click",confirm);$("importBackup").addEventListener("click",backup);$("importPeriod").addEventListener("change",syncOffice);
$("importHistory").addEventListener("click",e=>{const btn=e.target.closest("[data-period]");if(btn){$("importPeriod").value=btn.dataset.period;syncOffice()}});
window.dataEntryHasSaved=city=>Object.keys(store).some(k=>k.startsWith(city+"|"));
window.dataEntryOfficeChanged=syncOffice;
syncOffice();
})();