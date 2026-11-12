/* Staff directory demo: local only, vacation schedule dates are discarded. */
(function(){
"use strict";
const $=id=>document.getElementById(id),STORE="tian-kirov-staff-v1";
const defaultStaff=[
{name:"Захаров Евгений Александрович",position:"Руководитель обособленного подразделения"},
{name:"Сурнин Данил Дмитриевич",position:"Руководитель направления маркетплейсы"}
];
const safe=x=>String(x??"").replace(/[&<>"']/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c]));
const norm=x=>String(x||"").toLowerCase().replace(/ё/g,"е").replace(/\s+/g," ").trim();
const validName=x=>/^[А-ЯЁ][а-яё-]+(?:\s+[А-ЯЁ][а-яё-]+){1,2}$/.test(String(x??"").trim());
let custom=[],proposed=[];
try{const stored=JSON.parse(localStorage.getItem(STORE)||"[]");if(Array.isArray(stored))custom=stored.filter(x=>x&&validName(x.name)&&typeof x.position==="string").slice(0,250)}catch{}
function list(){const names=new Set(defaultStaff.map(x=>norm(x.name)));return [...defaultStaff,...custom.filter(x=>{const k=norm(x.name);if(names.has(k))return false;names.add(k);return true})]}
function render(){const q=norm($("staffSearch").value),all=list(),rows=all.filter(x=>norm(x.name+" "+x.position).includes(q));$("staffCount").textContent=all.length+" сотрудн.";const out=rows.map(x=>'<div class="staff-member"><span class="staff-initial">'+safe(x.name[0])+'</span><div><b>'+safe(x.name)+'</b><small>'+safe(x.position||"Должность уточняется")+'</small></div></div>').join("");$("staffRoster").innerHTML=out||'<p class="help">Поиск не дал результатов.</p>'}
function setMsg(m,err){$("staffMessage").textContent=m;$("staffMessage").style.color=err?"#b84c3c":"#288a55"}
function parse(wb){let res=[];for(const sheetName of wb.SheetNames){const ws=wb.Sheets[sheetName];const rows=XLSX.utils.sheet_to_json(ws,{header:1,defval:"",raw:false,blankrows:false});for(let r=0;r<rows.length;r++){const row=rows[r];if(!Array.isArray(row))continue;for(let c=0;c<Math.min(row.length,24);c++){let value=String(row[c]??"").trim().replace(/\s+/g," ");if(!validName(value))continue;const before=rows[r-1]||[],header=String(before[c]||"").toLowerCase();let position="";for(const off of [-2,-1,1,2]){const v=String(row[c+off]??"").trim();if(!v||v.length>100||validName(v)||/\d{2}[./-]\d{2}/.test(v))continue;if(/^(менеджер|специалист|руководитель|бухгалтер|инженер|сотрудник|кладовщик|логист|дизайнер|директор|начальник|торговый|оператор|маркетолог)/i.test(v)){position=v;break}}if(/(месяц|дата|отпуск|год)/.test(header))continue;res.push({name:value,position:position||"Должность уточняется",sheet:sheetName})}}}
const seen=new Set();return res.filter(x=>{let k=norm(x.name);if(seen.has(k))return false;seen.add(k);return true}).slice(0,250)}
async function loadFile(file){if(!file)return;if(!/\.xlsx?$/i.test(file.name)||file.size>15*1024*1024)return setMsg("Нужен Excel .xlsx/.xls размером до 15 МБ.",true);if(!window.XLSX)return setMsg("Библиотека Excel не загрузилась. Проверьте подключение к интернету.",true);
try{let wb=XLSX.read(await file.arrayBuffer(),{type:"array"});proposed=parse(wb);$("staffPreview").hidden=false;$("staffFileNote").textContent=file.name+" · найдено кандидатов: "+proposed.length;$("staffCandidateList").innerHTML=proposed.length?proposed.map((x,i)=>'<div class="staff-candidate"><label><input type="checkbox" data-staff-index="'+i+'" '+(list().some(y=>norm(y.name)===norm(x.name))?'disabled':'checked')+'><span>'+safe(x.name)+'<small>'+safe(x.position)+' · '+safe(x.sheet)+'</small></span></label></div>').join(""):'<p class="help">Не удалось уверенно найти ФИО. Проверьте, что в файле ФИО написаны в отдельных ячейках в формате «Фамилия Имя Отчество».</p>';setMsg("Проверьте кандидатов: из графика отпусков автоматически нельзя гарантировать актуальность штатного состава.",false)}catch(e){setMsg("Ошибка чтения файла: "+e.message,true)}}
$("staffSearch").addEventListener("input",render);
$("staffImportTrigger").addEventListener("click",()=>$("staffFile").click());
$("staffFile").addEventListener("change",e=>{loadFile(e.target.files?.[0]);e.target.value=""});
$("staffCancel").addEventListener("click",()=>{$("staffPreview").hidden=true;proposed=[]});
$("staffApply").addEventListener("click",()=>{const selected=[...document.querySelectorAll("[data-staff-index]:checked")].map(x=>proposed[Number(x.dataset.staffIndex)]).filter(Boolean);if(!selected.length)return setMsg("Не выбрано новых сотрудников.",true);let seen=new Set(list().map(x=>norm(x.name)));let added=[];for(const s of selected){const k=norm(s.name);if(!seen.has(k)){added.push({name:s.name,position:s.position});seen.add(k)}}try{localStorage.setItem(STORE,JSON.stringify([...custom,...added]));custom.push(...added);$("staffPreview").hidden=true;proposed=[];render();setMsg("Добавлено сотрудников: "+added.length+". Список хранится только в этом браузере.")}catch{setMsg("Не удалось сохранить список в этом браузере.",true)}});
render();
})();