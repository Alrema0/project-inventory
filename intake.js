import { labels, censusStatuses, cleanProject, projectErrors } from './census-model.js';
import { importWorkbook, templateColumns } from './excel-import.js';
const q=s=>document.querySelector(s);
const esc=value=>String(value??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const fmt=value=>new Intl.NumberFormat('ar-SA',{maximumFractionDigits:2}).format(value);
let drafts=[], submissions=[], activeReference=null, activeSubmission=null, busy=false, loaded=false;
let importing=false;
const contactKeys=['departmentName','phone','email'];
const newDraft=()=>({uid:crypto.randomUUID(),...Object.fromEntries(Object.keys(labels).map(k=>[k,''])),status:'لم يبدأ',completionPercent:'0'});
/* ضع رابط Apps Script Web App هنا */
const SHEET_URL='https://script.google.com/macros/s/AKfycby7Q7QN4p-NkLffkCVVQ5SOKyqvQ6VRWexfvTf546pHxurC3inryqrXXSxTGwqrK4x6/exec';
const STORE_KEY='nwc-census-submissions';
function readStore(){try{return JSON.parse(localStorage.getItem(STORE_KEY)||'[]');}catch{return [];}}
function writeStore(list){try{localStorage.setItem(STORE_KEY,JSON.stringify(list.slice(0,20)));}catch{}}
function saveLocal(s){writeStore([s,...readStore().filter(r=>r.reference!==s.reference)]);}
async function sendToSheet(s){
  const rows=s.records.map((p,i)=>{
    const row={'الرقم المرجعي':s.reference,'رقم المشروع في الإرسال':i+1,'تاريخ الإرسال':s.submittedAt};
    for(const [k,v] of Object.entries(p)) row[labels[k]||k]=v;
    row['الحالة المحسوبة']=typeof state==='function'?state(reportProject(p,s.reference,i,s.submittedAt)):p.status;
    return row;
  });
  const response=await fetch(SHEET_URL,{method:'POST',body:JSON.stringify({reference:s.reference,rows}),signal:AbortSignal.timeout(35000)});
  const data=await response.json();
  if(!data.ok) throw Error(data.error||'تعذر الحفظ في Google Sheets.');
  return data;
}
function showError(selector,message){q(selector).hidden=!message;q(selector).textContent=message;}
function reportProject(p,reference,index,submittedAt='') {
  return {businessUnit:'',location:'',contractNumber:'',contractor:'',consultant:'',projectGoal:p.description||'',notes:'',...cleanProject(p),id:`${reference}-${index}`,projectNumber:`${reference.slice(0,8)}-${index+1}`,reference,submittedAt,delayOverride:'auto'};
}
function syncReports(){
  projects=submissions.flatMap(s=>s.records.map((p,i)=>reportProject(p,s.reference,i,s.submittedAt)).filter((_,i)=>!(s.deleted||[]).includes(i)));
  if(!submissions.some(s=>s.reference===activeReference)) projects.push(...drafts.filter(p=>p.projectName.trim()).map((p,i)=>reportProject(p,'مسودة',i)));
  render();
  const total=drafts.reduce((sum,p)=>sum+(Number.isFinite(Number(p.contractValue))?Number(p.contractValue):0),0);
  q('#quick-summary').innerHTML=`<div class="quick-summary"><strong>${fmt(drafts.length)} مشروع <span>·</span> ${fmt(total)} ر.س</strong><small>ملخص قبل الإرسال</small></div>`;
}
window.deleteProject=async(ref,idx)=>{
  if(busy||importing)return;
  if(ref==='مسودة'){
    const named=drafts.filter(p=>p.projectName.trim());const target=named[idx];if(!target)return;
    if(!confirm(`حذف المشروع «${target.projectName}» من النموذج؟`))return;
    drafts=drafts.filter(p=>p!==target);if(!drafts.length)drafts=[newDraft()];
    activeReference=null;drawDrafts();return;
  }
  const s=submissions.find(r=>r.reference===ref);if(!s||!s.records[idx])return;
  const sent=s.dispatched.includes(idx);
  if(!confirm(`حذف المشروع «${s.records[idx].projectName}»؟`+(sent?'\nسيُحذف أيضاً من Google Sheets.':'')))return;
  try{
    if(sent&&SHEET_URL){
      const response=await fetch(SHEET_URL,{method:'POST',body:JSON.stringify({action:'delete',reference:ref,index:idx+1}),signal:AbortSignal.timeout(35000)});
      const data=await response.json();if(!data.ok)throw Error(data.error);
    }
    const updated={...s,deleted:[...(s.deleted||[]),idx]};
    if(activeSubmission?.reference===ref)activeSubmission=updated;
    saveLocal(updated);updateSaved(updated);
  }catch{alert('تعذّر حذف المشروع من Google Sheets. يرجى المحاولة مرة أخرى.');}
};
function updateSaved(s){submissions=[s,...submissions.filter(r=>r.reference!==s.reference)];syncReports();drawPending();}
const pending=document.createElement('div');pending.id='session-pending';q('#session-error').after(pending);
function drawPending(){
  pending.replaceChildren();
  for(const s of submissions.filter(s=>s.dispatched.length<s.records.length)) {
    const box=document.createElement('div');box.className='pending-notice';
    box.innerHTML=`<strong>توجد مشاريع محفوظة لم يكتمل حفظها في Google Sheets.</strong><br>المرجع: <b dir="ltr">${esc(s.reference)}</b> — الطلبات المنفذة: ${fmt(s.dispatched.length)} من ${fmt(s.records.length)}. <br>تعاد محاولة حفظ المشاريع المتبقية باستخدام معرّف مستقل لكل مشروع.`;
    const button=document.createElement('button');button.className='button secondary';button.textContent='مراجعة واستكمال الإرسال';button.onclick=()=>resume(s);box.append(button);pending.append(box);
  }
}
async function loadSession(){
  showError('#session-error','');
  try{submissions=readStore();loaded=true;syncReports();drawPending();}
  catch(e){loaded=false;const el=q('#session-error');el.hidden=false;el.textContent='تعذر تحميل بيانات الجلسة. '+e.message+' ';const retry=document.createElement('button');retry.className='button secondary';retry.textContent='إعادة المحاولة';retry.onclick=loadSession;el.append(retry);}
}
async function loadSheetStatus(){
  const notice=q('#sheet-connection-status');
  try{
    const configured=!!SHEET_URL;
    notice.hidden=configured;
    notice.textContent=configured?'':'الحفظ في Google Sheets بانتظار تفعيل الربط. تُحفظ البيانات على هذا الجهاز عند الإرسال، ولم تُرسل إلى الجدول بعد.';
    return configured;
  }catch{
    notice.hidden=false;notice.textContent='تعذر التحقق من إعداد الحفظ في Google Sheets. يمكن إعادة المحاولة دون فقد البيانات.';
    throw Error('تعذر التحقق من إعداد الحفظ في Google Sheets.');
  }
}
function drawDrafts(){
  q('#send-projects').textContent='إرسال';
  const field=(key,value,prefix)=>{
    const id=`${prefix}-${key}`,label=labels[key];let input;
    if(key==='status') {
      const items=censusStatuses;
      input=`<select id="${id}" name="${key}" required><option value="">اختر من القائمة</option>${value&&!items.includes(value)?`<option selected value="${esc(value)}">${esc(value)} — يرجى مراجعة القيمة</option>`:''}${items.map(s=>`<option ${s===value?'selected':''}>${esc(s)}</option>`).join('')}</select>`;
    }else if(['description','notes'].includes(key)) input=`<textarea id="${id}" name="${key}" maxlength="4000" rows="2" ${key==='notes'?'':'required'}>${esc(value)}</textarea>`;
    else {
      const type=key==='email'?'email':key==='phone'?'tel':key.endsWith('Date')?'date':['contractValue','completionPercent'].includes(key)?'number':'text';
      const autocomplete=({departmentName:'organization',projectManager:'name',phone:'tel',email:'email'})[key];
      input=`<input id="${id}" name="${key}" type="${type}" value="${esc(value)}" required maxlength="4000" ${autocomplete?`autocomplete="${autocomplete}"`:''} ${['email','phone'].includes(key)?'dir="ltr"':''} ${type==='number'?`min="${key==='contractValue'?'0.01':'0'}" max="${key==='contractValue'?'1000000000000':'100'}" step="${key==='contractValue'?'0.01':'1'}"`:''} ${key==='departmentName'?'placeholder="اسم إدارتك"':''}>`;
    }
    return `<label class="${['description','notes','projectName'].includes(key)?'wide':''}" for="${id}">${label}${key==='notes'?' (اختياري)':''}${input}</label>`;
  };
  q('#contact-fields').innerHTML=`<section class="panel census-card contact-card"><div class="census-card-head"><div><h2>بيانات الإدارة والتواصل</h2><p>بيانات مشتركة. تظهر بيانات المشاريع المختلفة في أقسامها أدناه.</p></div></div><div class="census-grid">${contactKeys.map(k=>field(k,drafts[0]?.[k]||'','contact')).join('')}</div></section>`;
  q('#contact-fields').querySelectorAll('input').forEach(el=>el.oninput=()=>{
    const previous=drafts[0][el.name];drafts.forEach((p,i)=>{if(i===0||p[el.name]===previous)p[el.name]=el.value;});q('#draft-projects').querySelectorAll(`[name="${el.name}"]`).forEach(input=>{input.value=drafts.find(p=>p.uid===input.closest('[data-draft]').dataset.draft)[el.name];});changed();
  });
  q('#draft-projects').innerHTML=drafts.map((p,i)=>`<section class="panel census-card" data-draft="${p.uid}"><div class="census-card-head"><h2>المشروع ${fmt(i+1)}</h2><button type="button" class="button secondary" data-remove="${p.uid}" ${drafts.length===1?'hidden':''}>حذف</button></div><div class="project-groups">${contactKeys.some(k=>p[k]!==drafts[0][k])?`<section class="project-field-group project-contact"><div class="project-group-heading"><div><h3>بيانات الإدارة والتواصل لهذا المشروع</h3><p>بيانات مستقلة كما وردت في الملف المرفق.</p></div></div><div class="census-grid group-fields">${contactKeys.map(k=>field(k,p[k]??'',p.uid)).join('')}</div></section>`:''}${[
    ['٠١','البيانات الأساسية',['projectName','projectManager','description'],'اسم المشروع والمسؤول عنه والوصف المختصر.'],
    ['٠٢','الجدول الزمني والقيمة',['startDate','plannedEndDate','contractValue'],'تواريخ المشروع وقيمته التعاقدية بالريال السعودي.'],
    ['٠٣','الحالة ونسبة الإنجاز',['status','completionPercent','notes'],'الحالة الحالية ونسبة الإنجاز والملاحظات أو المعوقات.']
  ].map(([number,title,keys,description])=>`<section class="project-field-group"><div class="project-group-heading"><span>${number}</span><div><h3>${title}</h3><p>${description}</p></div></div><div class="census-grid group-fields">${keys.map(k=>field(k,p[k]??'',p.uid)).join('')}</div></section>`).join('')}</div></section>`).join('');
  q('#draft-projects').querySelectorAll('input,textarea,select').forEach(el=>el.addEventListener('input',()=>{
    const draft=drafts.find(p=>p.uid===el.closest('[data-draft]').dataset.draft);draft[el.name]=el.value;
    if(el.name==='status'&&['مكتمل','لم يبدأ'].includes(el.value)){
      draft.completionPercent=el.value==='مكتمل'?100:0;
      el.closest('[data-draft]').querySelector('[name="completionPercent"]').value=draft.completionPercent;
    }
    changed();
  }));
  q('#draft-projects').querySelectorAll('[data-remove]').forEach(b=>b.onclick=()=>{drafts=drafts.filter(p=>p.uid!==b.dataset.remove);activeReference=null;drawDrafts();});
  syncReports();
}
function changed(){activeReference=null;showError('#census-errors','');showError('#send-error','');q('#send-projects').textContent='إرسال';syncReports();}
q('#another-project').onclick=()=>{
  const last=drafts.at(-1),draft=newDraft();
  for(const key of contactKeys) draft[key]=last?.[key]||'';
  drafts.push(draft);activeReference=null;drawDrafts();q(`[data-draft="${draft.uid}"] input`)?.focus();
};
q('#census-form').onsubmit=e=>{e.preventDefault();q('#send-projects').click();};
function resume(s){
  if(busy||importing)return;
  activeSubmission=s;activeReference=s.reference;
  drafts=s.records.map(p=>({...p,uid:crypto.randomUUID()}));drawDrafts();
  q('#census-form').querySelectorAll('input,select,textarea,button').forEach(el=>el.disabled=true);
  q('#send-projects').disabled=false;q('#send-projects').hidden=false;
  q('#send-projects').textContent='إرسال المتبقي';q('#success-view').hidden=true;
  q('#send-progress').textContent=`طلبات الإرسال المسجّلة: ${fmt(s.dispatched.length)} من ${fmt(s.records.length)}.`;
  q('#submission-reference').value=s.reference;q('#submission-date').value=s.submittedAt;
  showError('#send-error','');q('#submission-controls').scrollIntoView({block:'center'});
}
q('#new-census').onclick=()=>{
  activeSubmission=null;activeReference=null;drafts=[newDraft()];
  q('#success-view').hidden=true;q('#send-projects').hidden=false;
  q('#send-progress').textContent='';showError('#send-error','');
  q('#submission-reference').value='';q('#submission-date').value='';
  q('#census-form').querySelectorAll('input,select,textarea,button').forEach(el=>el.disabled=false);
  drawDrafts();q('#contact-fields input')?.focus();
};
q('#send-projects').onclick=async e=>{
  e.preventDefault();
  if(busy||importing)return;
  if(!activeSubmission){
    if(!q('#census-form').reportValidity())return;
    const errors=drafts.flatMap((p,i)=>projectErrors(p).map(error=>`المشروع ${fmt(i+1)}: ${error}`));
    if(errors.length){showError('#census-errors',errors.join('\n'));q('#census-errors').scrollIntoView({block:'center'});return;}
    activeReference??=crypto.randomUUID();q('#submission-reference').value=activeReference;
  }
  busy=true;const button=q('#send-projects');button.disabled=true;
  q('#census-form').querySelectorAll('input,select,textarea,button').forEach(el=>el.disabled=true);
  showError('#send-error','');
  const records=activeSubmission?.records||drafts.map(cleanProject);
  try{
    if(!loaded){await loadSession();if(!loaded)throw Error('يلزم استعادة الاتصال ببيانات الجلسة أولاً.');}
    button.textContent='جارٍ حفظ البيانات…';
    if(!activeSubmission) {
      activeSubmission={reference:activeReference,submittedAt:new Date().toISOString(),records,dispatched:[]};
      saveLocal(activeSubmission);updateSaved(activeSubmission);
    }
    q('#submission-date').value=activeSubmission.submittedAt;
    if(!await loadSheetStatus())throw Error('الربط مع Google Sheets بانتظار التفعيل. بياناتك محفوظة على هذا الجهاز ولم تُرسل إلى الجدول بعد.');
    if(activeSubmission.dispatched.length<records.length){
      button.textContent='جارٍ الحفظ في Google Sheets…';
      q('#send-progress').textContent=`جارٍ حفظ ${fmt(records.length)} مشروع في Google Sheets.`;
      await sendToSheet(activeSubmission);
      activeSubmission={...activeSubmission,dispatched:records.map((_,i)=>i)};
      saveLocal(activeSubmission);updateSaved(activeSubmission);
    }
    q('#success-reference').textContent=activeSubmission.reference;
    syncReports();q('#success-view').hidden=false;button.hidden=true;
    q('#send-progress').textContent='';q('#success-view h2').focus();
  } catch(e){
    showError('#send-error',(e.name==='TimeoutError'?'انتهت مهلة الاتصال.':e.message)+'\nلم يكتمل الحفظ في Google Sheets. يمكنك إعادة المحاولة لاستكمال المشاريع المتبقية.');
    q('#send-progress').textContent=activeSubmission?`حُفظت البيانات بالمرجع ${activeSubmission.reference}. طلبات الإرسال المسجّلة: ${activeSubmission.dispatched.length} من ${records.length}.`:'لم تُحفظ المشاريع في Google Sheets.';
    button.textContent='إعادة المحاولة';
  }finally{busy=false;button.disabled=false;q('#census-form').querySelectorAll('input,select,textarea,button').forEach(el=>el.disabled=!!activeSubmission);button.disabled=false;q('#excel-file').disabled=!!activeSubmission;q('#download-template').disabled=false;q('#new-census').disabled=false;}
};
let excelLoader;
function loadExcel(){
  if(window.ExcelJS)return Promise.resolve(window.ExcelJS);
  if(!excelLoader)excelLoader=new Promise((resolve,reject)=>{
    const script=document.createElement('script');script.src='./vendor/exceljs.min.js';script.onload=()=>window.ExcelJS?resolve(window.ExcelJS):reject(Error('تعذر تشغيل قارئ Excel.'));script.onerror=()=>{script.remove();excelLoader=null;reject(Error('تعذر تحميل قارئ Excel؛ يرجى إعادة المحاولة.'));};document.head.append(script);
  });
  return excelLoader;
}
function toggleImportLock(locked){
  importing=locked;q('#census-form').querySelectorAll('input,select,textarea,button').forEach(el=>el.disabled=locked||!!activeSubmission);
  q('#download-template').disabled=locked;
}
q('#excel-file').onchange=async e=>{
  const file=e.target.files?.[0];if(!file)return;
  if(busy||importing||activeSubmission){e.target.value='';return;}
  showError('#import-error','');q('#import-result').hidden=true;
  if(!/\.xlsx$/i.test(file.name)){showError('#import-error','يرجى إرفاق ملف Excel بصيغة .xlsx. يمكن حفظ الملفات القديمة بهذه الصيغة من Excel.');e.target.value='';return;}
  toggleImportLock(true);q('#import-progress').hidden=false;q('#import-progress').textContent='جارٍ قراءة الملف وتوزيع بيانات المشاريع…';
  try{
    const Excel=await loadExcel();const workbook=new Excel.Workbook();await workbook.xlsx.load(await file.arrayBuffer());
    const result=importWorkbook(workbook,drafts[0]||{});
    const existing=drafts.filter(p=>Object.keys(labels).some(k=>!contactKeys.includes(k)&&!['status','completionPercent'].includes(k)&&String(p[k]??'').trim())||p.status!=='لم يبدأ'||Number(p.completionPercent)!==0);
    drafts=[...existing,...result.records.map(p=>({...p,uid:crypto.randomUUID()}))];activeReference=null;drawDrafts();
    const box=q('#import-result');box.replaceChildren();box.hidden=false;
    const title=document.createElement('p');title.textContent=`تمت قراءة «${file.name}» وإضافة ${fmt(result.records.length)} مشروعاً إلى النموذج وجميع أقسامه. يرجى المراجعة قبل الإرسال.`;box.append(title);
    if(result.warnings.length){
      const details=document.createElement('details');details.open=true;const summary=document.createElement('summary');summary.textContent=`${fmt(result.warnings.length)} صفاً يحتاج إلى استكمال أو تصحيح`;
      const list=document.createElement('ul');for(const w of result.warnings){const li=document.createElement('li');li.textContent=`${w.sheet} — الصف ${fmt(w.row)}: ${w.errors.join(' ')}`;list.append(li);}details.append(summary,list);box.append(details);
    }
    showError('#census-errors','');showError('#send-error','');
  }catch(error){showError('#import-error',error.message?.includes('الورقة')||error.message?.includes('يرجى')||error.message?.includes('الملف')||error.message?.includes('مشروع')?error.message:'تعذرت قراءة الملف. تأكد من أنه ملف Excel سليم وغير محمي بكلمة مرور. لم تتغير البيانات الحالية.');}
  finally{toggleImportLock(false);q('#import-progress').hidden=true;e.target.value='';}
};
q('#download-template').onclick=async()=>{
  if(importing||busy)return;
  toggleImportLock(true);showError('#import-error','');
  try{
    const Excel=await loadExcel(),workbook=new Excel.Workbook(),sheet=workbook.addWorksheet('المشاريع',{views:[{rightToLeft:true,state:'frozen',ySplit:1}]});
    sheet.columns=templateColumns.map(([key,header])=>({key,header,width:key==='description'||key==='notes'?40:26}));
    sheet.getRow(1).eachCell(cell=>{cell.font={bold:true,color:{argb:'FFFFFFFF'}};cell.fill={type:'pattern',pattern:'solid',fgColor:{argb:'FF1976B8'}};cell.alignment={horizontal:'right',wrapText:true};});sheet.getRow(1).height=34;sheet.getColumn('phone').numFmt='@';
    sheet.getCell(1,templateColumns.findIndex(([key])=>key==='startDate')+1).note='التاريخ بصيغة YYYY-MM-DD أو خلية تاريخ Excel.';sheet.getCell(1,templateColumns.findIndex(([key])=>key==='plannedEndDate')+1).note='التاريخ بصيغة YYYY-MM-DD أو خلية تاريخ Excel.';sheet.getCell(1,templateColumns.findIndex(([key])=>key==='completionPercent')+1).note='نسبة الإنجاز من 0 إلى 100، أو خلية منسقة كنسبة مئوية.';
    const bytes=await workbook.xlsx.writeBuffer(),url=URL.createObjectURL(new Blob([bytes],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'}));
    const link=document.createElement('a');link.href=url;link.download='قالب-مشاريع-الإدارات.xlsx';link.click();setTimeout(()=>URL.revokeObjectURL(url),1000);
  }catch{showError('#import-error','تعذر تنزيل القالب. يرجى إعادة المحاولة.');}
  finally{toggleImportLock(false);}
};
drafts=[newDraft()];drawDrafts();loadSession();loadSheetStatus().catch(()=>{});