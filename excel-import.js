import { labels, numeric, projectErrors } from './census-model.js';
export const templateColumns=Object.entries(labels);
const contactKeys=['departmentName','phone','email'];
const aliases={
  departmentName:['اسم الإدارة','الإدارة','اسم الاداره','department','departmentName'],
  projectManager:['المسؤول عن المشروع','اسم المسؤول','المسؤول','مدير المشروع','مدير المشروع / المسؤول','projectManager','manager'],
  phone:['رقم التواصل','رقم الجوال','رقم الهاتف','الجوال','phone','mobile'],
  email:['البريد الإلكتروني','البريد الالكتروني','الايميل','email'],
  projectName:['اسم المشروع','المشروع','projectName','project name'],
  description:['وصف مختصر','الوصف','وصف المشروع','الهدف من المشروع','هدف المشروع','description','projectGoal'],
  projectType:['القطاع','نوع المشروع','projectType','sector'],
  contractValue:['القيمة التعاقدية بالريال السعودي','القيمة التعاقدية','قيمة العقد','قيمة العقد (ريال)','القيمة التعاقدية (ر.س)','contractValue','budget'],
  startDate:['تاريخ البداية','تاريخ البدء','startDate','start date'],
  plannedEndDate:['تاريخ النهاية المتوقع','تاريخ النهاية','تاريخ الانتهاء المتوقع','تاريخ الانتهاء المخطط','plannedEndDate','end date'],
  status:['الحالة','حالة المشروع','status'],
  completionPercent:['نسبة الإنجاز %','نسبة الإنجاز','نسبة الانجاز','completionPercent','progress'],
  notes:['ملاحظات أو معوقات','الملاحظات','ملاحظات','المعوقات','notes']
};
const normalize=s=>String(s??'').normalize('NFKC').toLowerCase().replace(/[أإآ]/g,'ا').replace(/ى/g,'ي').replace(/ة/g,'ه').replace(/[\s\p{P}\p{S}\u064B-\u065F\u200E\u200F]/gu,'');
const headerKey=value=>Object.keys(aliases).find(k=>aliases[k].some(v=>normalize(v)===normalize(value)));
function cellValue(cell){
  const v=cell.value;
  if(v===null||v===undefined)return '';
  if(typeof v==='object'&&!(v instanceof Date)){
    if('formula' in v||'sharedFormula' in v)return v.result??'';
    if(v.richText)return v.richText.map(t=>t.text).join('');
    if('text' in v)return v.text;
    return '';
  }
  return v;
}
function dateValue(value,date1904){
  if(value instanceof Date)return Number.isNaN(value.getTime())?'':value.toISOString().slice(0,10);
  if(typeof value==='number'&&value>0&&value<2958466){
    const epoch=date1904?Date.UTC(1904,0,1):Date.UTC(1899,11,30);
    return new Date(epoch+Math.floor(value)*86400000).toISOString().slice(0,10);
  }
  const text=String(value).trim().replace(/[٠-٩]/g,c=>'٠١٢٣٤٥٦٧٨٩'.indexOf(c));
  const iso=text.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})$/);
  const local=text.match(/^(\d{1,2})[-/](\d{1,2})[-/](\d{4})$/);
  if(iso)return `${iso[1]}-${iso[2].padStart(2,'0')}-${iso[3].padStart(2,'0')}`;
  if(local)return `${local[3]}-${local[2].padStart(2,'0')}-${local[1].padStart(2,'0')}`;
  return text;
}
function readCell(cell,key,date1904){
  const value=cellValue(cell);
  if(key.endsWith('Date'))return dateValue(value,date1904);
  if(key==='phone')return String(typeof value==='number'&&cell.numFmt&&/^0+$/.test(cell.numFmt)?String(value).padStart(cell.numFmt.length,'0'):value).trim();
  if(key==='completionPercent'){
    if(value==='')return '';
    if(typeof value==='number'&&/%/.test(cell.numFmt||''))return Number((value*100).toFixed(8));
    const text=String(value).trim().replace(/[%٪]/g,'');return numeric(text);
  }
  if(key==='contractValue'){
    if(value==='')return '';
    return numeric(String(value).replace(/,/g,'').replace(/SAR|ريال سعودي|ريال|ر\.س\.?/gi,'').trim());
  }
  return String(value??'').trim();
}
function extraValue(cell,date1904){
  const v=cellValue(cell);
  if(v instanceof Date)return dateValue(v,date1904);
  if(typeof v==='number'&&/[dmy]/i.test(cell.numFmt||'')&&!/%/.test(cell.numFmt||''))return dateValue(v,date1904);
  if(typeof v==='number'&&/%/.test(cell.numFmt||''))return Number((v*100).toFixed(4))+'%';
  return String(v??'').trim();
}
function detectHeader(sheet){
  for(let i=1;i<=Math.min(sheet.rowCount,25);i++){
    const columns=new Map(),extras=new Map(),used=new Set();
    sheet.getRow(i).eachCell((cell,col)=>{
      const text=String(cellValue(cell)??'').trim();const key=headerKey(text);
      if(key){if(columns.has(key))throw Error(`الورقة «${sheet.name}»: يوجد عمود مكرر للحقل «${labels[key]}».`);columns.set(key,col);}
      else if(text){let name=text.slice(0,100),n=2;while(used.has(name))name=`${text.slice(0,95)} (${n++})`;used.add(name);extras.set(col,name);}
    });
    if(columns.has('projectName')||(columns.has('departmentName')&&columns.size>=2))return {row:i,columns,extras};
  }
  return null;
}
export function importWorkbook(workbook,defaults={}){
  const contacts=new Map(),tables=[],warnings=[],records=[],sources=[];
  for(const sheet of workbook.worksheets){
    const header=detectHeader(sheet);if(!header)continue;
    if(header.columns.has('projectName')){tables.push({sheet,...header});continue;}
    for(let r=header.row+1;r<=sheet.rowCount;r++){
      const contact={};for(const k of contactKeys)if(header.columns.has(k))contact[k]=readCell(sheet.getRow(r).getCell(header.columns.get(k)),k,false);
      if(contact.departmentName)contacts.set(contact.departmentName,contact);
    }
  }
  if(!tables.length)throw Error('لم يُعثر على عمود «اسم المشروع». يرجى استخدام القالب أو عناوين أعمدة عربية واضحة.');
  const onlyContact=contacts.size===1?[...contacts.values()][0]:{};
  for(const {sheet,row,columns,extras} of tables){
    let previous={};
    for(let r=row+1;r<=sheet.rowCount;r++){
      const excelRow=sheet.getRow(r);
      const data=Object.fromEntries([...Object.keys(labels),'projectType'].map(k=>[k,columns.has(k)?readCell(excelRow.getCell(columns.get(k)),k,workbook.properties.date1904):'']));
      data.extras=[...extras].map(([col,label])=>({label,value:extraValue(excelRow.getCell(col),workbook.properties.date1904)})).filter(x=>x.value!=='');
      if(!Object.keys(labels).some(k=>data[k]!==''))continue;
      if(!data.projectName)throw Error(`الورقة «${sheet.name}»، الصف ${r}: توجد بيانات دون اسم مشروع. لم تُستورد أي بيانات؛ يرجى استكمال الاسم أو حذف الصف.`);
      const dept=data.departmentName||previous.departmentName||onlyContact.departmentName||defaults.departmentName||'';
      const base=contacts.get(dept)||(previous.departmentName===dept?previous:null)||(defaults.departmentName===dept?defaults:{});
      for(const k of contactKeys)if(!data[k])data[k]=k==='departmentName'?dept:(base[k]||'');
      previous=Object.fromEntries(contactKeys.map(k=>[k,data[k]]));
      if(!data.status)data.status='لم يبدأ';
      if(data.completionPercent==='')data.completionPercent=data.status==='مكتمل'?100:data.status==='لم يبدأ'?0:'';
      if(data.status==='جاري التنفيذ')data.status='قيد التنفيذ';
      if(data.projectType==='تطوير عقاري')data.projectType='تطوير عقاري ومباني';
      if(data.projectType==='تحول رقمي')data.projectType='تحول رقمي وتقنية';
      const errors=projectErrors(data);
      if(errors.length)warnings.push({sheet:sheet.name,row:r,errors});
      records.push(data);sources.push({sheet:sheet.name,row:r});
    }
  }
  if(!records.length)throw Error('الملف لا يحتوي على مشاريع. يرجى إضافة صفوف بيانات تحت عناوين الأعمدة.');
  return {records,warnings,sources};
}