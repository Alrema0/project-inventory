export const censusStatuses = ['لم يبدأ','قيد التنفيذ','متأخر','مكتمل'];
export const labels = {departmentName:'اسم الإدارة',projectManager:'المسؤول عن المشروع',phone:'رقم التواصل',email:'البريد الإلكتروني',projectName:'اسم المشروع',projectType:'نوع المشروع',description:'وصف مختصر',contractValue:'القيمة التعاقدية بالريال السعودي',startDate:'تاريخ البداية',plannedEndDate:'تاريخ النهاية المتوقع',status:'الحالة',completionPercent:'نسبة الإنجاز %',notes:'ملاحظات أو معوقات'};
export const numeric = value => Number(String(value).replace(/[٠-٩۰-۹]/g,c=>'٠١٢٣٤٥٦٧٨٩'.includes(c)?'٠١٢٣٤٥٦٧٨٩'.indexOf(c):'۰۱۲۳۴۵۶۷۸۹'.indexOf(c)).replace(/٬/g,'').replace(/٫/g,'.'));
export function cleanProject(input) {
  return {...Object.fromEntries(Object.keys(labels).map(k=>[k,['contractValue','completionPercent'].includes(k)?numeric(input?.[k]??''):String(input?.[k]??'').trim()]))};
}
export function projectErrors(input) {
  const p=cleanProject(input), errors=[];
  for(const [key,label] of Object.entries(labels)) {
    if(key!=='notes' && (input?.[key]===undefined || String(input[key]).trim()==='')) errors.push(`${label}: يرجى استكمال الحقل.`);
    if(String(input?.[key]??'').length>4000) errors.push(`${label}: الحد الأقصى ٤٠٠٠ حرف.`);
  }
  if(!censusStatuses.includes(p.status)) errors.push('يرجى اختيار حالة المشروع.');
  if(!Number.isFinite(p.contractValue)||p.contractValue<=0||p.contractValue>1e12) errors.push('يرجى مراجعة القيمة التعاقدية: قيمة موجبة لا تتجاوز تريليون ريال.');
  if(!Number.isInteger(p.completionPercent)||p.completionPercent<0||p.completionPercent>100) errors.push('نسبة الإنجاز عدد صحيح بين ٠ و١٠٠.');
  if(p.status==='مكتمل'&&p.completionPercent!==100) errors.push('المشروع المكتمل يحتاج نسبة إنجاز ١٠٠٪؛ راجع النسبة أو الحالة.');
  if(p.status==='لم يبدأ'&&p.completionPercent!==0) errors.push('المشروع الذي لم يبدأ تكون نسبة إنجازه ٠٪.');
  if(['قيد التنفيذ','متأخر'].includes(p.status)&&p.completionPercent===100) errors.push('الإنجاز ١٠٠٪؛ يرجى مراجعة الحالة واختيار مكتمل عند اكتمال المشروع.');
  if(p.status==='قيد التنفيذ'&&p.completionPercent===0) errors.push('قيد التنفيذ يتطلب نسبة أكبر من ٠٪؛ راجع النسبة أو اختر لم يبدأ.');
  const validDate=d=>/^\d{4}-\d{2}-\d{2}$/.test(d)&&Number.isFinite(Date.parse(d))&&new Date(d).toISOString().slice(0,10)===d;
  if(!validDate(p.startDate)||!validDate(p.plannedEndDate)) errors.push('يرجى إدخال تاريخين صحيحين.');
  else if(p.plannedEndDate<p.startDate) errors.push('تاريخ النهاية يسبق البداية؛ يرجى تصحيح أحد التاريخين.');
  if(!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(p.email)) errors.push('يرجى إدخال بريد إلكتروني صحيح.');
  if(!/^\+?[\d\s()\-]{7,22}$/.test(String(p.phone).replace(/[٠-٩]/g,c=>'٠١٢٣٤٥٦٧٨٩'.indexOf(c)))) errors.push('يرجى إدخال رقم تواصل صحيح مع مفتاح الدولة عند الحاجة.');
  return [...new Set(errors)];
}
