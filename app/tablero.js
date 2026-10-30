(function(){
'use strict';
var OLS=XL.OLS, ITEMS=XL.ITEMS, ILAB=XL.ITEM_LABEL;
var MES=['Ene','Feb','Mar','Abr','May','Jun','Jul','Ago','Sep','Oct','Nov','Dic'];
var MESL=['Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre'];
var COL={Mantenimiento:'#4F8CFF',Neumaticos:'#38D39F',Combustible:'#F6B800',Multas:'#FB7185',Otros:'#A78BFA'};
var DATA_JSON='datos/datos_presupuesto.json', DATA_URL='datos/Plantilla_Presupuesto_OL.xlsx';

/* ---------- estado ---------- */
var model=XL.emptyModel(), info={ok:false,msg:'Cargando datos…',when:null};
var flt='Todos', cut=5, metric='real';

/* ---------- utilidades ---------- */
function $(id){return document.getElementById(id);}
function clone(o){return JSON.parse(JSON.stringify(o));}
var nf0=new Intl.NumberFormat('es-AR',{maximumFractionDigits:0});
var nf1=new Intl.NumberFormat('es-AR',{minimumFractionDigits:1,maximumFractionDigits:1});
var nf2=new Intl.NumberFormat('es-AR',{maximumFractionDigits:2});
function money(n){return '$ '+nf0.format(Math.round(n||0));}
function compact(n){n=n||0;var a=Math.abs(n);if(a>=1e6)return '$ '+(a>=1e8?nf0.format(n/1e6):nf1.format(n/1e6))+' M';if(a>=1e3)return '$ '+nf0.format(n/1e3)+' k';return '$ '+nf0.format(n);}
function pct(x,d){return (d===0?nf0.format(x*100):nf1.format(x*100))+'%';}
function esc(s){return String(s).replace(/[&<>"]/g,function(c){return {'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c];});}
function sumr(a,from,to){var s=0;for(var i=from;i<=to;i++)s+=a[i]||0;return s;}
function niceMax(v){if(v<=0)return 1;var mag=Math.pow(10,Math.floor(Math.log10(v))),n=v/mag,s=n<=1?1:n<=2?2:n<=2.5?2.5:n<=5?5:10;return s*mag;}
function fmtDate(iso){if(!iso)return '';var d=new Date(iso);return d.toLocaleDateString('es-AR',{day:'2-digit',month:'2-digit',year:'numeric'})+' '+d.toLocaleTimeString('es-AR',{hour:'2-digit',minute:'2-digit'});}

/* ---------- series ---------- */
function scope(){return flt==='Todos'?OLS:[flt];}
function ser(t,it,ols){var o=Array(12).fill(0);(ols||scope()).forEach(function(ol){model.data[ol][t][it].forEach(function(v,i){o[i]+=v||0;});});return o;}
function serTot(t,ols){var o=Array(12).fill(0);ITEMS.forEach(function(it){ser(t,it,ols).forEach(function(v,i){o[i]+=v;});});return o;}
function defaultCut(){
  var best=-1;
  for(var m=0;m<12;m++){var n=0;OLS.forEach(function(ol){var s=0;ITEMS.forEach(function(it){s+=model.data[ol].real[it][m]||0;});if(s>0)n++;});if(n>=3)best=m;}
  if(best>=0)return best;
  for(m=11;m>=0;m--){if(serTot('real',OLS)[m]>0)return m;}
  return Math.max(0,new Date().getMonth()-1);
}

/* ---------- carga de datos desde la plantilla ---------- */
function download(name,bytes,type){
  var blob=new Blob([bytes],{type:type||'application/octet-stream'}),a=document.createElement('a');
  a.href=URL.createObjectURL(blob);a.download=name;document.body.appendChild(a);a.click();
  setTimeout(function(){URL.revokeObjectURL(a.href);a.remove();},1500);
}
async function parseBuffer(buf){
  var res=XL.parseWorkbook(await XL.readXlsx(buf),{year:2026});
  if(!res.kind||!res.ols)throw new Error((res.warnings&&res.warnings[0])||'El archivo no tiene el formato de la plantilla (hojas "Presupuesto" y "Gastos").');
  var m=XL.emptyModel();XL.applyImport(m,res.ols);
  return {model:m,warnings:res.warnings||[]};
}
function applyLoaded(r,label,when){
  model=r.model;info={ok:true,msg:label,when:when};cut=defaultCut();
  if(r.warnings.length)toast('<b>Aviso:</b> '+esc(r.warnings.join(' · ')),9000);
}
function fromJson(j){
  if(!j||!j.data)throw new Error('El archivo de datos no tiene el formato esperado.');
  var m=XL.emptyModel();
  OLS.forEach(function(ol){if(!j.data[ol])return;['presu','real','target'].forEach(function(t){ITEMS.forEach(function(it){var a=j.data[ol][t]&&j.data[ol][t][it];if(a&&a.length===12)m.data[ol][t][it]=a.map(function(v){return +v||0;});});});});
  m.meta=j.meta||m.meta;
  return {model:m,warnings:[]};
}
async function loadFromUrl(){
  var resp=await fetch(DATA_JSON+'?_='+Date.now(),{cache:'no-store'});
  if(resp.ok){
    var j=await resp.json(),r=fromJson(j);
    applyLoaded(r,'Datos del repositorio',(j.meta&&j.meta.savedAt)||new Date().toISOString());
    return;
  }
  resp=await fetch(DATA_URL+'?_='+Date.now(),{cache:'no-store'});
  if(!resp.ok)throw new Error('No se encontró '+DATA_JSON+' ni '+DATA_URL+'. Ejecutá publicar.bat para generar los datos.');
  var lm=resp.headers.get('Last-Modified');
  applyLoaded(await parseBuffer(await resp.arrayBuffer()),'Plantilla del repositorio',lm?new Date(lm).toISOString():new Date().toISOString());
}
async function refresh(manual){
  var b=$('btnRefresh');b.disabled=true;
  try{
    if(location.protocol==='file:'){ if(manual){b.disabled=false;$('fileIn').value='';$('fileIn').click();return;} throw new Error('Abierto desde un archivo local: usá "Actualizar datos" para elegir el archivo de datos o la plantilla.'); }
    await loadFromUrl();
    if(manual)toast('<b>Datos actualizados.</b>',4000);
  }catch(e){
    info={ok:false,msg:e.message||String(e),when:null};
    if(manual)toast('<b>No se pudo actualizar.</b> '+esc(info.msg),9000);
  }
  b.disabled=false;render();
}
async function loadPicked(file){
  try{var r=/\.json$/i.test(file.name)?fromJson(JSON.parse(await file.text())):await parseBuffer(await file.arrayBuffer());applyLoaded(r,'Archivo cargado manualmente: '+file.name,new Date(file.lastModified||Date.now()).toISOString());toast('<b>Datos cargados</b> desde '+esc(file.name)+'.',4000);}
  catch(e){info={ok:false,msg:e.message||String(e),when:null};toast('<b>No se pudo leer el archivo.</b> '+esc(info.msg),9000);}
  render();
}

/* ---------- toast y modal ---------- */
var toastT=null;
function toast(html,ms){var t=$('toast');t.innerHTML=html;t.style.display='block';clearTimeout(toastT);toastT=setTimeout(function(){t.style.display='none';},ms||6000);}
/* ---------- render principal ---------- */
function render(){renderTop();renderFilters();renderDerived();}
function renderDerived(){renderEmpty();renderCharts();renderTargets();renderDetail();renderOLTable();}
function renderTop(){
  var st=$('status');
  if(info.ok){st.className='status';st.innerHTML='<i></i>'+esc(info.msg)+(info.when?' · '+esc(fmtDate(info.when)):'');}
  else{st.className='status dirty';st.innerHTML='<i></i>'+esc(info.msg);}
}
function renderFilters(){
  var c=$('olChips');
  c.innerHTML=['Todos'].concat(OLS).map(function(o){return '<button type="button" class="chip-btn" data-o="'+o+'" aria-pressed="'+(flt===o)+'">'+o+'</button>';}).join('');
  c.querySelectorAll('button').forEach(function(b){b.onclick=function(){flt=b.dataset.o;renderFilters();renderDerived();};});
  var s=$('cutSel');
  s.innerHTML=MESL.map(function(n,i){return '<option value="'+i+'"'+(i===cut?' selected':'')+'>'+n+'</option>';}).join('');
  s.onchange=function(){cut=+s.value;renderDerived();};
}

/* Aviso de vacío */
function hasData(){var s=0;OLS.forEach(function(ol){['presu','real'].forEach(function(t){ITEMS.forEach(function(it){model.data[ol][t][it].forEach(function(v){s+=Math.abs(v||0);});});});});return s>0;}
function renderEmpty(){
  var e=$('emptyNote');
  if(hasData()){e.innerHTML='';return;}
  e.innerHTML='<div class="card" style="border-left:6px solid var(--gold);margin-bottom:14px"><h3 style="margin:0 0 6px;font-size:15px;text-transform:uppercase;letter-spacing:.08em">Todavía no hay datos cargados</h3><p style="margin:0;color:var(--mut)">Los datos salen de <b>datos/datos_presupuesto.json</b>, que se genera desde los Excel de la carpeta <b>base_excels</b>. Completala (hojas Presupuesto, Gastos y Targets) y presioná <b>Actualizar datos</b>.</p></div>';
}
/* Targets */
function badge(real,tgt){
  if(!tgt)return '<span class="st">Sin target</span>';
  return real<=tgt?'<span class="st ok">● Dentro</span>':'<span class="st no">● Fuera</span>';
}
function renderTargets(){
  var head='<thead><tr><th>Operador</th>'+ITEMS.map(function(it){return '<th>'+ILAB[it]+'</th>';}).join('')+'<th>Total</th></tr></thead>',b='<tbody>',any=false;
  function cell(r,t){if(t)any=true;return '<td class="tg">'+badge(r,t)+'<small>'+(t?compact(r)+' / '+compact(t)+' · '+pct(r/t,0):compact(r))+'</small></td>';}
  var TR=Array(ITEMS.length).fill(0),TT=Array(ITEMS.length).fill(0);
  OLS.forEach(function(ol){
    var rr=0,tt=0,cells='';
    ITEMS.forEach(function(it,i){var r=sumr(model.data[ol].real[it],0,cut),t=sumr(model.data[ol].target[it],0,cut);if(t>0){rr+=r;tt+=t;TR[i]+=r;TT[i]+=t;}cells+=cell(r,t);});
    b+='<tr><td>'+ol+'</td>'+cells+cell(rr,tt)+'</tr>';
  });
  var gr=0,gt=0,cells2='';ITEMS.forEach(function(it,i){gr+=TR[i];gt+=TT[i];cells2+=cell(TR[i],TT[i]);});
  b+='<tr class="tot"><td>Flota</td>'+cells2+cell(gr,gt)+'</tr></tbody>';
  $('tblTarget').innerHTML=head+b;
  $('tgtNote').textContent='Enero a '+MESL[cut]+' · gasto real acumulado frente al target acumulado (solo ítems con target)'+(any?'.':'. Cargá la hoja "Targets" de la plantilla para ver el cumplimiento.');
}

/* Gráficos */
var tipEl=$('tip');
function showTip(e,html){tipEl.innerHTML=html;tipEl.style.display='block';var w=tipEl.offsetWidth,x=e.clientX+16;if(x+w>window.innerWidth-8)x=e.clientX-w-16;tipEl.style.left=Math.max(8,x)+'px';tipEl.style.top=Math.min(e.clientY+14,window.innerHeight-tipEl.offsetHeight-8)+'px';}
function hideTip(){tipEl.style.display='none';}
function axis(maxV,W,H,pl,pr,pt,pb){
  var s='',t;for(t=0;t<=4;t++){var v=maxV/4*t,y=pt+(H-pt-pb)-(v/maxV)*(H-pt-pb);
    s+='<line x1="'+pl+'" x2="'+(W-pr)+'" y1="'+y+'" y2="'+y+'" stroke="'+(t?'var(--gridA)':'var(--gridB)')+'"/><text x="'+(pl-8)+'" y="'+(y+4)+'" text-anchor="end">'+(v===0?'0':nf0.format(v/1e6)+' M')+'</text>';}
  return s;
}
function chartMonthly(){
  var W=860,H=340,pl=64,pr=12,pt=12,pb=30,ph=H-pt-pb,pw=W-pl-pr;
  var real={};ITEMS.forEach(function(it){real[it]=ser('real',it);});
  var pp=serTot('presu'),stack=Array(12).fill(0).map(function(_,m){return ITEMS.reduce(function(a,it){return a+real[it][m];},0);});
  var maxV=niceMax(Math.max.apply(null,stack.concat(pp))*1.05),slot=pw/12,bw=slot*0.58,s='<svg class="chart" viewBox="0 0 '+W+' '+H+'" role="img" aria-label="Gasto real por ítem y presupuesto total, mes a mes">'+axis(maxV,W,H,pl,pr,pt,pb);
  var cx=pl+slot*(cut+1);
  if(cut<11)s+='<rect x="'+cx+'" y="'+pt+'" width="'+(W-pr-cx)+'" height="'+ph+'" fill="var(--shade)"/>';
  for(var m=0;m<12;m++){
    var x=pl+slot*m+(slot-bw)/2,y=pt+ph;
    ITEMS.forEach(function(it){var h=real[it][m]/maxV*ph;y-=h;if(h>0)s+='<rect x="'+x.toFixed(1)+'" y="'+y.toFixed(1)+'" width="'+bw.toFixed(1)+'" height="'+h.toFixed(1)+'" fill="'+COL[it]+'"/>';});
    s+='<text x="'+(pl+slot*m+slot/2)+'" y="'+(H-10)+'" text-anchor="middle" '+(m>cut?'style="fill:var(--dim)"':'')+'>'+MES[m]+'</text>';
  }
  var pts=pp.map(function(v,m){return (pl+slot*m+slot/2).toFixed(1)+','+(pt+ph-v/maxV*ph).toFixed(1);}).join(' ');
  s+='<polyline points="'+pts+'" fill="none" stroke="var(--halo)" stroke-width="6" stroke-linejoin="round"/><polyline points="'+pts+'" fill="none" stroke="var(--ink)" stroke-width="2.5" stroke-dasharray="7 4" stroke-linejoin="round"/>';
  pp.forEach(function(v,m){s+='<circle cx="'+(pl+slot*m+slot/2).toFixed(1)+'" cy="'+(pt+ph-v/maxV*ph).toFixed(1)+'" r="4" fill="var(--ink)" stroke="var(--halo)" stroke-width="2"/>';});
  for(m=0;m<12;m++)s+='<rect class="hov" data-m="'+m+'" x="'+(pl+slot*m)+'" y="'+pt+'" width="'+slot+'" height="'+ph+'" fill="transparent"/>';
  s+='</svg>';
  $('chMonthly').innerHTML=s;
  $('chMonthly').querySelectorAll('.hov').forEach(function(r){
    r.onmousemove=function(e){var m=+r.dataset.m,t='<b>'+MESL[m]+'</b>';
      ITEMS.forEach(function(it){if(real[it][m])t+='<div><span>'+ILAB[it]+'</span><span>'+money(real[it][m])+'</span></div>';});
      t+='<div style="border-top:1px solid var(--line);margin-top:4px;padding-top:4px"><span>Gasto real</span><b style="color:var(--text);margin:0">'+money(stack[m])+'</b></div><div><span>Presupuesto</span><span>'+money(pp[m])+'</span></div>';
      if(pp[m]&&stack[m])t+='<div><span>Desvío</span><span>'+(stack[m]>pp[m]?'+':'−')+pct(Math.abs((stack[m]-pp[m])/pp[m]))+'</span></div>';
      showTip(e,t);};
    r.onmouseleave=hideTip;
  });
  $('lgMonthly').innerHTML=ITEMS.map(function(it){return '<span><i style="background:'+COL[it]+'"></i>'+ILAB[it]+'</span>';}).join('')+'<span><i class="ln" style="background:var(--ink)"></i>Presupuesto total</span>';
  $('monthlySub').textContent='Los meses posteriores a '+MESL[cut]+' se muestran sombreados (aún sin cierre).';
}
function chartCum(){
  var W=420,H=340,pl=64,pr=14,pt=14,pb=30,ph=H-pt-pb,pw=W-pl-pr;
  var pp=serTot('presu'),rl=serTot('real'),cp=[],cr=[],a=0,b=0;
  for(var i=0;i<12;i++){a+=pp[i];b+=rl[i];cp.push(a);cr.push(b);}
  var maxV=niceMax(Math.max(a,b)*1.05),step=pw/11;
  function X(i){return pl+step*i;} function Y(v){return pt+ph-v/maxV*ph;}
  var s='<svg class="chart" viewBox="0 0 '+W+' '+H+'" role="img" aria-label="Presupuesto y gasto real acumulados">'+axis(maxV,W,H,pl,pr,pt,pb);
  for(i=0;i<12;i+=1)if(i%2===0||i===11)s+='<text x="'+X(i)+'" y="'+(H-10)+'" text-anchor="middle">'+MES[i]+'</text>';
  s+='<polyline points="'+cp.map(function(v,i){return X(i).toFixed(1)+','+Y(v).toFixed(1);}).join(' ')+'" fill="none" stroke="var(--ink)" stroke-width="2.5" stroke-dasharray="7 4"/>';
  var up=cr.slice(0,cut+1);
  s+='<polyline points="'+up.map(function(v,i){return X(i).toFixed(1)+','+Y(v).toFixed(1);}).join(' ')+'" fill="none" stroke="var(--halo)" stroke-width="8"/><polyline points="'+up.map(function(v,i){return X(i).toFixed(1)+','+Y(v).toFixed(1);}).join(' ')+'" fill="none" stroke="#F6B800" stroke-width="4.5" stroke-linejoin="round"/>';
  s+='<circle cx="'+X(cut)+'" cy="'+Y(cr[cut])+'" r="6" fill="#F6B800" stroke="var(--halo)" stroke-width="2.5"/>';
  s+='<text x="'+Math.min(X(cut),W-pr-4)+'" y="'+(Y(cr[cut])-12)+'" text-anchor="end" style="font-weight:700;fill:var(--gold-t)">'+compact(cr[cut])+'</text>';
  for(i=0;i<12;i++)s+='<rect class="hov" data-m="'+i+'" x="'+(X(i)-step/2)+'" y="'+pt+'" width="'+step+'" height="'+ph+'" fill="transparent"/>';
  s+='</svg>';
  $('chCum').innerHTML=s;
  $('chCum').querySelectorAll('.hov').forEach(function(r){
    r.onmousemove=function(e){var m=+r.dataset.m;showTip(e,'<b>Hasta '+MESL[m]+'</b><div><span>Presupuesto acum.</span><span>'+money(cp[m])+'</span></div>'+(m<=cut?'<div><span>Gasto real acum.</span><span>'+money(cr[m])+'</span></div>':''));};
    r.onmouseleave=hideTip;
  });
  $('lgCum').innerHTML='<span><i class="ln" style="background:#F6B800;height:5px"></i>Gasto real</span><span><i class="ln" style="background:var(--ink)"></i>Presupuesto</span>';
}
function olBars(){
  var rows=OLS.map(function(ol){var r=0,p=0;ITEMS.forEach(function(it){r+=sumr(model.data[ol].real[it],0,cut);p+=sumr(model.data[ol].presu[it],0,cut);});return {ol:ol,r:r,p:p};});
  var mx=Math.max.apply(null,rows.map(function(x){return Math.max(x.r,x.p);}).concat([1]));
  $('olBars').innerHTML=rows.map(function(x){
    var dv=x.p?(x.r-x.p)/x.p:null;
    return '<div class="olrow'+(flt===x.ol?' sel':'')+'"><b>'+x.ol+'</b><div class="track" role="img" aria-label="'+x.ol+': gasto real '+compact(x.r)+' de '+compact(x.p)+' presupuestados"><div class="pp" style="width:'+(x.p/mx*100)+'%"></div><div class="rl'+(x.r>x.p?' over':'')+'" style="width:'+(x.r/mx*100)+'%"></div></div><div class="olnums"><span>Real <strong>'+compact(x.r)+'</strong> · Ppto. <strong>'+compact(x.p)+'</strong></span><span class="'+(dv==null?'':dv>0?'bad':'good')+'"><strong>'+(dv==null?'—':(dv>0?'▲ +':'▼ −')+pct(Math.abs(dv)))+'</strong></span></div></div>';
  }).join('');
  $('olSub').textContent='Enero a '+MESL[cut]+' · la marca clara es el presupuesto acumulado.';
}
function mixChart(){
  var vals=ITEMS.map(function(it){return sumr(ser('real',it),0,cut);}),tot=vals.reduce(function(a,b){return a+b;},0)||1;
  $('mix').innerHTML='<div class="mixbar" role="img" aria-label="Composición porcentual del gasto real por ítem">'+ITEMS.map(function(it,i){return '<span style="width:'+(vals[i]/tot*100)+'%;background:'+COL[it]+'" title="'+ILAB[it]+'"></span>';}).join('')+'</div>'+
    ITEMS.map(function(it,i){return '<div class="mixrow"><i style="background:'+COL[it]+'"></i><span>'+ILAB[it]+'</span><strong>'+compact(vals[i])+'</strong><em>'+pct(vals[i]/tot)+'</em></div>';}).join('');
  $('mixSub').textContent='Enero a '+MESL[cut]+' · '+(flt==='Todos'?'los 4 operadores':'operador '+flt);
}
function renderCharts(){chartMonthly();chartCum();olBars();mixChart();}

/* Tablas */
function setMetric(m){metric=m;document.querySelectorAll('#metricSeg button').forEach(function(b){b.setAttribute('aria-pressed',String(b.dataset.m===m));});renderDetail();}
function renderDetail(){
  var head='<thead><tr><th>Ítem</th>'+MES.map(function(m){return '<th>'+m+'</th>';}).join('')+'<th>Acum. '+MES[cut]+'</th><th>Año</th></tr></thead>';
  function row(label,color,r,p,isTot,tg){
    var v=metric==='real'?r:metric==='presu'?p:metric==='target'?tg:r.map(function(x,i){return x-p[i];});
    var cells=v.map(function(x,i){var dim=(metric!=='presu'&&metric!=='target'&&i>cut);var cl=(metric==='desv'&&!dim&&x!==0)?(x>0?'bad':'good'):'';return '<td class="'+(dim?'dim ':'')+cl+'">'+(metric==='real'&&dim&&!x?'·':nf0.format(Math.round(x)))+'</td>';}).join('');
    var ac=sumr(v,0,cut),yr=sumr(v,0,11);
    var cl2=metric==='desv'?(ac>0?'bad':ac<0?'good':''):'';
    return '<tr'+(isTot?' class="tot"':'')+'><td>'+(color?'<span class="sw" style="background:'+color+'"></span>':'')+label+'</td>'+cells+'<td class="'+cl2+'"><b>'+nf0.format(Math.round(ac))+'</b></td><td>'+nf0.format(Math.round(yr))+'</td></tr>';
  }
  var body='<tbody>';
  ITEMS.forEach(function(it){body+=row(ILAB[it],COL[it],ser('real',it),ser('presu',it),false,ser('target',it));});
  body+=row('Total',null,serTot('real'),serTot('presu'),true,serTot('target'))+'</tbody>';
  $('tblDetail').innerHTML=head+body;
}
function renderOLTable(){
  var head='<thead><tr><th>Operador</th><th>Ppto. anual</th><th>Ppto. acum. '+MES[cut]+'</th><th>Real acum. '+MES[cut]+'</th><th>Desvío $</th><th>Desvío %</th><th>% del ppto. anual</th><th>Target acum.</th><th>Estado</th></tr></thead>',b='<tbody>',TP=0,TR=0,TY=0,TT=0,TRT=0;
  function orow(l,y,p,r,tg,t,rt){var d=r-p,dp=p?d/p:null,c=d>0?'bad':d<0?'good':'';return '<tr'+(t?' class="tot"':'')+'><td>'+l+'</td><td>'+money(y)+'</td><td>'+money(p)+'</td><td>'+money(r)+'</td><td class="'+c+'">'+(d>0?'+':d<0?'−':'')+money(Math.abs(d))+'</td><td class="'+c+'">'+(dp==null?'—':(d>0?'+':d<0?'−':'')+pct(Math.abs(dp)))+'</td><td>'+(y?pct(r/y):'—')+'</td><td>'+(tg?money(tg):'—')+'</td><td>'+badge(rt,tg)+'</td></tr>';}
  OLS.forEach(function(ol){var r=0,p=0,y=0,tg=0,rt=0;ITEMS.forEach(function(it){r+=sumr(model.data[ol].real[it],0,cut);if(sumr(model.data[ol].target[it],0,cut)>0)rt+=sumr(model.data[ol].real[it],0,cut);p+=sumr(model.data[ol].presu[it],0,cut);y+=sumr(model.data[ol].presu[it],0,11);tg+=sumr(model.data[ol].target[it],0,cut);});TP+=p;TR+=r;TY+=y;TT+=tg;TRT+=rt;b+=orow(ol,y,p,r,tg,false,rt);});
  b+=orow('Total',TY,TP,TR,TT,true,TRT)+'</tbody>';
  $('tblOL').innerHTML=head+b;
  $('olTblNote').textContent='Enero a '+MESL[cut]+'. Se calcula siempre con los 4 operadores.';
}

/* ---------- publicar ---------- */
function openModal(inner){$('modalRoot').innerHTML='<div class="ov" id="ov"><div class="modal" role="dialog" aria-modal="true">'+inner+'</div></div>';$('ov').addEventListener('mousedown',function(e){if(e.target.id==='ov')closeModal();});document.addEventListener('keydown',escClose);}
function escClose(e){if(e.key==='Escape')closeModal();}
function closeModal(){$('modalRoot').innerHTML='';document.removeEventListener('keydown',escClose);}


/* ---------- tema ---------- */
function applyTheme(t){
  document.documentElement.setAttribute('data-theme',t);
  var b=$('btnTheme');if(b)b.textContent=t==='light'?'☾ Tema oscuro':'☀ Tema claro';
  try{localStorage.setItem('mc_theme',t);}catch(e){}
}
function initTheme(){
  var t=null;try{t=localStorage.getItem('mc_theme');}catch(e){}
  if(t!=='light'&&t!=='dark')t=(window.matchMedia&&matchMedia('(prefers-color-scheme: light)').matches)?'light':'dark';
  applyTheme(t);
}
$('btnTheme').onclick=function(){applyTheme(document.documentElement.getAttribute('data-theme')==='light'?'dark':'light');};

/* ---------- botones ---------- */
$('btnRefresh').onclick=function(){refresh(true);};
$('fileIn').onchange=function(){if(this.files.length)loadPicked(this.files[0]);};
document.querySelectorAll('#metricSeg button').forEach(function(b){b.onclick=function(){setMetric(b.dataset.m);};});

/* ---------- inicio ---------- */
(async function init(){
  initTheme();
  render();
  await refresh(false);
})();
})();
