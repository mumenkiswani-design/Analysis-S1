(function(){
  function fmt(n){return Number(n||0).toLocaleString('en-US',{maximumFractionDigits:0});}
  function pct(n){return (Number(n||0)*100).toFixed(1)+'%';}
  function addTitle(slide,title,sub){slide.addText(title,{x:0.45,y:0.28,w:12.2,h:0.45,fontSize:24,bold:true,color:'18304C'});if(sub)slide.addText(sub,{x:0.47,y:0.78,w:12,h:0.3,fontSize:10.5,color:'5A646E'});}
  async function buildPpt(){
    try{
      var P=window.PptxGenJS||window.pptxgenjs||window.pptxgen;
      var T=window.S1_TARGETS;
      if(!P){alert('PowerPoint library is not loaded yet. Please refresh once and try again.');return;}
      if(!T){alert('Dashboard target data is not loaded yet. Please refresh and try again.');return;}
      var pptx=new P();pptx.layout='LAYOUT_WIDE';pptx.author='S1 Business Review';pptx.subject='S1 2026 Business Review';pptx.title='S1 Business Review Reconciled';
      var s=pptx.addSlide();s.background={color:'F8FAFD'};addTitle(s,'S1 2026 Business Review','Editable PowerPoint generated from the live dashboard');
      var cards=[['2025 Restated','221,560 JOD'],['2026 Sales',fmt(T.overall.sales)+' JOD'],['Target',fmt(T.overall.target)+' JOD'],['Achievement',pct(T.overall.ach)]];
      cards.forEach(function(c,i){var x=0.65+i*3.05;s.addShape(pptx.ShapeType.roundRect,{x:x,y:1.35,w:2.75,h:1.15,fill:{color:'F3F7FB'},line:{color:'D9E3EC'}});s.addText(c[0],{x:x+0.15,y:1.53,w:2.45,h:0.22,fontSize:9,bold:true,color:'5A646E'});s.addText(c[1],{x:x+0.15,y:1.9,w:2.45,h:0.3,fontSize:18,bold:true,color:'18304C'});});
      s.addText('2026 is reconciled to the official Medical Rep with-subagents report. Brand achievement uses the official Private-only brand report.',{x:0.8,y:3.15,w:11.7,h:1.0,fontSize:18,color:'1C232D'});
      s=pptx.addSlide();addTitle(s,'Medical Rep Achievement vs Target','Official S1 2026 control');
      s.addChart(pptx.ChartType.bar,[{name:'Sales',labels:T.reps.map(r=>r.name),values:T.reps.map(r=>r.sales)},{name:'Target',labels:T.reps.map(r=>r.name),values:T.reps.map(r=>r.target)}],{x:0.55,y:1.15,w:7.2,h:5.1,showLegend:true,legendPos:'b',catAxisLabelFontSize:9,valAxisLabelFontSize:8});
      s.addTable([['Medical Rep','Sales','Target','Achievement']].concat(T.reps.map(r=>[r.name,fmt(r.sales),fmt(r.target),pct(r.ach)])),{x:7.95,y:1.2,w:4.8,h:4.4,fontSize:9,border:{type:'solid',color:'D9E3EC',pt:0.5},fill:'FFFFFF',margin:0.04});
      s=pptx.addSlide();addTitle(s,'Brand Achievement vs Target','Official S1 2026 brand control');
      s.addChart(pptx.ChartType.bar,[{name:'Sales',labels:T.brands.map(r=>r.name),values:T.brands.map(r=>r.sales)},{name:'Target',labels:T.brands.map(r=>r.name),values:T.brands.map(r=>r.target)}],{x:0.55,y:1.15,w:7.2,h:5.1,showLegend:true,legendPos:'b',catAxisLabelFontSize:10,valAxisLabelFontSize:8});
      s.addTable([['Brand','Sales','Target','Achievement']].concat(T.brands.map(r=>[r.name,fmt(r.sales),fmt(r.target),pct(r.ach)])),{x:7.95,y:1.2,w:4.8,h:3.8,fontSize:9,border:{type:'solid',color:'D9E3EC',pt:0.5},fill:'FFFFFF',margin:0.04});
      s=pptx.addSlide();addTitle(s,'Mohammad Al-Omari Reconciliation','Corrected to the official with-subagents report');
      var mo=T.reps.find(r=>r.name.indexOf('محمد')>=0);s.addText('Official Sales: '+fmt(mo.sales)+' JOD',{x:0.8,y:1.4,w:5.8,h:0.5,fontSize:24,bold:true,color:'18304C'});s.addText('Target: '+fmt(mo.target)+' JOD',{x:0.8,y:2.1,w:5.8,h:0.45,fontSize:20,color:'1C232D'});s.addText('Achievement: '+pct(mo.ach),{x:0.8,y:2.75,w:5.8,h:0.5,fontSize:24,bold:true,color:'2A9D8F'});s.addText('Previous model sales: '+fmt(mo.originalCalcSales)+' JOD\nCorrection: +'+fmt(mo.sales-mo.originalCalcSales)+' JOD\nReason: reconstructed account/subagent allocation was below the official report across multiple SKUs, mainly Hi Dee.',{x:6.3,y:1.45,w:5.8,h:2.2,fontSize:16,color:'1C232D'});
      s=pptx.addSlide();addTitle(s,'Management Priorities','S1 2026');s.addText('• Protect Hi Dee overachievement\n• Accelerate Olaxy recovery\n• Improve Unicast performance\n• Use official targets for performance evaluation\n• Use territory/account allocation for diagnostic analysis',{x:0.9,y:1.4,w:11.3,h:3.6,fontSize:22,color:'1C232D',breakLine:false,fit:'shrink'});
      await pptx.writeFile({fileName:'S1_Business_Review_Reconciled_Editable.pptx'});
    }catch(e){console.error('PowerPoint export failed',e);alert('PowerPoint export failed: '+(e&&e.message?e.message:e));}
  }
  function ensureButton(){var actions=document.querySelector('.actions')||document.querySelector('.top')||document.body;if(!actions)return;var old=document.getElementById('downloadPpt');if(old)old.remove();var b=document.createElement('button');b.id='downloadPpt';b.className='btn';b.textContent='Download PowerPoint';b.style.marginLeft='8px';b.addEventListener('click',buildPpt);actions.appendChild(b);}
  if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',ensureButton);else ensureButton();
})();