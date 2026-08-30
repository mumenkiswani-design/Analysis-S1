(function(){
  function fmt(n){ return Number(n||0).toLocaleString('en-US',{maximumFractionDigits:0}); }
  function pct(n){ return (Number(n||0)*100).toFixed(1)+'%'; }
  function rowsToTable(rows){ return rows; }
  function ensureButton(){
    var actions=document.querySelector('.actions') || document.querySelector('.top') || document.body;
    if(!actions || document.getElementById('downloadPpt')) return;
    var btn=document.createElement('button');
    btn.id='downloadPpt'; btn.className='btn'; btn.textContent='Download PowerPoint';
    btn.style.marginLeft='8px';
    btn.addEventListener('click', buildPpt);
    actions.appendChild(btn);
  }
  function addTitle(slide,title,sub){
    slide.addText(title,{x:0.35,y:0.25,w:12.6,h:0.45,fontSize:25,bold:true,color:'18304C'});
    if(sub) slide.addText(sub,{x:0.38,y:0.72,w:12.4,h:0.32,fontSize:10.5,color:'5A646E'});
  }
  function addCard(slide,x,y,w,h,label,value,sub,color){
    slide.addShape(pptx.ShapeType.roundRect,{x:x,y:y,w:w,h:h,rectRadius:0.06,fill:{color:'F5F8FC'},line:{color:'E0E8F0'}});
    slide.addText(label,{x:x+0.12,y:y+0.1,w:w-0.24,h:0.2,fontSize:8.5,bold:true,color:'5A646E'});
    slide.addText(value,{x:x+0.12,y:y+0.38,w:w-0.24,h:0.36,fontSize:19,bold:true,color:color||'18304C'});
    if(sub) slide.addText(sub,{x:x+0.12,y:y+0.82,w:w-0.24,h:0.2,fontSize:7.5,color:'5A646E'});
  }
  function simpleTable(slide,x,y,w,h,headers,rows){
    var data=[headers].concat(rows);
    slide.addTable(data,{x:x,y:y,w:w,h:h,border:{type:'solid',color:'D9E2EC',pt:0.6},fontSize:8.5,color:'1C232D',margin:0.05,fill:'FFFFFF',autoFit:false,
      bold:true,fill:'FFFFFF',
      rowH:0.3,
      valign:'mid',
      margin:0.04,
      breakLine:false,
      fit:'shrink',
      // pptxgen applies cell style arrays poorly, keep simple
    });
    slide.addShape(pptx.ShapeType.rect,{x:x,y:y,w:w,h:0.28,fill:{color:'18304C'},line:{color:'18304C'}});
    var cw=w/headers.length;
    headers.forEach(function(hd,i){ slide.addText(hd,{x:x+i*cw+0.04,y:y+0.05,w:cw-0.08,h:0.18,fontSize:7.8,bold:true,color:'FFFFFF'}); });
  }
  function buildPpt(){
    var P=window.pptxgen || window.pptxgenjs;
    if(!P){ alert('PowerPoint library is still loading. Please try again in a few seconds.'); return; }
    var T=window.S1_TARGETS;
    if(!T){ alert('Target data is not loaded yet. Please refresh and try again.'); return; }
    window.pptx = new P();
    pptx.layout='LAYOUT_WIDE'; pptx.author='ChatGPT'; pptx.subject='S1 Business Review'; pptx.title='S1 Business Review Reconciled'; pptx.company='MS Pharma';
    pptx.defineLayout({name:'LAYOUT_WIDE',width:13.333,height:7.5});
    var slide;
    slide=pptx.addSlide(); slide.background={color:'F8FAFD'}; addTitle(slide,'S1 2026 Business Review Dashboard','Reconciled official targets and achievement | January–June 2026');
    addCard(slide,0.6,1.25,2.75,1.08,'Official Sales',fmt(T.overall.sales)+' JOD','Private with subagents','2A9D8F');
    addCard(slide,3.55,1.25,2.75,1.08,'Target',fmt(T.overall.target)+' JOD','Official S1 target','18304C');
    addCard(slide,6.5,1.25,2.75,1.08,'Achievement',pct(T.overall.ach),'Sales / Target','E09F3E');
    addCard(slide,9.45,1.25,2.75,1.08,'Original Model Gap',fmt(T.overall.sales-T.overall.originalCalcSales)+' JOD','Now reconciled','BE5050');
    slide.addText([
      {text:'Key message\n',options:{bold:true}},
      {text:'2026 Medical Rep × SKU results are reconciled to the official “Private with subagents” report. Brand achievement uses the official “Private only” brand report as the brand control source. Mohammad Al-Omari is corrected to the official 69,333 JOD and 85.75% achievement.'}
    ],{x:0.75,y:2.8,w:11.8,h:1.15,fontSize:17,color:'1C232D',breakLine:false,fit:'shrink'});
    slide.addText('Use the official target values for performance evaluation and the allocation model for territory, channel and area diagnosis.',{x:0.75,y:4.25,w:11.8,h:0.55,fontSize:16,color:'1C232D'});
    
    slide=pptx.addSlide(); addTitle(slide,'Medical Rep Achievement vs Target','Official “Private with subagents” control');
    slide.addChart(pptx.ChartType.bar,[{name:'Sales',labels:T.reps.map(r=>r.name),values:T.reps.map(r=>r.sales)},{name:'Target',labels:T.reps.map(r=>r.name),values:T.reps.map(r=>r.target)}],{x:0.55,y:1.05,w:7.05,h:4.8,showLegend:true,showValue:false,catAxisLabelFontSize:8,valAxisLabelFontSize:8,legendPos:'b'});
    simpleTable(slide,7.85,1.05,4.95,4.3,['Rep','Sales','Target','Ach.','Gap'],T.reps.map(r=>[r.name,fmt(r.sales),fmt(r.target),pct(r.ach),fmt(r.sales-r.target)]));
    
    slide=pptx.addSlide(); addTitle(slide,'Brand Achievement vs Target','Official brand report control');
    slide.addChart(pptx.ChartType.bar,[{name:'Sales',labels:T.brands.map(r=>r.name),values:T.brands.map(r=>r.sales)},{name:'Target',labels:T.brands.map(r=>r.name),values:T.brands.map(r=>r.target)}],{x:0.55,y:1.05,w:7.1,h:4.8,showLegend:true,catAxisLabelFontSize:9,valAxisLabelFontSize:8,legendPos:'b'});
    simpleTable(slide,7.95,1.1,4.75,3.3,['Brand','Sales','Target','Ach.','Gap'],T.brands.map(r=>[r.name,fmt(r.sales),fmt(r.target),pct(r.ach),fmt(r.sales-r.target)]));
    slide.addText('Hi Dee is above target and is the strongest brand. Olaxy and Unicast remain the biggest recovery priorities.',{x:8.05,y:4.7,w:4.55,h:0.75,fontSize:13,color:'1C232D'});
    
    slide=pptx.addSlide(); addTitle(slide,'Mohammad Al-Omari Reconciliation','Reason for the previous sales and achievement gap');
    var mo=T.reps.filter(r=>r.name.indexOf('محمد')>=0)[0];
    addCard(slide,0.65,1.15,2.8,1.05,'Original model',fmt(mo.originalCalcSales)+' JOD',null,'18304C');
    addCard(slide,3.65,1.15,2.8,1.05,'Official sales',fmt(mo.sales)+' JOD',null,'2A9D8F');
    addCard(slide,6.65,1.15,2.8,1.05,'Correction','+'+fmt(mo.sales-mo.originalCalcSales)+' JOD',null,'E09F3E');
    addCard(slide,9.65,1.15,2.8,1.05,'Correct achievement',pct(mo.ach),null,'2A9D8F');
    simpleTable(slide,0.85,2.65,4.35,3.05,['SKU','Correction JOD'],[['Hi Dee','+1,017'],['Dinixir 80 ml','+266'],['Dinixir 40 ml','+190'],['Dinixir 300','+168'],['Unicast 4','+159'],['Olaxy 40','+97'],['Olaxy 20','+63'],['Unicast 5','-47']]);
    slide.addText('The gap was not caused by the achievement formula. It came from the reconstructed 2026 account/subagent allocation being below the official “Private with subagents” report across several SKUs, especially Hi Dee. The fix reconciles 2026 Rep × SKU values to the official report while preserving month/area proportions underneath.',{x:5.55,y:2.75,w:6.8,h:2.1,fontSize:15,color:'1C232D',fit:'shrink'});
    
    slide=pptx.addSlide(); addTitle(slide,'Management Priorities','Actions after reconciliation');
    var priorities=['Protect Hi Dee overachievement while monitoring returns and account continuity.','Build an Olaxy recovery plan by rep, focusing on both 20s and 40s target gaps.','Address Unicast underachievement, especially Unicast 5, through customer-level action plans.','Use official targets for performance evaluation and the allocation model for area/account diagnosis.','For Mohammad Al-Omari, use 69,333 JOD official sales and 85.75% achievement in management review.'];
    slide.addText(priorities.map(x=>'• '+x).join('\n'),{x:0.85,y:1.25,w:11.8,h:4.4,fontSize:18,color:'1C232D',breakLine:false,fit:'shrink'});
    
    slide=pptx.addSlide(); addTitle(slide,'ROTT Business Review Frame','Role · Objective · Type · Tone');
    var rott=[['Role','First-Line Sales Manager presenting reconciled sales performance.'],['Objective','Explain achievement, gaps, growth drivers and recovery actions.'],['Type','Interactive S1 business review with official target validation.'],['Tone','Analytical, accountable, concise and action-oriented.']];
    rott.forEach(function(r,i){var x=0.75+(i%2)*6.1,y=1.3+Math.floor(i/2)*2.25;slide.addShape(pptx.ShapeType.roundRect,{x:x,y:y,w:5.65,h:1.55,fill:{color:'F5F8FC'},line:{color:'E0E8F0'}});slide.addText(r[0],{x:x+0.2,y:y+0.18,w:5.2,h:0.35,fontSize:20,bold:true,color:'18304C'});slide.addText(r[1],{x:x+0.2,y:y+0.65,w:5.2,h:0.55,fontSize:13.5,color:'1C232D'});});
    pptx.writeFile({fileName:'S1_Business_Review_Reconciled_Editable.pptx'});
  }
  if(document.readyState==='loading') document.addEventListener('DOMContentLoaded',ensureButton); else ensureButton();
})();