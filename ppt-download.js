(function(){
  const btn=document.getElementById('downloadPptBtn');
  if(!btn) return;

  const reps=[
    ['Yazeed Hammad',82978.897,78801.570,105.301],
    ['Abdallah Hijjeh',66189.520,79661.258,83.089],
    ['Mohammad Al-Omari',69333.317,80853.102,85.752],
    ['Mahmoud Qatawneh',82841.675,86882.131,95.349],
    ['Batool Ababneh',69699.718,81381.308,85.646]
  ];
  const skus=[
    ['Hi Dee Drops',196127.64],['Dinixir 300',61104.95],['Olaxy 40',41019.36],
    ['Dinixir 80 ml',16696.16],['Unicast 5',15688.13],['Dinixir 40 ml',14464.05],
    ['Unicast 4',13459.10],['Olaxy 20',12491.74]
  ];
  const total2025=122577.33, total2026=371043.126, target2026=407579.368, achievement=91.036;

  function loadPptxGen(){
    if(window.PptxGenJS) return Promise.resolve(window.PptxGenJS);
    return new Promise((resolve,reject)=>{
      const s=document.createElement('script');
      s.src='https://cdn.jsdelivr.net/npm/pptxgenjs@3.12.0/dist/pptxgen.bundle.js';
      s.onload=()=>window.PptxGenJS?resolve(window.PptxGenJS):reject(new Error('PowerPoint library did not initialize'));
      s.onerror=()=>reject(new Error('Could not load PowerPoint library. Check your internet connection.'));
      document.head.appendChild(s);
    });
  }

  function addTitle(slide,title,subtitle){
    slide.addShape('rect',{x:0,y:0,w:13.333,h:0.10,fill:{color:'2563EB'},line:{color:'2563EB'}});
    slide.addText(title,{x:0.65,y:0.35,w:12,h:0.42,fontFace:'Aptos Display',fontSize:22,bold:true,color:'13233A',margin:0});
    if(subtitle) slide.addText(subtitle,{x:0.65,y:0.82,w:12,h:0.25,fontFace:'Aptos',fontSize:9,color:'667085',margin:0});
  }
  function addFooter(slide,n){
    slide.addText('S1 Business Review · January–June 2026',{x:0.65,y:7.15,w:5.5,h:0.18,fontFace:'Aptos',fontSize:7,color:'98A2B3',margin:0});
    slide.addText(String(n),{x:12.15,y:7.15,w:0.5,h:0.18,fontFace:'Aptos',fontSize:7,color:'98A2B3',align:'right',margin:0});
  }
  function card(slide,x,y,w,label,value,accent){
    slide.addShape('roundRect',{x,y,w,h:1.02,rectRadius:0.06,fill:{color:'FFFFFF'},line:{color:'D0D5DD',width:1}});
    slide.addShape('rect',{x,y,w:0.06,h:1.02,fill:{color:accent},line:{color:accent}});
    slide.addText(label,{x:x+0.20,y:y+0.16,w:w-0.3,h:0.18,fontFace:'Aptos',fontSize:7,bold:true,color:'667085',margin:0});
    slide.addText(value,{x:x+0.20,y:y+0.45,w:w-0.3,h:0.30,fontFace:'Aptos Display',fontSize:17,bold:true,color:'13233A',margin:0});
  }

  async function buildDeck(){
    const PptxGenJS=await loadPptxGen();
    const pptx=new PptxGenJS();
    pptx.layout='LAYOUT_WIDE';
    pptx.author='S1 Business Review Dashboard';
    pptx.subject='Pharmaceutical Sales Business Review';
    pptx.title='S1 Medical Rep Business Review';
    pptx.company='Medical Sales Analytics';
    pptx.lang='en-US';
    pptx.theme={headFontFace:'Aptos Display',bodyFontFace:'Aptos',lang:'en-US'};

    // Slide 1 — Overview
    let s=pptx.addSlide(); s.background={color:'F8FAFC'};
    addTitle(s,'S1 Sales Performance Dashboard','January–June 2025 vs 2026 · Medical Rep, Brand, SKU and Area Analytics');
    s.addText('FIRST-LINE SALES MANAGER | PHARMACEUTICAL SALES',{x:0.65,y:1.32,w:6.2,h:0.22,fontSize:8,bold:true,color:'475467',margin:0});
    card(s,0.65,1.80,2.65,'S1 2026 NET SALES',Math.round(total2026).toLocaleString()+' JOD','2563EB');
    card(s,3.50,1.80,2.65,'S1 2026 TARGET',Math.round(target2026).toLocaleString()+' JOD','F59E0B');
    card(s,6.35,1.80,2.65,'ACHIEVEMENT',achievement.toFixed(1)+'%','16A34A');
    card(s,9.20,1.80,2.65,'GROWTH VS 2025',((total2026/total2025-1)*100).toFixed(1)+'%','14B8A6');
    s.addShape('roundRect',{x:0.65,y:3.25,w:11.85,h:2.1,fill:{color:'FFFFFF'},line:{color:'E4E7EC'}});
    s.addText('Management message',{x:0.95,y:3.55,w:2.2,h:0.25,fontSize:11,bold:true,color:'13233A',margin:0});
    s.addText('S1 2026 delivered strong growth versus 2025, but the business remains below the official target. The management priority is to close the remaining gap through SKU-level and rep-level execution while protecting the leading growth drivers.',{x:0.95,y:4.00,w:10.9,h:0.85,fontSize:15,color:'344054',breakLine:false,margin:0.03,fit:'shrink'});
    addFooter(s,1);

    // Slide 2 — Executive summary
    s=pptx.addSlide(); s.background={color:'F8FAFC'}; addTitle(s,'Executive Summary','Validated S1 2026 control total and management perspective');
    s.addChart(pptx.ChartType.bar,[{name:'Sales / Target',labels:['S1 2025','S1 2026','2026 Target'],values:[total2025,total2026,target2026]}],{
      x:0.70,y:1.35,w:5.6,h:4.8,catAxisLabelFontSize:10,valAxisLabelFontSize:9,showLegend:false,showTitle:true,title:'S1 Sales vs Target',showValue:false,
      chartColors:['4472C4'],showCatName:false,valGridLine:{color:'D0D5DD',width:1},showBorder:false
    });
    s.addText('S1 2026 net sales: 371,043 JOD vs S1 2025: 122,577 JOD. Official achievement: 91.0% against 407,579 JOD target. The remaining gap is 36,536 JOD. Management focus should protect top drivers, improve gap SKUs, and reinforce account-level execution.',{x:6.65,y:1.55,w:5.8,h:2.55,fontSize:15,color:'344054',margin:0.04,fit:'shrink'});
    s.addText('Key priorities',{x:6.65,y:4.35,w:2,h:0.25,fontSize:11,bold:true,color:'13233A',margin:0});
    s.addText('• Close the target gap\n• Protect Hi Dee and Dinixir 300 contribution\n• Improve low-achievement SKU mix\n• Focus weekly coaching on under-target reps',{x:6.65,y:4.75,w:5.5,h:1.2,fontSize:12,color:'475467',margin:0.02,breakLine:false});
    addFooter(s,2);

    // Slide 3 — Medical Rep
    s=pptx.addSlide(); s.background={color:'F8FAFC'}; addTitle(s,'Medical Rep Achievement','Official S1 2026 net sales, target and achievement');
    s.addChart(pptx.ChartType.bar,[
      {name:'Net Sales',labels:reps.map(r=>r[0]),values:reps.map(r=>r[1])},
      {name:'Target',labels:reps.map(r=>r[0]),values:reps.map(r=>r[2])}
    ],{x:0.65,y:1.25,w:7.1,h:5.3,catAxisLabelFontSize:10,valAxisLabelFontSize:8,showLegend:true,legendPos:'b',chartColors:['4472C4','C55A11'],showTitle:false,showBorder:false,valGridLine:{color:'D0D5DD'}});
    const rows=[['Medical Rep','Sales','Target','Ach %'],...reps.map(r=>[r[0],Math.round(r[1]).toLocaleString(),Math.round(r[2]).toLocaleString(),r[3].toFixed(1)+'%'])];
    s.addTable(rows,{x:7.95,y:1.35,w:4.65,h:4.8,border:{type:'solid',color:'D0D5DD',pt:1},fill:'FFFFFF',color:'344054',fontFace:'Aptos',fontSize:9,margin:0.05,rowH:0.52,
      bold:false,autoFit:false,
      fillHeader:'13233A'});
    addFooter(s,3);

    // Slide 4 — SKU performance
    s=pptx.addSlide(); s.background={color:'F8FAFC'}; addTitle(s,'Brand / SKU Performance','S1 2026 value contribution by SKU');
    s.addChart(pptx.ChartType.bar,[{name:'Net Sales',labels:skus.map(x=>x[0]),values:skus.map(x=>x[1])}],{x:0.85,y:1.25,w:11.65,h:5.45,catAxisLabelFontSize:10,valAxisLabelFontSize:9,showLegend:false,chartColors:['4472C4'],showTitle:true,title:'S1 2026 Net Sales by SKU',showBorder:false,valGridLine:{color:'D0D5DD'}});
    addFooter(s,4);

    // Slide 5 — ROTT
    s=pptx.addSlide(); s.background={color:'F8FAFC'}; addTitle(s,'ROTT Business Review Brief','Role · Objective · Type · Tone');
    const rott=[
      ['Role','Act as a First-Line Sales Manager presenting a validated S1 business review to senior management.'],
      ['Objective','Explain S1 results, drivers, gaps, opportunities and corrective actions by rep, brand, SKU and area.'],
      ['Type','Interactive performance review supported by editable PowerPoint, data validation and action tracking.'],
      ['Tone','Professional, concise, data-driven, accountable and action-oriented.']
    ];
    rott.forEach((r,i)=>{const x=i%2===0?0.80:6.80, y=i<2?1.55:4.00; s.addShape('roundRect',{x,y,w:5.65,h:1.65,fill:{color:'FFFFFF'},line:{color:'D0D5DD'}}); s.addText(r[0],{x:x+0.25,y:y+0.20,w:1.3,h:0.25,fontSize:11,bold:true,color:'2563EB',margin:0}); s.addText(r[1],{x:x+0.25,y:y+0.60,w:5.05,h:0.72,fontSize:11,color:'475467',margin:0.02,fit:'shrink'});});
    addFooter(s,5);

    // Slide 6 — Management actions
    s=pptx.addSlide(); s.background={color:'F8FAFC'}; addTitle(s,'Recommended Management Actions','Next steps for closing the S1 gap');
    s.addTable([
      ['Priority','Action','Business implication'],
      ['1','Close target gap','Prioritize reps/SKUs below target and define weekly recovery run-rate.'],
      ['2','Protect top drivers','Secure Hi Dee and Dinixir 300 availability, coverage and prescribing momentum.'],
      ['3','Improve SKU mix','Push lower-achievement SKUs through focused account lists and coaching.'],
      ['4','Track execution','Weekly dashboard follow-up by rep, area, brand, SKU, achievement and variance.']
    ],{x:0.75,y:1.35,w:11.8,h:3.9,border:{type:'solid',color:'D0D5DD',pt:1},fill:'FFFFFF',color:'344054',fontFace:'Aptos',fontSize:11,margin:0.07,rowH:0.64});
    s.addText('All slide elements are editable in Microsoft PowerPoint, including text, shapes, tables and charts.',{x:0.80,y:5.80,w:11.5,h:0.45,fontSize:12,bold:true,color:'2563EB',margin:0});
    addFooter(s,6);

    await pptx.writeFile({fileName:'S1_Medical_Rep_Business_Review_Editable.pptx'});
  }

  btn.addEventListener('click',async()=>{
    const old=btn.textContent;
    btn.disabled=true; btn.textContent='Preparing PowerPoint…';
    try{ await buildDeck(); }
    catch(err){ console.error(err); alert('PowerPoint export failed: '+err.message); }
    finally{ btn.disabled=false; btn.textContent=old; }
  });
})();