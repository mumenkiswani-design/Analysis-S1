(function(){
  const E=window.S1_DATA;
  if(!E||!E.d||!Array.isArray(E.r)) return;
  const D=E.d;
  const official={
    'يزيد حماد':{'Dinixir80':3165.061,'Unicast5':2631.186,'Olaxy20':2402.970,'Olaxy40':10198.456,'Hi Dee':42275.736,'Unicast4':3699.476,'Dinixir300':15017.535,'Dinixir40':3588.477},
    'عبدالله حجية':{'Unicast4':1551.916,'Dinixir80':3152.989,'Unicast5':2544.861,'Olaxy40':8273.471,'Olaxy20':2646.045,'Dinixir40':2379.447,'Hi Dee':34356.456,'Dinixir300':11284.335},
    'محمد العمري':{'Olaxy40':5178.342,'Dinixir80':3672.552,'Olaxy20':2368.060,'Unicast5':3876.798,'Unicast4':3393.829,'Dinixir300':11262.081,'Dinixir40':2669.488,'Hi Dee':36912.168},
    'محمود قطاونة':{'Dinixir80':3736.528,'Unicast5':3762.504,'Olaxy40':8770.003,'Unicast4':3116.965,'Olaxy20':2948.245,'Hi Dee':41744.520,'Dinixir40':3388.386,'Dinixir300':15374.525},
    'بتول عبابنه':{'Unicast4':1696.915,'Dinixir80':2969.026,'Unicast5':2872.781,'Olaxy20':2126.420,'Olaxy40':8599.089,'Dinixir300':8158.475,'Dinixir40':2438.253,'Hi Dee':40838.760}
  };
  function keySku(s){s=String(s||'');if(s.includes('80 ml'))return'Dinixir80';if(s.includes('40 ml'))return'Dinixir40';if(s.includes('Dinixir 300'))return'Dinixir300';if(s.includes('Hi Dee'))return'Hi Dee';if(s.includes('OLAXY')&&s.includes("20's"))return'Olaxy20';if(s.includes('OLAXY')&&s.includes("40's"))return'Olaxy40';if(s.includes('Unicast 4'))return'Unicast4';if(s.includes('Unicast 5'))return'Unicast5';return null;}
  const sums={};
  E.r.forEach(x=>{if(x[0]!==2026)return;const rep=D.rep[x[2]],sku=keySku(D.sku[x[5]]);if(!rep||!sku)return;const k=rep+'|'+sku;sums[k]=(sums[k]||0)+Number(x[6]||0);});
  const factors={},audit=[];
  Object.keys(official).forEach(rep=>Object.keys(official[rep]).forEach(sku=>{const k=rep+'|'+sku,base=sums[k]||0,target=official[rep][sku],f=base?target/base:1;factors[k]=f;audit.push({rep,sku,before:base,official:target,factor:f,gap:target-base});}));
  E.r.forEach(x=>{if(x[0]!==2026)return;const rep=D.rep[x[2]],sku=keySku(D.sku[x[5]]),k=rep+'|'+sku;if(factors[k])x[6]=Number(x[6]||0)*factors[k];});
  window.S1_RECONCILIATION={basis:'Ach per Rep S1(1).pdf — Private with subagents — Jan–Jun 2026',method:'Scale each 2026 Rep×SKU group to the official net amount while preserving its month/area mix.',audit:audit};
})();