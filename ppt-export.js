(function(){
  function ensureButton(){
    var actions=document.querySelector('.actions') || document.querySelector('.top') || document.body;
    if(!actions || document.getElementById('downloadPpt')) return;
    var a=document.createElement('a');
    a.id='downloadPpt';
    a.className='btn';
    a.textContent='Download PowerPoint';
    a.href='downloads/S1_Business_Review_Reconciled_Editable.pptx?v=12';
    a.setAttribute('download','S1_Business_Review_Reconciled_Editable.pptx');
    a.style.marginLeft='8px';
    a.style.textDecoration='none';
    a.style.display='inline-flex';
    a.style.alignItems='center';
    actions.appendChild(a);
  }
  if(document.readyState==='loading') document.addEventListener('DOMContentLoaded',ensureButton); else ensureButton();
})();