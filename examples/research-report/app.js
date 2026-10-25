'use strict';
// Standalone interface studies. All content is synthetic; no network calls.
(() => {
 const q=(root,s)=>root.querySelector(s);
 const n=(v)=>new Intl.NumberFormat('en-US',{maximumFractionDigits:1}).format(v);
 const head=(brand,sub)=>`<div class="demo-head"><span class="demo-brand">${brand}</span><span class="demo-sub">${sub}</span></div>`;
 const foot=(left,right='Synthetic example')=>`<div class="demo-foot"><span>${left}</span><span>${right}</span></div>`;
 const demos={
  vfund(root){
   const series={steady:[100,101,100,103,105,104,107,108,110,109,112,114,113,116,117,120,118,121,123,124,123,126,125,128],volatile:[100,108,102,111,105,114,104,110,119,112,121,116,125,118,126,120,132,123,136,129,137,130,142,138],drawdown:[100,103,106,104,109,112,107,101,94,88,82,85,89,87,93,98,102,100,107,111,114,119,117,121]};
   root.innerHTML=head('Research report','INTERFACE STUDY')+`<div class="demo-content"><div class="control-row"><label><span class="demo-label">Synthetic dataset</span><select class="demo-select" aria-label="Synthetic dataset"><option value="steady">Steady path</option><option value="volatile">Volatile path</option><option value="drawdown">Drawdown & recovery</option></select></label><span class="demo-sub">Indexed starting value: 100</span></div><div class="chart-slot" aria-live="polite"></div><table class="demo-table"><thead><tr><th>Artifact</th><th>Purpose</th><th class="num">State</th></tr></thead><tbody><tr><td>Run manifest</td><td>Code, inputs, configuration</td><td class="num">Recorded</td></tr><tr><td>Portfolio history</td><td>Sequential accounting</td><td class="num">Recorded</td></tr><tr><td>Reconciliation</td><td>Independent checks</td><td class="num">Reviewed</td></tr></tbody></table><p class="demo-note">Illustrates report organization only. No strategy, market data, backtest result, or investment return is represented.</p></div>`+foot('SYNTHETIC / NO INVESTMENT STRATEGY');
   function render(){
    const values=series[q(root,'select').value];const W=680,H=240,left=45,right=15,top=30,bottom=37;
    const X=i=>left+i*(W-left-right)/(values.length-1),Y=v=>H-bottom-(v-70)*(H-top-bottom)/80;
    const path=values.map((v,i)=>`${i?'L':'M'}${X(i).toFixed(1)},${Y(v).toFixed(1)}`).join(' ');
    const grid=[80,100,120,140].map(v=>`<g><line x1="${left}" x2="${W-right}" y1="${Y(v)}" y2="${Y(v)}" stroke="#e1e7ef"/><text x="${left-10}" y="${Y(v)+4}" text-anchor="end" fill="#5b6878" font-size="11">${v}</text></g>`).join('');
    q(root,'.chart-slot').innerHTML=`<svg class="demo-chart" viewBox="0 0 ${W} ${H}" role="img" aria-label="Synthetic indexed-value path; no actual performance is shown"><title>Synthetic indexed-value path</title>${grid}<path d="${path}" fill="none" stroke="#245bb5" stroke-width="2.5"/><circle cx="${X(values.length-1)}" cy="${Y(values.at(-1))}" r="3.5" fill="#245bb5"/><text x="${left}" y="${H-12}" fill="#5b6878" font-size="11">Observation 1</text><text x="${W-right}" y="${H-12}" text-anchor="end" fill="#5b6878" font-size="11">Observation 24</text><text x="${left}" y="16" fill="#5b6878" font-size="10">Indexed value (synthetic)</text></svg>`;
   }
   q(root,'select').addEventListener('change',render);render();
  },
  liftlog(root){
   const key='varun-portfolio:workout-study:v1';let timer;const defaults=()=>[{weight:95,reps:8},{weight:95,reps:8},{weight:95,reps:6}];let sets=defaults();
   try{const saved=JSON.parse(localStorage.getItem(key));if(Array.isArray(saved)&&saved.length&&saved.length<=8&&saved.every(s=>Number.isFinite(s.weight)&&s.weight>=0&&s.weight<=1000&&Number.isInteger(s.reps)&&s.reps>=0&&s.reps<=100))sets=saved;}catch{}
   root.innerHTML=head('LiftLog','WORKOUT INTERFACE STUDY')+`<div class="demo-content"><h3 class="workout-name">Upper body / Bench press</h3><p class="workout-context">Last sample session: 3 sets · 95 lb · 8 / 8 / 6 reps</p><table class="demo-table workout-table"><thead><tr><th>Set</th><th>Weight (lb)</th><th>Reps</th><th class="num">Volume (lb)</th><th><span class="sr-only">Remove</span></th></tr></thead><tbody></tbody></table><div class="workout-total"><span>Session volume</span><strong class="total" aria-live="polite"></strong></div><div class="control-row"><button class="mini-button add" type="button">+ Add set</button><button class="mini-button reset" type="button">Reset sample</button></div><p class="demo-note">Edit a set, then reload to try local recovery. This example never connects to LiftLog or its database.</p></div>`+foot('<span class="save-state" aria-live="polite">Browser-local sample</span>','Invented workout data');
   function total(){q(root,'.total').textContent=n(sets.reduce((t,s)=>t+s.weight*s.reps,0))+' lb';root.querySelectorAll('[data-volume]').forEach((td,i)=>td.textContent=n(sets[i].weight*sets[i].reps));q(root,'.add').disabled=sets.length>=8;}
   function save(){clearTimeout(timer);q(root,'.save-state').textContent='Saving sample…';timer=setTimeout(()=>{try{localStorage.setItem(key,JSON.stringify(sets));q(root,'.save-state').textContent='Saved on this device';}catch{q(root,'.save-state').textContent='Storage unavailable; sample stays in this tab';}},600);}
   function render(){q(root,'tbody').innerHTML=sets.map((s,i)=>`<tr><td>${i+1}</td><td><input type="number" min="0" max="1000" step="2.5" inputmode="decimal" aria-label="Set ${i+1} weight in pounds" value="${s.weight}" data-field="weight" data-index="${i}"></td><td><input type="number" min="0" max="100" step="1" inputmode="numeric" aria-label="Set ${i+1} repetitions" value="${s.reps}" data-field="reps" data-index="${i}"></td><td class="num" data-volume></td><td><button type="button" class="remove-set" data-remove="${i}" aria-label="Remove set ${i+1}" ${sets.length===1?'disabled':''}>×</button></td></tr>`).join('');total();}
   root.addEventListener('input',e=>{if(!e.target.matches('[data-field]'))return;const field=e.target.dataset.field;let value=Number(e.target.value);if(!Number.isFinite(value)||value<0||value>(field==='weight'?1000:100))return;if(field==='reps'&&!Number.isInteger(value))return;sets[Number(e.target.dataset.index)][field]=value;total();save();});
   root.addEventListener('change',e=>{if(e.target.matches('[data-field]'))e.target.value=sets[Number(e.target.dataset.index)][e.target.dataset.field];});
   root.addEventListener('click',e=>{const b=e.target.closest('[data-remove]');if(b&&sets.length>1){const index=Number(b.dataset.remove);sets.splice(index,1);render();save();q(root,'.add').focus();}});
   q(root,'.add').addEventListener('click',()=>{if(sets.length<8){sets.push({...sets.at(-1)});render();save();q(root,`input[data-index="${sets.length-1}"]`).focus();}});
   q(root,'.reset').addEventListener('click',()=>{sets=defaults();render();save();});render();
  }
 };
 document.querySelectorAll("[data-demo]").forEach(root=>{const demo=demos[root.dataset.demo];if(demo)demo(root);});
})();
