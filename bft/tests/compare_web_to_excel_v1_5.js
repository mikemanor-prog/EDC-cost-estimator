const fs=require('fs');
const html=fs.readFileSync(__dirname+'/../web/bft_estimator_v1.5.html','utf8');
const src=html.match(/<script id="calc">([\s\S]*?)<\/script>/)[1];
const m={exports:{}}; new Function('module',src)(m); const BFT=m.exports;
const cases=JSON.parse(fs.readFileSync(__dirname+'/cases_v1_5.json')), xl=JSON.parse(fs.readFileSync(process.argv[2]||__dirname+'/excel_results_v1_5.json'));
const map={ // cell -> setter
 "Estimator!B5":(M,v)=>M.est.voltage=v,"Estimator!B6":(M,v)=>M.est.load=v,"Estimator!B7":(M,v)=>M.est.cuT=v,"Estimator!B10":(M,v)=>M.est.bpStud=v,
 "Estimator!B11":(M,v)=>M.est.bpCount=v,"Estimator!B12":(M,v)=>M.est.bpConn=v,"Estimator!B13":(M,v)=>M.est.retention=v,"Estimator!B14":(M,v)=>M.est.joining=v,
 "Estimator!B46":(M,v)=>M.est.outCu=v,"Estimator!B18":(M,v)=>M.est.fuses[0].qty=v,"Estimator!B19":(M,v)=>M.est.fuses[1].qty=v,"Estimator!B20":(M,v)=>M.est.fuses[2].qty=v,
 "Assumptions!B63":(M,v)=>M.a.derate=v,"Assumptions!B58":(M,v)=>M.a.k=v,
 "Busbar Sizer!B5":(M,v)=>M.bs.current=v,"Busbar Sizer!B6":(M,v)=>M.bs.rise=v,"Busbar Sizer!B7":(M,v)=>M.bs.length=v,"Busbar Sizer!B8":(M,v)=>M.bs.width=v,
 "Busbar Sizer!B9":(M,v)=>M.bs.thick=v,"Busbar Sizer!B11":(M,v)=>M.bs.maxWidth=v,"Busbar Sizer!A26":(M,v)=>M.bs.ratios[0]=v};
const zrow={20:"include",21:"placement",22:"L",23:"W",24:"H",25:"studs",26:"size",27:"current",28:"branch"};
const lay={"Layout!B5":(M,v)=>M.layout.walls=v,"Layout!B6":(M,v)=>M.layout.exitEnd=v,"Layout!B7":(M,v)=>M.layout.maxLen=v,"Layout!B8":(M,v)=>M.layout.maxDepth=v};
for(let i=0;i<3;i++){ lay["Layout!C"+(12+i)]=(M,v)=>M.est.fuses[i].short=v; lay["Layout!D"+(12+i)]=(M,v)=>M.est.fuses[i].pitchOverride=v; }
Object.assign(map,lay);
const est={B26:"busCur",B27:"busBolt",B28:"busW",B34:"slots",B35:"bpSec",B36:"baseW",B37:"wallH",B38:"baseL",B39:"divH",B40:"height",B41:"divN",B44:"busLen",B45:"busMass",
 B47:"zCu",B48:"cuTotal",B49:"cuCost",B51:"bolts",B52:"bossVol",B53:"floorVol",B54:"wallVol",B55:"divVol",B57:"moldVol",B58:"baseMass",B59:"baseCost",B66:"coverMass",B67:"coverCost",
 B70:"fuseHw",B71:"retAdd",B72:"joinAdd",B73:"bpPart",B74:"zHw",B75:"zLaser",B86:"total",B87:"perPos",B88:"plasticMass",E62:"overhang",E63:"bpExt",E64:"outExt",E65:"coverL",E66:"coverW",E13:"fitLen",E14:"fitDepth",B40:"height"};
const kz={31:"incl",32:"end",33:"back",34:"wAdd",35:"wBack",36:"dBack",37:"lNeed",38:"bumpH",39:"bumpVol",40:"studs",41:"conn",42:"branchW",43:"autoLen",44:"lenUsed",45:"cuMass",46:"hw",47:"bossU",48:"laser"};
const lyb={51:"nShort",52:"nLong",53:"shortSlots",54:"slots",55:"passes",60:"backDepth",61:"inRow",62:"busFront",63:"colIn",64:"colW",65:"longEnd",66:"busLenRow",67:"busLenShort",68:"shortDepth",69:"shortDiv",70:"longDiv",71:"divLen",72:"divN",
 75:"zIncl",76:"zW",77:"zWBack",78:"zL",79:"zTop",80:"bumpVol",81:"zStuds",82:"zCu",83:"zHw",84:"zBoss",85:"zLaser",86:"fitLen",87:"fitDepth"};
let n=0,bad=0;
const chk=(c,i,k,a,b)=>{n++; const ok=(typeof a==='string'||typeof b==='string')?a===b:Math.abs(a-b)<=1e-9*Math.max(1,Math.abs(a)); if(!ok){bad++;console.log('case',i,k,'xl',a,'js',b);} };
cases.forEach((c,i)=>{
  const M=BFT.defaults();
  for(const [k,v] of Object.entries(c)){ const mm=k.match(/^Layout!([BC])(\d+)$/); if(mm&&!zrow[mm[2]]) { map[k](M,v); continue; } if(mm) M.zones[mm[1]==='B'?0:1][zrow[mm[2]]]=v; else map[k](M,v); }
  const r=BFT.estimate(M), b=BFT.busbar(M), X=xl[i];
  for(const [cell,key] of Object.entries(est)) chk(c,i,cell,X['Estimator!'+cell],r[key]);
  for(const [row,key] of Object.entries(kz)) ['B','C'].forEach((col,z)=>chk(c,i,'LY '+col+row,X['Layout!'+col+row],r.zones[z][key]));
  for(const [row,key] of Object.entries(lyb)) chk(c,i,'LY B'+row,X['Layout!B'+row],r[key]);
  r.fuses.forEach((fz,j)=>{ const rr=12+j; [['E','autoSlot'],['F','slot'],['G','nLong'],['H','pitchCheck'],['I','nShort']].forEach(([col,key])=>chk(c,i,'LY '+col+rr,X['Layout!'+col+rr],fz[key])); });
  const bk={B13:"ratio",B14:"area",B15:"perim",B16:"density",B17:"rated",B18:"riseAt",B20:"mass",B21:"cost"};
  for(const [cell,key] of Object.entries(bk)) chk(c,i,'BS '+cell,X['Busbar Sizer!'+cell],b[key]);
  b.sweep.forEach((s,j)=>{const r0=26+j; [['B','t'],['C','w'],['D','area'],['E','density'],['F','mass'],['G','cost'],['H','fits']].forEach(([col,key])=>chk(c,i,'BS '+col+r0,X['Busbar Sizer!'+col+r0],s[key]));});
  chk(c,i,'BS best',X['Busbar Sizer!B39'],b.best?b.best.ratio:'None fit'); chk(c,i,'BS bestcost',X['Busbar Sizer!B40'],b.best?b.best.cost:'-');
  console.log('case',i,'total',r.total.toFixed(4),'short',r.nShort,'base',r.baseL.toFixed(1)+'x'+r.baseW.toFixed(1)+'x'+r.height.toFixed(1),'cover',r.coverL.toFixed(1)+'x'+r.coverW.toFixed(1),'cu',r.cuTotal.toFixed(1),'best',b.best&&b.best.ratio);
});
console.log(`${n-bad}/${n} values match`); process.exit(bad?1:0);
