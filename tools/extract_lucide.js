const fs=require('fs');const vm=require('vm');
const src=fs.readFileSync(process.argv[2],'utf8');
const sb={};sb.self=sb;sb.globalThis=sb;vm.runInNewContext(src,sb);
const ic=sb.lucide.icons||sb.lucide;
const kebab=s=>s.replace(/([a-zA-Z])(\d)/g,'$1-$2').replace(/([a-z])([A-Z])/g,'$1-$2').replace(/([A-Z])([A-Z][a-z])/g,'$1-$2').toLowerCase();
const out={};
for(const k of Object.keys(ic)){
  const node=ic[k];
  if(!Array.isArray(node)) continue;
  const pairs = Array.isArray(node[0]) ? node : (Array.isArray(node[2])?node[2]:[]);
  const parts=pairs.filter(p=>Array.isArray(p)&&typeof p[0]==='string').map(([tag,attrs])=>'<'+tag+' '+Object.entries(attrs||{}).filter(([a])=>a!=='key').map(([a,v])=>a+'="'+v+'"').join(' ')+'/>');
  if(parts.length) out[kebab(k)]=parts.join('');
}
fs.writeFileSync(process.argv[3],JSON.stringify(out));
console.log(Object.keys(out).length, ['circle-help','trash-2','file-text','a-arrow-up','grid-3x3','chart-column'].map(n=>n+':'+!!out[n]).join(' '));
