import fs from 'node:fs';
import path from 'node:path';
const problems=[];let checked=0;
function scan(directory){
  for(const item of fs.readdirSync(directory,{withFileTypes:true})){
    const file=path.join(directory,item.name);
    if(item.isDirectory()){
      if(['.runtime','.git','node_modules'].includes(item.name))problems.push(file+': forbidden runtime directory');
      else scan(file);
    }else{
      checked++;
      if(/^\.env(?:\.|$)/.test(item.name)||item.name==='endpoint.json'||item.name==='office-session-manifest.xml')problems.push(file+': forbidden secret/session file');
      if(/\.(?:cs|js|mjs|json|xml|md|txt|ps1|cmd|html|css|config|trx|tap)$/i.test(item.name)){
        const bytes=fs.readFileSync(file);
        const value=bytes.toString(bytes[0]===255&&bytes[1]===254?'utf16le':'utf8');
        if(/\bsk-(?:proj-|svcacct-)?[A-Za-z0-9_-]{24,}/.test(value))problems.push(file+': possible API credential');
        if(/ll-auth\.[a-f0-9]{64}|#session=[a-f0-9]{64}|Bearer [a-f0-9]{64}/i.test(value))problems.push(file+': session credential');
        if(envKey&&value.includes(envKey))problems.push(file+': configured key detected');
      }
    }
  }
}
const envKey=process.env.OPENAI_API_KEY;
for(const directory of process.argv.slice(2))scan(path.resolve(directory));
if(problems.length){console.error(problems.join('\n'));process.exitCode=1}else console.log(`PASS: ${checked} packaged files inspected; no secret/session files or credential patterns found.`);
