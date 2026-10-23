import {build} from 'esbuild';
import {mkdir,readFile,writeFile} from 'node:fs/promises';
await mkdir('vendor',{recursive:true});
for(const [entry,name,globalName] of [['src/mail-parsers.js','mail-parsers','RBD_MAIL_PARSERS'],['src/entra-auth.js','entra-auth','RBD_ENTRA']])await build({entryPoints:[entry],outfile:`vendor/${name}.js`,bundle:true,format:'iife',globalName,platform:'browser',target:['chrome109','safari16'],minify:true,legalComments:'eof'});
const packages=['postal-mime','@kenjiuno/msgreader','dompurify','@azure/msal-browser','@azure/msal-common','iconv-lite','safer-buffer','@kenjiuno/decompressrtf','@kenjiuno/rtf-parser','@kenjiuno/rtf-parser'];
let notices='# Third-party licenses\n\nBrowser bundles are built locally with pinned dependencies; no runtime CDN imports.\n';
for(const pkg of new Set(packages)){try{const p=JSON.parse(await readFile(`node_modules/${pkg}/package.json`));let license='';for(const file of ['LICENSE','LICENSE.txt','LICENSE.md','LICENSE.TXT']){try{license=await readFile(`node_modules/${pkg}/${file}`,'utf8');break;}catch{}}notices+=`\n## ${pkg} ${p.version} (${p.license||'see license'})\n\n${license}\n`;}catch{}}
await writeFile('vendor/THIRD_PARTY_LICENSES.md',notices);
