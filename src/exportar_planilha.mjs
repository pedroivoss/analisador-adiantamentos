import fs from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { Workbook, SpreadsheetFile } from '@oai/artifact-tool';
const run=process.argv[2];
const payload=JSON.parse(await fs.readFile(path.join(run,'dados/workbook_payload.json'),'utf8'));
const wb=Workbook.create();
const qaDir=path.resolve(path.dirname(fileURLToPath(import.meta.url)),'..','work',path.basename(run));
await fs.mkdir(qaDir,{recursive:true});
const summaries=Object.entries(payload.summary).filter(([k,v])=>typeof v!=='object').map(([indicador,valor])=>({indicador,valor}));
const collections={Resumo:summaries,...payload.tables,'Base normalizada':payload.rows};
const currency=new Set(['adiantado','utilizado','devolvido','media_adiantado','diferenca_financeira','consolidado','recalculado','diferenca']);
let tableCount=0;
function columnName(n){let s='';while(n){n--;s=String.fromCharCode(65+n%26)+s;n=Math.floor(n/26);}return s;}
for(const [name,records0] of Object.entries(collections)){
  const records=records0.length?records0:[{resultado:'Sem registros elegíveis'}];
  const headers=Object.keys(records[0]);
  const sheet=wb.worksheets.add(name.slice(0,31));
  sheet.showGridLines=false;
  sheet.getRange('A2').values=[[name==='Resumo'?'Adiantamentos - resultados da execução':name]];
  sheet.getRange('A2').format.font={name:'Arial',size:14,bold:true};
  const values=[headers,...records.map(r=>headers.map(k=>{
    const value=r[k];
    if(typeof value==='string'&&/^[=+@-]/.test(value))return "'"+value;
    if(typeof value==='object'&&value!==null)return JSON.stringify(value);
    return value??null;
  }))];
  const range=sheet.getRangeByIndexes(3,0,values.length,headers.length);range.values=values;
  range.format.font={name:'Arial',size:10};range.format.rowHeight=20;range.format.columnWidth=20;
  const header=sheet.getRangeByIndexes(3,0,1,headers.length);header.format.fill='#24476B';header.format.font={bold:true,color:'#FFFFFF',name:'Arial',size:10};header.format.wrapText=true;header.format.rowHeight=34;
  sheet.getRangeByIndexes(4,0,records.length,1).format.columnWidth=name==='Resumo'?39:24;
  for(let j=0;j<headers.length;j++){
    const h=headers[j];const col=sheet.getRangeByIndexes(4,j,records.length,1);
    if(currency.has(h))col.setNumberFormat('"R$" #,##0.00');
    if(h.includes('percentual')||h.startsWith('taxa_'))col.setNumberFormat('0.00%');
    if(['servidor','tipo_original','descricao'].includes(h))col.format.columnWidth=h==='descricao'?100:55;
    if(h==='processo')col.format.columnWidth=37;
    if(h.includes('prazo')||h.includes('dias_'))col.format.columnWidth=26;
  }
  sheet.freezePanes.freezeRows(4);
  sheet.tables.add('A4:'+columnName(headers.length)+(values.length+3),true,'Tabela'+(++tableCount));
  if(name==='Resumo'){
    for(let i=0;i<records.length;i++)if(currency.has(records[i].indicador))sheet.getRangeByIndexes(i+4,1,1,1).setNumberFormat('"R$" #,##0.00');
  }
}
wb.recalculate();
console.log((await wb.inspect({kind:'region',sheetId:'Resumo',range:'A4:B16',maxChars:2500,tableMaxRows:12,tableMaxCols:2})).ndjson);
const preview=await wb.render({sheetName:'Resumo',range:'A2:B17',scale:1.4,format:'png'});
await fs.writeFile(path.join(qaDir,'preview_planilha.png'),new Uint8Array(await preview.arrayBuffer()));
const output=await SpreadsheetFile.exportXlsx(wb);await output.save(path.join(run,'Planilhas/Analise_editavel.xlsx'));
try { await fs.rename(path.join(run,'Planilhas/Analise_editavel.xlsx.inspect.ndjson'),path.join(qaDir,'inspect_planilha.ndjson')); } catch(e) { if(e.code!=='ENOENT')throw e; }
console.log('Planilha exportada com',Object.keys(collections).length,'abas.');
