import fs from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { Workbook, SpreadsheetFile } from '@oai/artifact-tool';
const run=process.argv[2];
const payload=JSON.parse(await fs.readFile(path.join(run,'dados/apresentacao.json'),'utf8'));
const wb=Workbook.create();
const qaDir=path.resolve(path.dirname(fileURLToPath(import.meta.url)),'..','work',path.basename(run));
await fs.mkdir(qaDir,{recursive:true});
function col(n){let s='';while(n){n--;s=String.fromCharCode(65+n%26)+s;n=Math.floor(n/26);}return s;}
const formats={moeda:'#,##0.00',percentual:'0.00%',decimal:'0.0',inteiro:'0'};
let count=0;
for(const plan of payload.quadros){
  const sheet=wb.worksheets.add(plan.id.slice(0,31));sheet.showGridLines=false;
  const last=col(plan.cabecalhos.length),n=plan.linhas.length;
  sheet.mergeCells(`A2:${last}2`);sheet.getRange('A2').values=[[plan.titulo]];sheet.getRange(`A2:${last}2`).format.font={name:'Arial',size:14,bold:true};
  sheet.mergeCells(`A3:${last}3`);sheet.getRange('A3').values=[['Adiantamentos - '+payload.periodo.inicio+' a '+payload.periodo.fim]];sheet.getRange(`A3:${last}3`).format.font={name:'Arial',size:9,color:'#555555'};
  const matrix=[plan.cabecalhos,...plan.linhas];
  const range=sheet.getRangeByIndexes(4,0,matrix.length,plan.cabecalhos.length);range.values=matrix;
  range.format.font={name:'Arial',size:10};range.format.rowHeight=25;range.format.verticalAlignment='center';
  for(let j=0;j<plan.cabecalhos.length;j++){
    sheet.getRangeByIndexes(0,j,n+10,1).format.columnWidthPx=plan.larguras[j]*96/72;
    if(formats[plan.formatos[j]])sheet.getRangeByIndexes(5,j,n,1).setNumberFormat(formats[plan.formatos[j]]);
    if(j>0)sheet.getRangeByIndexes(5,j,n,1).format.horizontalAlignment='right';
  }
  const header=sheet.getRange(`A5:${last}5`);header.format.fill='#24476B';header.format.font={name:'Arial',size:10,color:'#FFFFFF',bold:true};header.format.wrapText=true;header.format.rowHeight=38;header.format.horizontalAlignment='center';
  if(plan.id==='Resumo')sheet.getRange('B10:B13').setNumberFormat('#,##0.00');
  for(let i=0;i<n;i++)if(i%2===1)sheet.getRangeByIndexes(i+5,0,1,plan.cabecalhos.length).format.fill='#F2F5F8';
  if(plan.tem_total){const total=sheet.getRangeByIndexes(n+4,0,1,plan.cabecalhos.length);total.format.fill='#E2EAF1';total.format.font.bold=true;}
  const noteRow=n+7;sheet.mergeCells(`A${noteRow}:${last}${noteRow}`);sheet.getRange(`A${noteRow}`).values=[['Fonte: '+payload.fonte+' | '+payload.aba]];
  sheet.getRange(`A${noteRow}:${last}${noteRow}`).format.font={name:'Arial',size:8,color:'#555555'};sheet.getRange(`A${noteRow}:${last}${noteRow}`).format.wrapText=true;sheet.getRange(`A${noteRow}:${last}${noteRow}`).format.rowHeight=28;
  sheet.freezePanes.freezeRows(5);
  for(const [name,offset] of [['O que mostra',3],['Leitura dos resultados',5],['Critério',8]]){
    const r=noteRow+offset,key={'O que mostra':'o_que_mostra','Leitura dos resultados':'conclusao','Critério':'criterio'}[name];
    sheet.mergeCells(`A${r}:${last}${r}`);sheet.getRange(`A${r}`).values=[[name+': '+plan[key]]];const rr=sheet.getRange(`A${r}:${last}${r}`);rr.format.font={name:'Arial',size:9};rr.format.wrapText=true;rr.format.rowHeight=offset===5?55:44;
  }
  // Cada aba é um quadro de no máximo seis colunas, pronto para copiar.
  count++;
}
wb.recalculate();
console.log((await wb.inspect({kind:'region',sheetId:'Mensal',range:'A5:D12',maxChars:1800,tableMaxRows:8,tableMaxCols:4})).ndjson);
for(const name of ['Mensal','Secretarias','Maiores valores','Durações']){
 const p=payload.quadros.find(x=>x.id===name),last=col(p.cabecalhos.length);
 const png=await wb.render({sheetName:name,range:`A2:${last}${p.linhas.length+7}`,scale:1.4,format:'png'});
 await fs.writeFile(path.join(qaDir,`preview ${name}.png`),new Uint8Array(await png.arrayBuffer()));
}
const output=await SpreadsheetFile.exportXlsx(wb);await output.save(path.join(run,'Planilhas/Analise_editavel.xlsx'));
try{await fs.rename(path.join(run,'Planilhas/Analise_editavel.xlsx.inspect.ndjson'),path.join(qaDir,'inspect_planilha.ndjson'));}catch(e){if(e.code!=='ENOENT')throw e;}
console.log('Quadros compactos exportados:',count);
