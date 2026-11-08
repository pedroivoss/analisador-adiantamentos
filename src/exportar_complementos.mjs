import fs from 'node:fs/promises';
import path from 'node:path';
import {Workbook,SpreadsheetFile} from '@oai/artifact-tool';
const run=process.argv[2];
const p=JSON.parse(await fs.readFile(path.join(run,'dados/complementos.json'),'utf8'));
const wb=Workbook.create();
// A pasta de QA fica fora das entregas.
const qaDir=path.resolve(import.meta.dirname,'..','work',path.basename(run),'complementos');
await fs.mkdir(qaDir,{recursive:true});
const col=n=>String.fromCharCode(64+n);
const formatMoney='#,##0.00;(#,##0.00);"-"';
function sheet(name,title,headers,rows,widths,moneycols=[]){
 const s=wb.worksheets.add(name);s.showGridLines=false;const end=col(headers.length);
 s.getRange('A2').values=[[title]];s.getRange('A2').format.font={name:'Arial',size:14,bold:true};
 s.mergeCells(`A3:${end}3`);s.getRange('A3').values=[['Fonte: controle de adiantamentos do Município de Ubatuba — janeiro a junho de 2026']];s.getRange(`A3:${end}3`).format.font={name:'Arial',size:9,color:'#555555'};
 const g=s.getRangeByIndexes(4,0,rows.length+1,headers.length);g.values=[headers,...rows];g.format.font={name:'Arial',size:10,color:'#000000'};g.format.verticalAlignment='center';g.format.rowHeight=30;
 const h=s.getRange(`A5:${end}5`);h.format.fill='#DDDDDD';h.format.font.bold=true;h.format.wrapText=true;h.format.rowHeight=42;h.format.horizontalAlignment='center';
 for(let j=0;j<headers.length;j++){s.getRangeByIndexes(0,j,rows.length+6,1).format.columnWidthPx=widths[j];if(moneycols.includes(j)){s.getRangeByIndexes(5,j,rows.length,1).setNumberFormat(formatMoney);s.getRangeByIndexes(5,j,rows.length,1).format.horizontalAlignment='right';}}
 for(let i=0;i<rows.length;i++)if(i%2===1)s.getRangeByIndexes(i+5,0,1,headers.length).format.fill='#F5F5F5';
 s.freezePanes.freezeRows(5);return s;
}
function comments(s,texts,start,last){
 for(let i=0;i<texts.length;i++){
  const r=start+i*3;s.mergeCells(`A${r}:${last}${r+1}`);s.getRange(`A${r}`).values=[[texts[i]]];const rr=s.getRange(`A${r}:${last}${r+1}`);rr.format.font={name:'Arial',size:10,color:'#000000'};rr.format.wrapText=true;rr.format.verticalAlignment='center';rr.format.rowHeight=31;
 }
}
for(const m of p.meses){
 const s=sheet(m.rotulo.replace('/',' '),m.titulo,['Sigla','Secretaria','Adiantado (R$)','Utilizado (R$)','Devolvido (R$)'],m.linhas,[65,350,140,140,140],[2,3,4]);s.getRange(`B6:B${m.linhas.length+5}`).format.wrapText=true;s.getRange(`A6:E${m.linhas.length+5}`).format.rowHeight=36;
 const row=m.linhas.length+6;s.getRange(`A${row}`).values=[['Total']];s.getRange(`C${row}:E${row}`).formulas=[[`=SUM(C6:C${row-1})`,`=SUM(D6:D${row-1})`,`=SUM(E6:E${row-1})`]];s.getRange(`C${row}:E${row}`).setNumberFormat(formatMoney);s.getRange(`A${row}:E${row}`).format.font.bold=true;s.getRange(`A${row}:E${row}`).format.fill='#E5E5E5';
 comments(s,m.comentarios,row+3,'E');
}
const pres=sheet('Presença mensal','Presença mensal das secretarias',['Secretaria',...p.meses.map(m=>m.rotulo),'Meses com adiantamento'],p.presenca,[110,90,90,90,90,90,90,155]);
for(let i=0;i<p.presenca.length;i++)pres.getRange(`H${i+6}`).formulas=[[`=COUNTIF(B${i+6}:G${i+6},"Com")`]];
comments(pres,p.comentarios_presenca,p.presenca.length+9,'H');
const detailWidths=[165,45,35,70,70,70,35,65,65,65,80,60,60];
const d=sheet('Relação individual','Relação individual dos adiantamentos',p.campos_detalhe,p.detalhe,detailWidths,[3,4,5]);
d.getRange(`A6:M${p.detalhe.length+5}`).format.wrapText=true;d.getRange(`A6:M${p.detalhe.length+5}`).format.rowHeight=44;d.getRange('A5:M5').format.rowHeight=52;
for(let i=0;i<p.detalhe.length;i++){
 const lines=Math.max(...p.detalhe[i].map((v,j)=>Math.ceil(String(v??'').length/Math.max(1,Math.floor(detailWidths[j]/6.5)))));
 d.getRange(`A${i+6}:M${i+6}`).format.rowHeight=Math.max(44,lines*12+8);
}
const tab=d.tables.add(`A5:M${p.detalhe.length+5}`,true,'RelacaoAdiantamentos');tab.style='TableStyleLight1';tab.showBandedRows=false;
const total=p.detalhe.length+6;d.getRange(`A${total}`).values=[['Total']];d.getRange(`D${total}:F${total}`).formulas=[[`=SUM(D6:D${total-1})`,`=SUM(E6:E${total-1})`,`=SUM(F6:F${total-1})`]];d.getRange(`D${total}:F${total}`).setNumberFormat(formatMoney);d.getRange(`A${total}:M${total}`).format.font.bold=true;
comments(d,p.comentarios_detalhe,total+3,'M');
wb.recalculate();
console.log((await wb.inspect({kind:'region',sheetId:'jan 26',range:'C27:E27',maxChars:1000,tableMaxRows:2,tableMaxCols:3})).ndjson);
console.log((await wb.inspect({kind:'match',searchTerm:'#REF!|#DIV/0!|#VALUE!|#NAME\\?|#NUM!|#SPILL!|#CALC!',options:{useRegex:true,maxResults:20},summary:'Erros de fórmula'})).ndjson);
for(const name of [...p.meses.map(m=>m.rotulo.replace('/',' ')),'Presença mensal','Relação individual']){
 const last=name==='Relação individual'?'M':name==='Presença mensal'?'H':'E';const bottom=name==='Relação individual'?18:27;
 const im=await wb.render({sheetName:name,range:`A2:${last}${bottom}`,scale:1.2,format:'png'});await fs.writeFile(path.join(qaDir,name+'.png'),new Uint8Array(await im.arrayBuffer()));
}
const file=await SpreadsheetFile.exportXlsx(wb);await file.save(path.join(run,'Planilhas','Analise mensal e detalhada.xlsx'));
console.log('Planilha complementar: oito abas exportadas.');
