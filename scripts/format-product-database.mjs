import fs from 'node:fs/promises';
import {FileBlob,SpreadsheetFile} from '@oai/artifact-tool';

const inputPath='/Users/evanfoo/Documents/Codex/2026-09-07/in/outputs/data/TY_Product_Database.xlsx';
const tempPath='/tmp/ty-product-db-layout/TY_Product_Database.xlsx';
const previewDir='/tmp/ty-product-db-layout/final-previews';
const font='Arial';
const navy='#0B5C7E';
const blue='#DDEFF7';
const mint='#E8F4EF';
const yellow='#FFF4CC';
const border='#BDD9E5';
const body='#173B4B';

const workbook=await SpreadsheetFile.importXlsx(await FileBlob.load(inputPath));
const instructions=workbook.worksheets.getItem('Instructions');
const specifications=workbook.worksheets.getItem('Specifications');
const guide=workbook.worksheets.getItem('Product Guide');
const lineup=workbook.worksheets.getItem('Brand Lineup');
const exstock=workbook.worksheets.getItem('Ex-stock');

// Rebuild the opening sheet as a compact, beginner-friendly workflow.
for(const table of instructions.tables.items)table.delete();
instructions.getRange('A1:C20').unmerge();
instructions.getRange('A1:C20').clear({applyTo:'all'});
instructions.showGridLines=false;
instructions.getRange('A1:C14').format.font={name:font,size:10,color:body};
instructions.getRange('A1').values=[['TY Product Database']];
instructions.getRange('A1').format.font={name:font,size:16,bold:true,color:body};
instructions.getRange('A2').values=[['Start here: add or update products in Specifications. Brand Lineup is only for optional manual entries.']];
instructions.getRange('A2:C2').merge();
instructions.getRange('A2:C2').format={font:{name:font,size:10,italic:true,color:'#647B86'},wrapText:true};
instructions.getRange('A4:C4').values=[['Step','Where','What to do']];
instructions.getRange('A5:C9').values=[
  [1,'Specifications','Add or update the SKU specification and prices. Use the prepared yellow rows at the bottom for product types that are still empty.'],
  [2,'Product Guide','Edit only when changing the guide wording, lead time, questions, notes or category structure.'],
  [3,'Brand Lineup (optional)','Use only for a focus SKU without a completed specification, or when you need a manual lineup category.'],
  [4,'Ex-stock','Update stock, clearance price, retail price and image filename. SKU must be unique.'],
  [5,'Publish','Replace data/TY_Product_Database.xlsx on GitHub. Vercel rebuilds the website automatically.'],
];
instructions.getRange('A11:C11').values=[['Colour guide',null,null]];
instructions.getRange('A12:C13').values=[
  ['Blue cells','Already prepared','Keep the Category and Product Type unless you are intentionally changing the guide section.'],
  ['Yellow cells','Fill these in','Enter Brand, SKU and the available specification details.'],
];
instructions.getRange('A4:C4').format={fill:navy,font:{name:font,size:10,bold:true,color:'#FFFFFF'},horizontalAlignment:'center',verticalAlignment:'center',borders:{preset:'all',style:'thin',color:'#FFFFFF'}};
instructions.getRange('A5:C9').format={font:{name:font,size:10,color:body},verticalAlignment:'center',wrapText:true,borders:{insideHorizontal:{style:'thin',color:border},bottom:{style:'thin',color:border}}};
instructions.getRange('A5:A9').format.horizontalAlignment='center';
instructions.getRange('A5:B9').format.font={name:font,size:10,bold:true,color:body};
instructions.getRange('A11:C11').format={fill:'#D9EAF1',font:{name:font,size:10,bold:true,color:body},borders:{preset:'outside',style:'thin',color:border}};
instructions.getRange('A12:A12').format.fill=blue;
instructions.getRange('A13:A13').format.fill=yellow;
instructions.getRange('A12:C13').format.borders={preset:'all',style:'thin',color:border};
instructions.getRange('A1:C1').format.rowHeight=28;
instructions.getRange('A2:C2').format.rowHeight=32;
instructions.getRange('A4:C4').format.rowHeight=25;
instructions.getRange('A5:C9').format.rowHeight=40;
instructions.getRange('A11:C13').format.rowHeight=26;
instructions.getRange('A:A').format.columnWidth=18;
instructions.getRange('B:B').format.columnWidth=27;
instructions.getRange('C:C').format.columnWidth=92;
instructions.freezePanes.freezeRows(4);
instructions.tabColor='#0B5C7E';

// Read the authoritative category/product pairs from Product Guide.
const guideValues=guide.getUsedRange(true).values;
const guideHeaders=guideValues[0].map(v=>String(v??'').trim());
const guideCategoryIndex=guideHeaders.indexOf('Category');
const guideProductIndex=guideHeaders.indexOf('Product');
const guidePairs=guideValues.slice(1)
  .map(row=>({category:String(row[guideCategoryIndex]??'').trim(),product:String(row[guideProductIndex]??'').trim()}))
  .filter(row=>row.category&&row.product);

// Preserve every existing row, including recent entries added below the old Excel table.
const specValues=specifications.getUsedRange(true).values;
const specHeaders=specValues[0].map(v=>String(v??'').trim());
const categoryIndex=specHeaders.indexOf('Category');
const productIndex=specHeaders.indexOf('Product Type');
const skuIndex=specHeaders.indexOf('SKU');
for(let row=1;row<specValues.length;row++){
  if(String(specValues[row][productIndex]??'').trim()==='Ice makers'&&String(specValues[row][skuIndex]??'').trim()){
    specValues[row][categoryIndex]='Ice';
  }
}

// Add one ready-to-fill row for every Product Guide section with no specification yet.
const populatedProducts=new Set(specValues.slice(1)
  .filter(row=>String(row[skuIndex]??'').trim())
  .map(row=>String(row[productIndex]??'').trim()));
const missingPairs=guidePairs.filter(row=>!populatedProducts.has(row.product));
const headers=[...specHeaders,'Clearance Price','Retail Price'];
const existingRows=specValues.slice(1).map(row=>[...row.map(value=>value===''?null:value),null,null]);
const blankRows=missingPairs.map(row=>[row.category,row.product,null,null,null,null,null,null,null,null,null,null,null,null]);
const finalSpecValues=[headers,...existingRows,...blankRows];
const firstPreparedRow=existingRows.length+2;
const lastPreparedRow=finalSpecValues.length;
for(const table of specifications.tables.items)table.delete();
specifications.getRange(`A1:N${Math.max(lastPreparedRow,300)}`).clear({applyTo:'contents'});
specifications.getRange('A1').write(finalSpecValues);
const specTable=specifications.tables.add(`A1:N${lastPreparedRow}`,true,'SpecificationsTable');
specTable.style='TableStyleMedium2';
specifications.showGridLines=false;
specifications.freezePanes.freezeRows(1);
specifications.freezePanes.freezeColumns(2);
specifications.tabColor='#1686B0';
specifications.getRange('A1:N1').format={fill:navy,font:{name:font,size:10,bold:true,color:'#FFFFFF'},horizontalAlignment:'center',verticalAlignment:'center',wrapText:true,borders:{preset:'all',style:'thin',color:'#FFFFFF'},rowHeight:28};
specifications.getRange(`A2:N${lastPreparedRow}`).format.font={name:font,size:9,color:body};
specifications.getRange(`A2:N${lastPreparedRow}`).format.verticalAlignment='center';
specifications.getRange(`A${firstPreparedRow}:B${lastPreparedRow}`).format={fill:blue,font:{name:font,size:9,bold:true,color:'#0B5C7E'},verticalAlignment:'center',wrapText:true,borders:{preset:'all',style:'thin',color:border}};
specifications.getRange(`C${firstPreparedRow}:N${lastPreparedRow}`).format={fill:yellow,font:{name:font,size:9,color:body},verticalAlignment:'center',wrapText:true,borders:{preset:'all',style:'thin',color:'#E5D79B'}};
specifications.getRange(`A${firstPreparedRow}:N${lastPreparedRow}`).format.rowHeight=30;
specifications.getRange(`M2:N${lastPreparedRow}`).format.numberFormat='"$"#,##0.00';
const specWidths=[18,30,18,22,34,32,20,20,15,15,15,38,18,18];
for(let col=0;col<specWidths.length;col++)specifications.getRangeByIndexes(0,col,lastPreparedRow,1).format.columnWidth=specWidths[col];
const categories=[...new Set(guidePairs.map(row=>row.category))];
const products=guidePairs.map(row=>row.product);
specifications.dataValidations.add({range:`A2:A${lastPreparedRow+100}`,rule:{type:'list',values:categories}});
specifications.dataValidations.add({range:`B2:B${lastPreparedRow+100}`,rule:{type:'list',values:products}});
specifications.dataValidations.add({range:`K2:K${lastPreparedRow+100}`,rule:{type:'list',values:['Yes','No','-']}});

// Keep the supporting sheets consistent and easy to scan without changing their data.
for(const [sheet,columns,color] of [
  [exstock,10,'#4E8A67'],
  [guide,8,'#7B6AA8'],
  [lineup,5,'#C18432'],
]){
  const used=sheet.getUsedRange(true);
  const rowCount=used.values.length;
  sheet.showGridLines=false;
  sheet.freezePanes.freezeRows(1);
  sheet.getRangeByIndexes(0,0,1,columns).format={fill:navy,font:{name:font,size:10,bold:true,color:'#FFFFFF'},horizontalAlignment:'center',verticalAlignment:'center',wrapText:true,borders:{preset:'all',style:'thin',color:'#FFFFFF'},rowHeight:28};
  sheet.getRangeByIndexes(1,0,Math.max(1,rowCount-1),columns).format.font={name:font,size:9,color:body};
  sheet.getRangeByIndexes(1,0,Math.max(1,rowCount-1),columns).format.verticalAlignment='center';
  sheet.tabColor=color;
}
exstock.freezePanes.freezeColumns(2);
guide.freezePanes.freezeColumns(2);
lineup.freezePanes.freezeColumns(2);
const lineupRows=lineup.getUsedRange(true).values.length;
lineup.dataValidations.add({range:`A2:A${lineupRows+100}`,rule:{type:'list',values:categories}});
lineup.dataValidations.add({range:`B2:B${lineupRows+100}`,rule:{type:'list',values:products}});

workbook.recalculate();
await fs.mkdir(previewDir,{recursive:true});
for(const [sheetName,range] of [
  ['Instructions','A1:C13'],
  ['Specifications',`A${Math.max(1,firstPreparedRow-2)}:N${lastPreparedRow}`],
  ['Product Guide','A1:H18'],
  ['Brand Lineup','A1:E24'],
  ['Ex-stock','A1:J20'],
]){
  const preview=await workbook.render({sheetName,range,scale:1.2,format:'png'});
  await fs.writeFile(`${previewDir}/${sheetName.replaceAll(' ','-')}.png`,new Uint8Array(await preview.arrayBuffer()));
}

const output=await SpreadsheetFile.exportXlsx(workbook);
await output.save(tempPath);
console.log(JSON.stringify({missingPreparedRows:blankRows.length,firstPreparedRow,lastPreparedRow,tempPath}));
