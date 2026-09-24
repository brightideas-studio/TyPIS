import fs from 'node:fs/promises';
import path from 'node:path';
import readXlsxFile from 'read-excel-file/node';

const root=process.cwd();
const workbookPath=process.env.TY_PRODUCT_DATABASE_PATH||path.join(root,'data','TY_Product_Database.xlsx');
const sourcePath=path.join(root,'index.html');
const outputDir=path.join(root,'dist');

const text=value=>value===null||value===undefined?'':String(value).trim();
const lines=value=>text(value).split(/\r?\n/).map(v=>v.trim()).filter(Boolean);
const brandName=value=>/^ty innovations$/i.test(text(value))?'TY Innovations':text(value);
const number=value=>value===''||value===null||value===undefined?0:Number(value);
const price=value=>{
  if(value===''||value===null||value===undefined)return '';
  const amount=Number(String(value).replace(/[$,]/g,''));
  if(!Number.isFinite(amount))throw new Error(`Invalid price: ${value}`);
  return `$${amount.toLocaleString('en-US',{maximumFractionDigits:2})}`;
};

async function rowsFromSheet(name,requiredHeaders){
  let matrix;
  try{matrix=await readXlsxFile(workbookPath,{sheet:name})}
  catch(error){throw new Error(`Cannot read worksheet “${name}”: ${error.message}`)}
  const headers=(matrix[0]||[]).map(text);
  for(const header of requiredHeaders)if(!headers.includes(header))throw new Error(`${name}: missing column “${header}”`);
  const rows=[];
  for(let index=1;index<matrix.length;index++){
    const record={};
    headers.forEach((header,column)=>record[header]=matrix[index][column]);
    if(Object.values(record).some(value=>text(value)!==''))rows.push({...record,__row:index+1});
  }
  return rows;
}

function ensureUnique(rows,key,sheet,allowedDuplicates=new Set()){
  const seen=new Map();
  for(const row of rows){
    const value=text(row[key]);
    if(!value)throw new Error(`${sheet}: ${key} is blank at row ${row.__row}`);
    if(seen.has(value)&&!allowedDuplicates.has(value))throw new Error(`${sheet}: duplicate ${key} “${value}” at rows ${seen.get(value)} and ${row.__row}`);
    seen.set(value,row.__row);
  }
}

function replaceLiteral(source,name,value){
  const marker=`const ${name}=`;
  const markerIndex=source.indexOf(marker);
  if(markerIndex<0)throw new Error(`index.html is missing ${name}`);
  let start=markerIndex+marker.length;
  while(/\s/.test(source[start]))start++;
  const open=source[start],close=open==='['?']':open==='{'?'}':null;
  if(!close)throw new Error(`Cannot replace ${name}`);
  let depth=0,quote='',escaped=false,end=-1;
  for(let i=start;i<source.length;i++){
    const char=source[i];
    if(quote){if(escaped)escaped=false;else if(char==='\\')escaped=true;else if(char===quote)quote='';continue}
    if(char==='"'||char==="'"||char==='`'){quote=char;continue}
    if(char===open)depth++;
    if(char===close&&--depth===0){end=i+1;break}
  }
  if(end<0)throw new Error(`Unclosed ${name} data`);
  return source.slice(0,start)+JSON.stringify(value)+source.slice(end);
}

const exRows=await rowsFromSheet('Ex-stock',['SKU','Series','Description','Physical Qty','Available Qty','Clearance Price','Retail Price','Image','Source Page','Lifecycle Status']);
ensureUnique(exRows,'SKU','Ex-stock',new Set(['SKU NOT LISTED']));
const stockItems=exRows.map(row=>({
  section:text(row.Series),sku:text(row.SKU),description:text(row.Description),
  physicalQty:number(row['Physical Qty']),availableQty:number(row['Available Qty']),
  dealerPrice:price(row['Clearance Price']),retailPrice:price(row['Retail Price']),
  sourcePage:number(row['Source Page']),lifecycleStatus:text(row['Lifecycle Status']),image:text(row.Image)?`product-images/${text(row.Image).replace(/^product-images\//,'')}`:''
}));

const specificationSheetRows=await rowsFromSheet('Specifications',['Category','Product Type','Brand','SKU','Product Name','Dimension','Capacity','Temperature','Refrigerant','Energy Rating','EEG Claimable','Power','Clearance Price','Retail Price','Lineup Category','Show in Product Guide','Remark']);
// Brand section headings and prepared blank entry rows intentionally have no SKU.
const specificationRows=specificationSheetRows.filter(row=>text(row.SKU)!=='');
const showInProductGuide=row=>text(row['Show in Product Guide']).toLowerCase()!=='no';
ensureUnique(specificationRows,'SKU','Specifications');
const productSpecifications=Object.fromEntries(specificationRows.map(row=>[text(row.SKU),{
  name:text(row['Product Name']),dimension:text(row.Dimension),capacity:text(row.Capacity),temperature:text(row.Temperature),
  refrigerant:text(row.Refrigerant),energyRating:text(row['Energy Rating']),eegClaimable:text(row['EEG Claimable']),power:text(row.Power),
  clearancePrice:price(row['Clearance Price']),retailPrice:price(row['Retail Price']),
  category:text(row.Category),productType:text(row['Product Type']),brand:text(row.Brand),remark:text(row.Remark)
}]));

const guideRows=await rowsFromSheet('Product Guide',['Category','Product','Brands & Series','Lead Time','Our Focus','Images','Questions to Ask','Important Notes']);
const guideItems=guideRows.map(row=>({
  category:text(row.Category),name:text(row.Product),brands:text(row['Brands & Series']),lead:text(row['Lead Time']),focus:text(row['Our Focus']),
  images:lines(row.Images).map(entry=>{const [file,...label]=entry.split('|');return {src:`product-images/${file.trim().replace(/^product-images\//,'')}`,label:label.join('|').trim()}}),
  questions:lines(row['Questions to Ask']),notes:lines(row['Important Notes'])
}));
const guideCategoryByProduct=new Map(guideItems.map(item=>[item.name.toLowerCase(),item.category]));
for(const row of specificationRows){
  const productType=text(row['Product Type']),category=text(row.Category),expectedCategory=guideCategoryByProduct.get(productType.toLowerCase());
  if(!expectedCategory)throw new Error(`Specifications: Product Type “${productType}” at row ${row.__row} does not match a Product Guide product`);
  if(category!==expectedCategory)throw new Error(`Specifications: Category “${category}” at row ${row.__row} should be “${expectedCategory}” for ${productType}`);
}

const lineupSheetRows=await rowsFromSheet('Brand Lineup',['Main Category','Product Guide Section','Brand','Lineup Category','SKU']);
const lineupRows=lineupSheetRows.filter(row=>text(row.SKU)!=='');
const productLineups={};
const specificationProductBySku=new Map(specificationRows.map(row=>[text(row.SKU),text(row['Product Type'])]));
const specificationVisibilityBySku=new Map(specificationRows.map(row=>[text(row.SKU),showInProductGuide(row)]));
for(const row of lineupRows){
  const mainCategory=text(row['Main Category']),product=text(row['Product Guide Section']),type=text(row['Lineup Category']),brand=brandName(row.Brand),sku=text(row.SKU);
  if(!mainCategory||!product||!type||!brand||!sku)throw new Error(`Brand Lineup: Main Category, Product Guide Section, Brand, Lineup Category and SKU are required at row ${row.__row}`);
  const expectedCategory=guideCategoryByProduct.get(product.toLowerCase());
  if(!expectedCategory)throw new Error(`Brand Lineup: Product Guide Section “${product}” at row ${row.__row} is missing from Product Guide`);
  if(mainCategory!==expectedCategory)throw new Error(`Brand Lineup: Main Category “${mainCategory}” at row ${row.__row} should be “${expectedCategory}” for ${product}`);
  // When a completed specification exists, its Product Guide section remains
  // authoritative so a focus SKU cannot accidentally appear in two sections.
  if(specificationProductBySku.has(sku)&&specificationProductBySku.get(sku)!==product)continue;
  if(specificationVisibilityBySku.get(sku)===false)continue;
  const key=`${product}||${brand}`;
  productLineups[key]??=[];
  if(!productLineups[key].some(item=>item.sku===sku))productLineups[key].push({sku,type});
}
// Every completed specification appears automatically in Product Guide.
// Brand Lineup remains available for manual-only SKUs and category overrides.
const automaticLineupType=row=>{
  const product=text(row['Product Type']).toLowerCase(),name=text(row['Product Name']).toLowerCase();
  if(product.includes('refrigerator')){
    if(name.includes('combination'))return 'Combination';
    if(name.includes('freezer'))return 'Freezer';
    return 'Chiller';
  }
  return text(row['Product Name'])||text(row['Product Type']);
};
for(const row of specificationRows){
  if(!showInProductGuide(row))continue;
  const product=text(row['Product Type']),brand=brandName(row.Brand),sku=text(row.SKU),type=text(row['Lineup Category'])||automaticLineupType(row);
  if(!product||!brand||!sku)continue;
  const key=`${product}||${brand}`;
  productLineups[key]??=[];
  const existing=productLineups[key].find(item=>item.sku===sku);
  if(existing)existing.type=type;
  else productLineups[key].push({sku,type});
}

let html=await fs.readFile(sourcePath,'utf8');
html=replaceLiteral(html,'stockItems',stockItems);
html=replaceLiteral(html,'productSpecifications',productSpecifications);
html=replaceLiteral(html,'guideItems',guideItems);
html=replaceLiteral(html,'productLineups',productLineups);
html=html.replace(/\d+ stock records · Updated/g,`${stockItems.length} stock records · Updated`);
html=html.replace(/Last updated \d{1,2} (?:Sep|September) 2026<\/footer>/g,'Last updated 24 Sep 2026</footer>');
html=html.replace(/Stock dated \d{1,2} (?:Sep|September) 2026/g,'Stock dated 24 Sep 2026');
html=html.replace(/Internal clearance price list · Last updated \d{1,2} (?:Sep|September) 2026/g,'Internal clearance price list · Last updated 24 September 2026');

await fs.rm(outputDir,{recursive:true,force:true});
await fs.mkdir(outputDir,{recursive:true});
await fs.writeFile(path.join(outputDir,'index.html'),html);
// Keep the root index synchronized as well, so opening or uploading index.html
// directly always uses the latest workbook data.
await fs.writeFile(sourcePath,html);
for(const directory of ['product-images','brand-logos']){
  await fs.cp(path.join(root,directory),path.join(outputDir,directory),{recursive:true});
}
for(const file of ['ty-export-logo-data.js','ty-logo.png','ty-os-logo.png','ty-equipment-wiki-social-preview.png','landing-stock-background.png','product-guide-hero-kitchen.png']){
  try{await fs.copyFile(path.join(root,file),path.join(outputDir,file))}catch(error){if(error.code!=='ENOENT')throw error}
}
const linkedSpecificationCount=new Set(Object.values(productLineups).flat().map(item=>item.sku).filter(sku=>productSpecifications[sku])).size;
console.log(`Website built: ${stockItems.length} stock records, ${specificationRows.length} specifications, ${linkedSpecificationCount} focus SKUs linked, ${guideItems.length} guide entries, ${lineupRows.length} manual lineup rows.`);
