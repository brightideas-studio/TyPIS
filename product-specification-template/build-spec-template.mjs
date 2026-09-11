import fs from "node:fs/promises";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const inputPath = "/Users/evanfoo/Documents/Codex/2026-09-07/in/outputs/index.html";
const outputDir = "/Users/evanfoo/Documents/Codex/2026-09-07/in/outputs/product-specification-template";
const outputPath = `${outputDir}/TY_Product_Specification_Template.xlsx`;
const previewPath = `${outputDir}/TY_Product_Specification_Template_preview.png`;

const html = await fs.readFile(inputPath, "utf8");
const match = html.match(/const stockItems=(\[.*?\]);\s*const stockCategories/s);
if (!match) throw new Error("Could not find stockItems in index.html");
const items = JSON.parse(match[1]);

function categoryFor(section) {
  const s = String(section).toUpperCase();
  if (s.includes("ICE MAKER")) return "Ice Makers";
  if (s.includes("DISHWASH")) return "Dishwasher";
  if (s === "ISA" || s.includes("GELATO") || s.includes("SOFT SERVE") || s.includes("VALMAR")) return "Gelato";
  if (s.includes("OCU") || s.includes("COMPRESSOR")) return "Components";
  if (s.includes("GLASS DOOR") || s.includes("SHOWCASE") || s.includes("SERVICE COUNTER") || s.includes("ISLAND CHEST") || s.includes("RTCK") || s.includes("RTOP")) return "Glass Door";
  if (s.includes("OVEN") || s.includes("RICE COOKER") || s.includes("COOKING") || s.includes("HOT RANGE") || s.includes("MARUZEN")) return "Cooking";
  if (s.includes("REFRIGERATOR") || s.includes("PANASONIC LV") || s.includes("PANASONIC KB") || s.includes("PANASONIC ES") || s.includes("PANASONIC HB") || s.includes("PANASONIC HP") || s.includes("DAIWA")) return "Refrigeration";
  return "Other";
}

const headers = [
  "SKU",
  "Product Name",
  "Category",
  "Source Section",
  "Record ID",
  "Dimension (W × D × H mm)",
  "Capacity",
  "Refrigerant",
  "Energy Rating",
  "EEG Claimable",
  "Power Supply",
  "Power Consumption",
  "Specification Notes",
  "Specification Source / URL",
];

const rows = items.map((item, index) => [
  item.sku,
  item.description,
  categoryFor(item.section),
  item.section,
  `TY-${String(index + 1).padStart(4, "0")}`,
  "",
  "",
  "",
  "",
  "",
  "",
  "",
  "",
  "",
]);

const workbook = Workbook.create();
const sheet = workbook.worksheets.add("Product Specifications");
sheet.showGridLines = false;
sheet.tabColor = "#0A5A82";

sheet.getRange("A2").values = [["TY Product Specification Template"]];
sheet.getRange("A3").values = [["Fill the yellow columns only. Keep SKU unchanged so the completed file can be matched back to the website. Use one value per cell and include units where applicable."]];
sheet.getRange("A4").values = [["EEG Claimable options: Yes, No, Pending, or Not applicable. Blank specification fields will be hidden in the website pop-up."]];
sheet.getRange("A5:N5").values = [headers];
sheet.getRange(`A6:N${rows.length + 5}`).values = rows;

const title = sheet.getRange("A2:N2");
title.format.font = { name: "Arial", size: 16, bold: true, color: "#0A5A82" };
title.format.rowHeight = 26;

const instructions = sheet.getRange("A3:N4");
instructions.format.font = { name: "Arial", size: 10, color: "#526873", italic: true };
instructions.format.rowHeight = 18;

const headersRange = sheet.getRange("A5:N5");
headersRange.format = {
  fill: "#0A5A82",
  font: { name: "Arial", size: 10, bold: true, color: "#FFFFFF" },
  horizontalAlignment: "center",
  verticalAlignment: "center",
  wrapText: true,
  borders: { preset: "inside", style: "thin", color: "#FFFFFF" },
};
headersRange.format.rowHeight = 42;

const body = sheet.getRange(`A6:N${rows.length + 5}`);
body.format.font = { name: "Arial", size: 10, color: "#20333D" };
body.format.verticalAlignment = "center";
body.format.rowHeight = 24;
body.format.borders = { preset: "inside", style: "thin", color: "#E4EAED" };

const sourceColumns = sheet.getRange(`A6:E${rows.length + 5}`);
sourceColumns.format.fill = "#F4F7F8";
const inputColumns = sheet.getRange(`F6:N${rows.length + 5}`);
inputColumns.format.fill = "#FFF7D6";

sheet.getRange(`A6:A${rows.length + 5}`).format.font = { name: "Arial", size: 10, bold: true, color: "#0A5A82" };
sheet.getRange(`J6:J${rows.length + 5}`).dataValidation = {
  rule: { type: "list", values: ["Yes", "No", "Pending", "Not applicable"] },
};

const table = sheet.tables.add(`A5:N${rows.length + 5}`, true, "ProductSpecificationsTable");
table.style = "TableStyleMedium2";
table.showFilterButton = true;
table.showBandedRows = false;

sourceColumns.format.fill = "#F4F7F8";
inputColumns.format.fill = "#FFF7D6";

const widths = [140, 310, 105, 190, 80, 175, 110, 105, 110, 115, 145, 135, 260, 230];
widths.forEach((width, index) => {
  sheet.getRangeByIndexes(0, index, rows.length + 5, 1).format.columnWidthPx = width;
});
sheet.getRange(`B6:B${rows.length + 5}`).format.wrapText = true;
sheet.getRange(`M6:N${rows.length + 5}`).format.wrapText = false;
sheet.freezePanes.freezeRows(5);
sheet.freezePanes.freezeColumns(2);

workbook.recalculate();

const inspection = await workbook.inspect({
  kind: "table",
  range: "Product Specifications!A2:N12",
  include: "values,formulas",
  tableMaxRows: 12,
  tableMaxCols: 14,
});
console.log(inspection.ndjson);

const errors = await workbook.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A|#NUM!|#NULL!|#SPILL!|#CALC!",
  options: { useRegex: true, maxResults: 100 },
  summary: "final formula error scan",
});
console.log(errors.ndjson);

const preview = await workbook.render({
  sheetName: "Product Specifications",
  range: "A1:N18",
  scale: 1.2,
  format: "png",
});
await fs.writeFile(previewPath, new Uint8Array(await preview.arrayBuffer()));

await fs.mkdir(outputDir, { recursive: true });
const output = await SpreadsheetFile.exportXlsx(workbook);
await output.save(outputPath);
console.log(JSON.stringify({ outputPath, previewPath, rowCount: rows.length }));
