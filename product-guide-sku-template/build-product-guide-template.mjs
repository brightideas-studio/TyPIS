import fs from "node:fs/promises";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const outputDir = new URL(".", import.meta.url).pathname;
const outputPath = `${outputDir}TY_Product_Guide_SKU_Lineup_Template.xlsx`;
const font = "Arial";

const sections = [
  "Upright refrigerators", "Undercounter refrigerators", "Blast freezers", "Ice makers",
  "Chest freezers", "Cake and pastry displays", "Flower displays", "Glass-door displays",
  "Undercounter displays", "Tabletop displays", "Combi ovens", "Convection ovens",
  "Speed ovens", "Cookers and ranges", "Dishwashers", "Gelato equipment"
];

const brands = [
  "Panasonic", "Daiwa", "YODA", "TY Innovations", "Nuovair", "Rational", "Eloma",
  "UNOX", "Electrolux", "Hatco", "ISA", "Valmar", "Gelmatic", "Other"
];

const starterRows = [
  ["Upright refrigerators", "Panasonic", "", ""],
  ["Upright refrigerators", "Daiwa", "", ""],
  ["Upright refrigerators", "YODA", "", ""],
  ["Upright refrigerators", "TY Innovations", "", ""],
  ["Undercounter refrigerators", "Panasonic", "", ""],
  ["Undercounter refrigerators", "Daiwa", "", ""],
  ["Undercounter refrigerators", "YODA", "", ""],
  ["Undercounter refrigerators", "TY Innovations", "", ""],
  ["Blast freezers", "Nuovair", "", ""],
  ["Blast freezers", "Daiwa", "", ""],
  ["Ice makers", "Panasonic", "", ""],
  ["Ice makers", "Daiwa", "", ""],
  ["Chest freezers", "YODA", "", ""],
  ["Chest freezers", "TY Innovations", "", ""],
  ["Cake and pastry displays", "YODA", "", ""],
  ["Cake and pastry displays", "TY Innovations", "", ""],
  ["Flower displays", "YODA", "", ""],
  ["Flower displays", "TY Innovations", "", ""],
  ["Glass-door displays", "Panasonic", "", ""],
  ["Glass-door displays", "Daiwa", "", ""],
  ["Glass-door displays", "YODA", "", ""],
  ["Glass-door displays", "TY Innovations", "", ""]
];

const workbook = Workbook.create();
const lineup = workbook.worksheets.add("SKU Lineup");
const logo = workbook.worksheets.add("Brand Logos");
const reference = workbook.worksheets.add("Reference");

lineup.showGridLines = false;
lineup.getRange("A2:D2").merge();
lineup.getRange("A2").values = [["TY Product Guide SKU Lineup"]];
lineup.getRange("A2:D2").format.font = { name: font, size: 16, bold: true, color: "#0A5A82" };
lineup.getRange("A3:D3").merge();
lineup.getRange("A3").values = [["Enter one SKU per row. Repeat the Product Guide Section and Brand for every SKU."]];
lineup.getRange("A3:D3").format.font = { name: font, size: 10, italic: true, color: "#667985" };
lineup.getRange("A5:D5").values = [["Product Guide Section", "Brand", "SKU", "Notes (optional)"]];
lineup.getRange("A5:D5").format = { fill: "#0A5A82", font: { name: font, size: 10, bold: true, color: "#FFFFFF" }, horizontalAlignment: "center", verticalAlignment: "center" };

const rows = Array.from({ length: 300 }, (_, index) => starterRows[index] || ["", "", "", ""]);
lineup.getRange("A6:D305").values = rows;
lineup.getRange("A6:D305").format.font = { name: font, size: 10, color: "#203844" };
lineup.getRange("A6:D305").format.verticalAlignment = "center";
lineup.getRange("A6:C305").format.fill = "#FFF8DF";
lineup.getRange("A6:D305").format.borders = { preset: "inside", style: "thin", color: "#E2E8EB" };
lineup.getRange("A6:A305").dataValidation = { rule: { type: "list", formula1: "Reference!$A$2:$A$17" } };
lineup.getRange("B6:B305").dataValidation = { rule: { type: "list", formula1: "Reference!$B$2:$B$15" } };
lineup.getRange("A:A").format.columnWidth = 31;
lineup.getRange("B:B").format.columnWidth = 20;
lineup.getRange("C:C").format.columnWidth = 27;
lineup.getRange("D:D").format.columnWidth = 34;
lineup.getRange("5:5").format.rowHeight = 30;
lineup.getRange("6:305").format.rowHeight = 23;
lineup.freezePanes.freezeRows(5);
lineup.tables.add("A5:D305", true, "SkuLineupTable").style = "TableStyleMedium2";

logo.showGridLines = false;
logo.getRange("A2:C2").merge();
logo.getRange("A2").values = [["Brand Logo File Map"]];
logo.getRange("A2:C2").format.font = { name: font, size: 16, bold: true, color: "#0A5A82" };
logo.getRange("A3:C3").merge();
logo.getRange("A3").values = [["Upload the logo files together with this workbook. Transparent PNG or SVG is preferred."]];
logo.getRange("A3:C3").format.font = { name: font, size: 10, italic: true, color: "#667985" };
logo.getRange("A5:C5").values = [["Brand", "Logo filename", "Notes (optional)"]];
logo.getRange("A5:C5").format = { fill: "#0A5A82", font: { name: font, size: 10, bold: true, color: "#FFFFFF" }, horizontalAlignment: "center", verticalAlignment: "center" };
logo.getRange("A6:C19").values = brands.map(brand => [brand, "", ""]);
logo.getRange("A6:B19").format.fill = "#FFF8DF";
logo.getRange("A6:C19").format.font = { name: font, size: 10, color: "#203844" };
logo.getRange("A6:C19").format.borders = { preset: "inside", style: "thin", color: "#E2E8EB" };
logo.getRange("A:A").format.columnWidth = 22;
logo.getRange("B:B").format.columnWidth = 34;
logo.getRange("C:C").format.columnWidth = 34;
logo.getRange("5:5").format.rowHeight = 30;
logo.getRange("6:19").format.rowHeight = 23;
logo.freezePanes.freezeRows(5);
logo.tables.add("A5:C19", true, "BrandLogoTable").style = "TableStyleMedium2";

reference.showGridLines = false;
reference.getRange("A1:B1").values = [["Product Guide Sections", "Brands"]];
reference.getRange("A1:B1").format = { fill: "#0A5A82", font: { name: font, size: 10, bold: true, color: "#FFFFFF" } };
reference.getRange(`A2:A${sections.length + 1}`).values = sections.map(value => [value]);
reference.getRange(`B2:B${brands.length + 1}`).values = brands.map(value => [value]);
reference.getRange("A1:B20").format.font = { name: font, size: 10 };
reference.getRange("A:A").format.columnWidth = 31;
reference.getRange("B:B").format.columnWidth = 22;
reference.tabColor = "#B7C8D0";

workbook.recalculate();
const check = await workbook.inspect({ kind: "table", range: "SKU Lineup!A2:D30", include: "values,formulas", tableMaxRows: 30, tableMaxCols: 4, maxChars: 8000 });
console.log(check.ndjson);
const errors = await workbook.inspect({ kind: "match", searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A|#NUM!|#NULL!|#SPILL!|#CALC!", options: { useRegex: true, maxResults: 100 }, summary: "final formula error scan" });
console.log(errors.ndjson);
const preview = await workbook.render({ sheetName: "SKU Lineup", range: "A1:D30", scale: 1.5, format: "png" });
await fs.writeFile(`${outputDir}TY_Product_Guide_SKU_Lineup_Template_preview.png`, new Uint8Array(await preview.arrayBuffer()));
const output = await SpreadsheetFile.exportXlsx(workbook);
await output.save(outputPath);
console.log(outputPath);
