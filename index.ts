#!/usr/bin/env ts-node
import * as program from "commander";
import * as fs from "fs";
import * as xl from "excel4node";
const pack = require("./package");

let file: string;

program
  .version(pack.version)
  .arguments("<file>")
  .action(filePath => {
    file = fs.readFileSync(filePath).toString();
  })
  .parse(process.argv);

const entries = file.split("\n").map(entry => {
  const priceString = entry.substring(0, entry.indexOf(" "));
  const description = entry.substring(entry.indexOf(" ") + 1);
  const priceStringEnglishNotation = priceString.replace(",", ".");
  const price = Number.parseFloat(priceStringEnglishNotation);
  return { price, description }
});

const excelWorkbook = (() => {
  const wb = new xl.Workbook();
  const ws = wb.addWorksheet("Expenses");

  ws.cell(1, 1)
    .string("Name");
  ws.cell(1, 2)
    .string("Price");

  entries.forEach((entry, index) => {
    ws.cell(index + 2, 1)
      .string(entry.description);
    ws.cell(index + 2, 2)
      .number(entry.price);
  });

  ws.cell(entries.length + 2, 2)
    .formula(`SUM(B2:B${entries.length + 1})`)

  return wb;
})();

excelWorkbook.writeToBuffer()
  .then(buffer => {
    process.stdout.write(buffer);
  });
