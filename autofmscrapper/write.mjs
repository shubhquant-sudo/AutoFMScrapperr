import fs from 'node:fs/promises';
import { createRequire } from 'node:module';
import { pathToFileURL } from 'node:url';
const require = createRequire(import.meta.url);
const {
  Workbook,
  SpreadsheetFile
} = await import(pathToFileURL(require.resolve('@oai/artifact-tool', {
  paths: [process.env.FM_ARTIFACT_MODULES]
})).href);
const [specPath, output] = process.argv.slice(2);
const modelSpec = JSON.parse(await fs.readFile(specPath, 'utf8'));
const workbook = Workbook.create();
const colors = {
  teal: '#264E58',
  slate: '#7896AF',
  ink: '#172D42',
  muted: '#657680',
  yellow: '#FFF5D2',
  blue: '#3271D2',
  green: '#008000',
  pale: '#E6EEF1'
};
function address(row, column) {
  let columnName = '';
  while (column) {
    columnName = String.fromCharCode(65 + (column - 1) % 26) + columnName;
    column = Math.floor((column - 1) / 26);
  }
  return columnName + row;
}
for (const sheetSpec of modelSpec.sheets) workbook.worksheets.add(sheetSpec.name);
for (const sheetSpec of modelSpec.sheets) {
  const worksheet = workbook.worksheets.getItem(sheetSpec.name);
  worksheet.showGridLines = false;
  const lastColumn = Math.max(sheetSpec.name === 'History' ? 4 : sheetSpec.name ===
    'Raw statements' ? 8 : 12, ...sheetSpec.cells.map(cellSpec => cellSpec.col));
  const end = address(sheetSpec.rows, lastColumn);
  const range = worksheet.getRange(`A1:${end}`);
  range.format.font = {
    name: 'Arial',
    size: 10,
    color: colors.ink
  };
  range.format.rowHeight = 22;
  range.format.verticalAlignment = 'center';
  worksheet.getRange('A:A').format.columnWidth = 1.15;
  worksheet.getRange('B:B').format.columnWidth = 43;
  worksheet.getRange('C:C').format.columnWidth = 3;
  worksheet.getRange(`D:${address(1,lastColumn).replace(/1$/,'')}`).format.columnWidth =
    15;
  if (sheetSpec.name === 'Assumptions') {
    worksheet.getRange('B:B').format.columnWidth = 38;
    worksheet.getRange('D:D').format.columnWidth = 13;
    worksheet.getRange('F:J').format.columnWidth = 13;
    worksheet.getRange('E:E').format.columnWidth = 3;
    worksheet.getRange('K:K').format.columnWidth = 3;
    worksheet.getRange('L:L').format.columnWidth = 52;
  }
  if (sheetSpec.name === 'Statements') {
    const gutter = modelSpec.forecast_columns[0] - 1;
    worksheet.getRange(
        `${address(1,gutter).replace(/1$/,'')}:${address(1,gutter).replace(/1$/,'')}`)
      .format.columnWidth = 3;
    worksheet.getRange('M:M').format.columnWidth = 3;
    worksheet.getRange('N:N').format.columnWidth = 70;
  }
  if (sheetSpec.name === 'DCF') {
    worksheet.getRange('I:I').format.columnWidth = 3;
    worksheet.getRange('J:J').format.columnWidth = 72;
  }
  if (sheetSpec.name === 'History') {
    worksheet.getRange('C:C').format.columnWidth = 23;
    worksheet.tabColor = '#D5C7AE';
  }
  if (sheetSpec.name === 'Raw statements') {
    worksheet.getRange('B:B').format.columnWidth = 52;
    worksheet.getRange('C:C').format.columnWidth = 3;
    worksheet.getRange(`D:${address(1,lastColumn).replace(/1$/,'')}`).format.columnWidth =
      21;
    worksheet.tabColor = '#D5C7AE';
  }
  const cells = new Map(sheetSpec.cells.map(cellSpec => [
    `${cellSpec.row}:${cellSpec.col}`, cellSpec
  ]));
  const values = Array.from({
    length: sheetSpec.rows
  }, () => Array(lastColumn).fill(null));
  const formulas = [];
  for (const cellSpec of cells.values()) {
    if (typeof cellSpec.value === 'string' && cellSpec.value.startsWith('=') && cellSpec
      .kind !== 'raw_source') formulas.push(cellSpec);
    else values[cellSpec.row - 1][cellSpec.col - 1] = cellSpec.kind === 'raw_source' &&
      typeof cellSpec.value === 'string' && cellSpec.value.trimStart().startsWith('=') ?
      "'" + cellSpec.value : cellSpec.value;
  }
  worksheet.getRange(`A1:${end}`).values = values;
  const formulasByRow = new Map();
  for (const cellSpec of formulas) {
    if (!formulasByRow.has(cellSpec.row)) formulasByRow.set(cellSpec.row, []);
    formulasByRow.get(cellSpec.row).push(cellSpec);
  }
  for (const [rowNumber, items] of formulasByRow) {
    items.sort((a, b) => a.col - b.col);
    let contiguousCells = [];
    const writeFormulaGroup = () => {
      if (contiguousCells.length) worksheet.getRange(
        `${address(rowNumber,contiguousCells[0].col)}:${address(rowNumber,contiguousCells.at(-1).col)}`
        ).formulas = [contiguousCells.map(cellSpec => cellSpec.value)];
      contiguousCells = [];
    };
    for (const cellSpec of items) {
      if (contiguousCells.length && cellSpec.col !== contiguousCells.at(-1).col + 1)
        writeFormulaGroup();
      contiguousCells.push(cellSpec);
    }
    writeFormulaGroup();
  }
  for (const cellSpec of cells.values()) {
    const cell = worksheet.getRange(address(cellSpec.row, cellSpec.col));
    cell.setNumberFormat(cellSpec.format || 'General');
    const formula = formulas.includes(cellSpec);
    cell.format.horizontalAlignment = typeof cellSpec.value === 'number' || formula ?
      'right' : 'left';
    if (formula) cell.format.font.color = [...cellSpec.value.matchAll(/'([^']+)'!/g)]
      .some(m => m[1] !== sheetSpec.name) ? colors.green : '#000000';
    if (cellSpec.kind === 'label' || cellSpec.kind === 'source_label') {
      cell.format.horizontalAlignment = 'left';
      if (String(cellSpec.value).length > 42) {
        cell.format.wrapText = true;
        worksheet.getRange(`${cellSpec.row}:${cellSpec.row}`).format.rowHeight = Math.max(
          30, Math.ceil(String(cellSpec.value).length / (sheetSpec.name ===
            'Raw statements' ? 52 : 43)) * 14);
      }
    }
    if (cellSpec.kind === 'input' || cellSpec.kind === 'missing_source') {
      cell.format.fill = colors.yellow;
      cell.format.font.color = colors.blue;
    }
    if (cellSpec.kind === 'source') cell.format.font.color = colors.green;
    if (cellSpec.kind === 'raw_source' && typeof cellSpec.value === 'number') cell
      .setNumberFormat(Number.isInteger(cellSpec.value) ? '#,##0;(#,##0);"-"' :
        '#,##0.00####;(#,##0.00####);"-"');
    if (cellSpec.kind === 'header') {
      cell.format.fill = colors.pale;
      cell.format.font = {
        name: 'Arial',
        size: 10,
        bold: true,
        color: colors.teal
      };
      cell.format.horizontalAlignment = 'center';
      cell.format.wrapText = true;
      worksheet.getRange(`${cellSpec.row}:${cellSpec.row}`).format.rowHeight = 32;
    }
    if (cellSpec.kind === 'explanation') {
      cell.format.font.color = colors.muted;
      cell.format.wrapText = true;
      cell.format.verticalAlignment = 'center';
      worksheet.getRange(`${cellSpec.row}:${cellSpec.row}`).format.rowHeight = Math.max(
        32, Math.ceil(String(cellSpec.value).length / (sheetSpec.name ===
          'Assumptions' ? 52 : 70)) * 14);
    }
    if (cellSpec.kind === 'status') {
      cell.format.font = {
        name: 'Arial',
        size: 10,
        bold: true,
        color: colors.teal
      };
      cell.format.horizontalAlignment = 'left';
      cell.format.wrapText = true;
      worksheet.getRange(`${cellSpec.row}:${cellSpec.row}`).format.rowHeight = 34;
    }
  }
  const titleEnd = sheetSpec.name === 'DCF' ? 'H' : sheetSpec.name === 'Assumptions' ?
    'J' : sheetSpec.name === 'History' ? address(1, Math.max(7, lastColumn)).replace(/1$/,
      '') : sheetSpec.name === 'Raw statements' ? 'H' : address(1, modelSpec
      .forecast_columns.at(-1)).replace(/1$/, '');
  worksheet.getRange('B2').values = [
    [`${modelSpec.company} — ${sheetSpec.name}`]
  ];
  worksheet.getRange(`B2:${titleEnd}2`).merge();
  worksheet.getRange(`B2:${titleEnd}2`).format.font = {
    name: 'Arial',
    size: 14,
    bold: true,
    color: colors.teal
  };
  worksheet.getRange('2:2').format.rowHeight = 30;
  worksheet.getRange('B3').values = [
    [
      `${sheetSpec.name==='Raw statements'?'Original source units':modelSpec.unit}. ${sheetSpec.name==='DCF'?'Shares in millions. Per-share values in '+modelSpec.currency+'. ':''}Latest source period: ${modelSpec.periods.at(-1)}.`]
  ];
  worksheet.getRange(`B3:${titleEnd}3`).merge();
  worksheet.getRange(`B3:${titleEnd}3`).format.font = {
    name: 'Arial',
    size: 9,
    italic: true,
    color: colors.muted
  };
  worksheet.getRange(`B3:${titleEnd}3`).format.borders = {
    bottom: {
      style: 'thin',
      color: colors.teal
    }
  };
  for (const [rowNumber, text] of sheetSpec.sections) {
    const block = worksheet.getRange(`B${rowNumber}:${titleEnd}${rowNumber}`);
    block.merge();
    block.values = [
      [text]
    ];
    block.format.fill = colors.teal;
    block.format.font = {
      name: 'Arial',
      size: 10,
      bold: true,
      color: '#FFFFFF'
    };
    block.format.rowHeight = 25;
  }
  for (const rowNumber of sheetSpec.totals) {
    const block = worksheet.getRange(`B${rowNumber}:${titleEnd}${rowNumber}`);
    block.format.font.bold = true;
    block.format.borders = {
      top: {
        style: 'thin',
        color: colors.slate
      }
    };
  }
  for (const [rowNumber, text] of sheetSpec.notes) {
    const block = worksheet.getRange(`B${rowNumber}:${titleEnd}${rowNumber+1}`);
    block.merge();
    worksheet.getRange(`B${rowNumber}`).values = [
      [text]
    ];
    block.format.font = {
      name: 'Arial',
      size: 10,
      color: colors.muted
    };
    block.format.wrapText = true;
    block.format.verticalAlignment = 'center';
    block.format.rowHeight = Math.max(22, Math.ceil(text.length / 110) * 9);
  }
  for (const item of Object.values(modelSpec.inputs).filter(x => x.sheet === sheetSpec
      .name)) {
    const rule = item.rule;
    let validation;
    if (rule.min !== undefined && rule.max !== undefined) validation = {
      type: 'decimal',
      operator: 'between',
      formula1: rule.min,
      formula2: rule.max
    };
    else if (rule.min !== undefined) validation = {
      type: 'decimal',
      operator: rule.strict ? 'greaterThan' : 'greaterThanOrEqual',
      formula1: rule.min
    };
    else if (rule.max !== undefined) validation = {
      type: 'decimal',
      operator: 'lessThanOrEqual',
      formula1: rule.max
    };
    if (validation) worksheet.getRange(item.cell).dataValidation = {
      rule: validation
    };
  }
  for (const area of sheetSpec.merges) worksheet.getRange(area).merge();
  if (sheetSpec.name === 'Statements') {
    for (const rowNumber of [107, 108, 109]) {
      const region = worksheet.getRange(
        `D${rowNumber}:${address(rowNumber,modelSpec.forecast_columns.at(-1))}`);
      region.conditionalFormats.add('cellIs', {
        operator: 'greaterThan',
        formula: 0.5,
        format: {
          fill: '#F8D8D5'
        }
      });
      region.conditionalFormats.add('cellIs', {
        operator: 'lessThan',
        formula: -0.5,
        format: {
          fill: '#F8D8D5'
        }
      });
    }
    worksheet.getRange(`D31:${address(31,modelSpec.forecast_columns.at(-1))}`)
      .conditionalFormats.add('cellIs', {
        operator: 'lessThan',
        formula: 0,
        format: {
          fill: '#F8D8D5'
        }
      });
  }
  if (sheetSpec.name === 'Assumptions') worksheet.getRange('F25:J25').conditionalFormats
    .add('cellIs', {
      operator: 'greaterThan',
      formula: 0,
      format: {
        fill: colors.yellow
      }
    });
  worksheet.freezePanes.freezeRows(7);
  worksheet.tabColor = sheetSpec.name === 'Assumptions' ? colors.slate : sheetSpec
    .name === 'Statements' || sheetSpec.name === 'DCF' ? colors.teal : '#D5C7AE';
}
workbook.recalculate();
const out = await SpreadsheetFile.exportXlsx(workbook);
await out.save(output);
console.log(JSON.stringify({
  output,
  sheets: modelSpec.sheets.length
}));
