import * as xlsx from "ts-xlsx";

const yargs = require("yargs");
const fs = require("fs");

interface MergedCell {
  start: {
    col: number;
    row: number;
  };
  end: {
    col: number;
    row: number;
  };
  sheet: number;
}

function parseMergedCells(file: Uint8Array): MergedCell[] {
  const workbook = xlsx.read(file, { type: "array" });
  const mergedCells: MergedCell[] = [];

  for (let i = 0; i < workbook.SheetNames.length; i++) {
    const worksheet = workbook.Sheets[workbook.SheetNames[i]];
    if (worksheet != null) {
      const merges = worksheet["!merges"];

      for (const mergedCell of merges) {
        mergedCells.push({
          start: {
            col: mergedCell.s.c,
            row: mergedCell.s.r,
          },
          end: {
            col: mergedCell.e.c,
            row: mergedCell.e.r,
          },
	  sheet: i
        });
      }
    }
  }

  return mergedCells;
}

function printMergedCells(mergedCells: MergedCell[]) {
  let count = 0;
  for (const mergedCell of mergedCells) {
    const startCol = mergedCell.start.col + 1;
    const startRow = mergedCell.start.row + 1;
    const endCol = mergedCell.end.col + 1;
    const endRow = mergedCell.end.row + 1;
    const colSpan = endCol - startCol + 1;
    const rowSpan = endRow - startRow + 1;
    const size = colSpan * rowSpan;
    const sheet = mergedCell.sheet + 1;
    count++;

    console.log(
      `Merged cell ${count}: \tStarts at (Col: ${startCol}, Row: ${startRow}) and ends at (Col: ${endCol}, Row: ${endRow})\n\t\tSize: ${size} cells (${
        endCol - startCol + 1
      } columns x ${endRow - startRow + 1} rows)\n\t\tSheet: ${sheet}\n`
    );
  }
}

const argv = yargs.argv;
const filePath = argv._[0];

if (!filePath) {
  console.error("Please specify the file path as an argument");
  process.exit(1);
}

const file = fs.readFileSync(filePath);
const mergedCells = parseMergedCells(new Uint8Array(file));
printMergedCells(mergedCells);
