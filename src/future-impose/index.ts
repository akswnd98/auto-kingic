import fsPromises from 'fs/promises';
import fs from 'fs';
import path from 'path';
import Jimp from 'jimp';
import ExcelJS from 'exceljs';

import { Document, Packer, Paragraph, TextRun, ImageRun } from 'docx';

const inputDir = 'datas';

const main = async () => {
  const workbox = new ExcelJS.Workbook();
  await workbox.xlsx.readFile(path.resolve(__dirname, inputDir, 'input.xlsx'));
  const sheet = workbox.getWorksheet('Sheet1');
  await generateDocx(sheet);
};

async function generateDocx (sheet: ExcelJS.Worksheet) {
  const pair = await getValidVehicleNumberFilePairs(sheet);
  const sections = await Promise.all(pair.map(async ([vehicleNumber, file]) => {
    const img = await Jimp.read(path.resolve(__dirname, inputDir, file));
    return {
      children: [
        new Paragraph({
          children: [
            new ImageRun({
              data: await fsPromises.readFile(path.resolve(__dirname, inputDir, file)),          
              transformation: {
                ...cropImage(img.getWidth(), img.getHeight()),
              },              
            }),
            new TextRun({ text: '• 위반사유: 장애인전용주차구역 불법주정차', break: 3, size: 28, bold: true }),
            new TextRun({ text: `• 차량번호: ${vehicleNumber}`, break: 2, size: 28, bold: true }),
          ],
        }),
      ],
    };
  }));

  const doc = new Document({
    sections,
  });

  Packer.toBuffer(doc).then((buffer) => {
    fs.writeFileSync(path.resolve(__dirname, 'datas', 'output.docx'), buffer);
  });
};

function cropImage (width: number, height: number) {
  let retWidth = width;
  let retHeight = height;
  if (width > 600) {
    retWidth = 600;
    retHeight = Math.round(height * 600 / width);
  }
  if (retHeight > 800) {
    retHeight = 800;
  }
  return {
    width: retWidth,
    height: retHeight,
  };
};

async function getValidVehicleNumberFilePairs (sheet: ExcelJS.Worksheet) {
  const vehicleNumbers = getSortedVehicleNumbers(sheet);
  const files = await fsPromises.readdir(path.resolve(__dirname, inputDir));
  const ret: [string, string][] = [];
  vehicleNumbers.forEach((v) => {
    try {
      const filteredFiles = findFilesByVehicleNumber(v, files);
      logFileVehicleNumberMismatch(v, filteredFiles);
      throwFileVehicleUndermatch(v, filteredFiles);
      ret.push([v, filteredFiles[0]]);
    } catch (e) {
    }
  });
  return ret;
};

function getSortedVehicleNumbers (sheet: ExcelJS.Worksheet) {
  const rows = sheet.getRows(1, sheet.rowCount)!;
  return rows.map((row) => (
     row.getCell(4).text
  ));
};

function findFilesByVehicleNumber (vehicleNumber: string, files: string[]) {
  return files.filter((v) => new RegExp(`.*${vehicleNumber}.*`, 'i').test(v));
};

function logFileVehicleNumberMismatch (vehicleNumber: string, filteredFiles: string[]) {
  [
    logFileVehicleNumberUndermatch,
    logFileVehicleNumberOvermatch,
  ].forEach((v) => {
    v(vehicleNumber, filteredFiles);
  });
}

function throwFileVehicleUndermatch (vehicleNumber: string, filteredFiles: string[]) {
  if (filteredFiles.length <= 0) {
    throw 'undermatch';
  }
}

function logFileVehicleNumberUndermatch (vehicleNumber: string, filteredFiles: string[]) {
  if (filteredFiles.length <= 0) {
    console.log(`filter match <= 0: ${vehicleNumber}`);
  }
};

function logFileVehicleNumberOvermatch (vehicleNumber: string, filteredFiles: string[]) {
  if (filteredFiles.length > 1) {
    console.log(`filter match > 1: ${vehicleNumber}`);
  }
};

main();
