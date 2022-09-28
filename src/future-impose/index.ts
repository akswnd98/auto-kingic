import express from 'express';
import fsPromises from 'fs/promises';
import fs from 'fs';
import path from 'path';
import Jimp from 'jimp';
import ExcelJS from 'exceljs';

import { Document, Packer, Paragraph, TextRun, ImageRun } from 'docx';

const inputDir = 'datas';

const cropImage = (width: number, height: number) => {
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

const main = async () => {
  const workbox = new ExcelJS.Workbook();
  await workbox.xlsx.readFile(path.resolve(__dirname, inputDir, 'input.xlsx'));
  const sheet = workbox.getWorksheet('Sheet1');
  await generateDocx(
    (await sortFiles(await fsPromises.readdir(path.resolve(__dirname, inputDir)), sheet)).filter((v) => v !== undefined) as string[],
  );
};

const sortFiles = async (files: string[], sheet: ExcelJS.Worksheet) => {
  const rows = sheet.getRows(1, sheet.rowCount)!;
  const ret = rows.map((row) => {
    const [date, vehicleNumber] = [1, 4].map((v) => row.getCell(v).text);
    const filteredFiles = files.filter((v) => new RegExp(`${date.substring(0, 8)} ${vehicleNumber}.*((\.jpg)|(\.png))`, 'i').test(v));
    if (filteredFiles.length <= 0) {
      console.log(`filter match <= 0: ${vehicleNumber}`);
      return undefined;
    }
    if (filteredFiles.length > 1) {
      console.log(`filter match > 1: ${vehicleNumber}`);
    }
    return filteredFiles[0];
  });
  return ret;
};

const generateDocx = async (files: string[]) => {
  const sections = await Promise.all(files.filter((value) => /(\.jpg)$/i.test(value)).map(async (file) => {
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
            new TextRun({ text: `• 차량번호: ${file.split(' ')[1]}`, break: 2, size: 28, bold: true }),
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

main();
