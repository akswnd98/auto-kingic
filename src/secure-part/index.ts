import moment from 'moment';
import ExcelJS from 'exceljs';
import path from 'path';

var date = moment("2022-10-01");

const hangulDay = ['일', '월', '화', '수', '목', '금', '토'];

const parkingManageTeam = {
  members: [
    '김선명',
    '신효경',
    '최원정',
    '김미옥',
  ],
  curIdx: 1,
};

const transportTeachTeam = {
  members: [
    '송인심',
    '김미림',
    '김보미',
  ],
  curIdx: 0,
};

const trafficPenaltyTeam = {
  members: [
    '이진희',
    '김대희',
  ],
  curIdx: 1,
};

const parkingControlTeam = {
  members: [
    '최진아',
    '현지원',
    '김신예',
    '정대교',
  ],
  curIdx: 3,
};

const chargeTeam = [
  parkingControlTeam,
  trafficPenaltyTeam,
  parkingManageTeam,
  transportTeachTeam,
  parkingManageTeam,
];

const additionalHoliday = [
  '2022-10-03',
  '2022-10-10',
];

const checkRestDay = (day: moment.Moment) => {
  if (day.day() === 0 || day.day() === 6) {
    return true;
  }
  for (let i = 0; i < additionalHoliday.length; i++) {
    if (day.format('YYYY-MM-DD') === additionalHoliday[i]) {
      return true;
    }
  }
}

const main = async () => {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(path.resolve(__dirname, '보안근무 명단(2022년 9월) .xlsx'));
  let beforeMonth = date.month() + 1;
  let curRowIdx = 0;
  while (date.format('YYYY-MM-DD') !== '2023-01-01') {
    if (beforeMonth !== date.month() + 1) {
      beforeMonth = date.month() + 1
      curRowIdx = 0;
    }
    const sheet = workbook.getWorksheet(`${date.format('MM')}월`);
    const row = sheet.getRow(3 + curRowIdx);
    row.getCell(3).value = date.format('MM월 DD일');
    if (!checkRestDay(date)) {
      const curDay = date.day() - 1;
      row.getCell(5).value = {
        richText: [{
          text: chargeTeam[curDay].members[chargeTeam[curDay].curIdx];
        }],
        
      row.getCell(5)
      row.getCell(4).value = hangulDay[date.day()];
      chargeTeam[curDay].curIdx = chargeTeam[curDay].curIdx % chargeTeam[curDay].members.length;
    } else {
    }
    date.add(1, 'day');
    curRowIdx++;
  }
  workbook.xlsx.writeFile(path.resolve(__dirname, 'output.xlsx'));
};

main();

const test = async () => {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(path.resolve(__dirname, '보안근무 명단(2022년 9월) .xlsx'));
  const sheet = workbook.getWorksheet(`${date.format('MM')}월`);
  console.log(sheet.getCell(3, 2).text);
};

// test();
