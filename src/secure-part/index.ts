import moment from 'moment';
import ExcelJS from 'exceljs';
import path from 'path';

let date = moment("2022-10-01");

const hangulDay = ['일', '월', '화', '수', '목', '금', '토'];

const parkingManageTeam = {
  members: [
    '김선명',
    '신효경',
    '최원정',
    '김미옥',
  ],
  curIdx: 1,
  teamName: '주차관리팀',
};

const transportTeachTeam = {
  members: [
    '송인심',
    '김미림',
    '김보미',
  ],
  curIdx: 0,
  teamName: '운수지도팀',
};

const trafficPenaltyTeam = {
  members: [
    '이진희',
    '김대환',
  ],
  curIdx: 1,
  teamName: '교통과징팀',
};

const parkingControlTeam = {
  members: [
    '최진아',
    '현지원',
    '김신예',
    '정대교',
  ],
  curIdx: 3,
  teamName: '주차단속팀',
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
    const curDay = date.day() - 1;
    const sheet = workbook.getWorksheet(`${date.format('MM')}월`);
    const row = sheet.getRow(3 + curRowIdx);
    row.getCell(2).value = curRowIdx + 1;
    row.getCell(3).value = date.format('MM월 DD일');
    if (!checkRestDay(date)) {
      row.getCell(5).value = {
        richText: [
          {
            text: chargeTeam[curDay].members[chargeTeam[curDay].curIdx],
            font: {
              color: {
                argb: '00000000',
              },
            },
          },
        ],
      };
      row.getCell(4).value = {
        richText: [
          {
            text: hangulDay[date.day()],
            font: {
              color: {
                argb: '00000000',
              },
            },
          },
        ],
      };
      row.getCell(6).value = chargeTeam[curDay].teamName;
      chargeTeam[curDay].curIdx = (chargeTeam[curDay].curIdx + 1) % chargeTeam[curDay].members.length;
    } else {
      row.getCell(4).value = {
        richText: [
          {
            text: hangulDay[date.day()],
            font: {
              color: {
                argb: '00FF0000',
              },
            },
          },
        ],
      };
      row.getCell(5).value = {
        richText: [
          {
            text: '공휴일',
            font: {
              color: {
                argb: '00FF0000',
              },
            },
          },
        ],
      };
      row.getCell(6).value = '';
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
