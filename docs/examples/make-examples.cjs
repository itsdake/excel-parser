/* Собирает примеры .xlsx для проверки универсального парсера: по файлу на раскладку из README («Входные данные»).
   Запуск: node docs/examples/make-examples.cjs — файлы появятся рядом со скриптом.
   Числа — числа, текст — строки внутри ячейки; оформления нет (оформление из файла читалка проверена отдельно). */
const fs = require('node:fs');
const path = require('node:path');
const { zip } = require('../../backend/test/helpers/make-xlsx.js');

const XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
const MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

const letters = (i) => (i >= 26 ? letters(Math.floor(i / 26) - 1) : '') + String.fromCharCode(65 + (i % 26));
const escape = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');

function sheetXml(rows) {
  const body = rows.map((line, r) => {
    const cells = line.map((v, c) => {
      if (v === null || v === undefined || v === '') return '';
      const ref = letters(c) + (r + 1);
      return typeof v === 'number'
        ? `<c r="${ref}"><v>${v}</v></c>`
        : `<c r="${ref}" t="inlineStr"><is><t>${escape(v)}</t></is></c>`;
    }).join('');
    return `<row r="${r + 1}">${cells}</row>`;
  }).join('');
  return XML + `<worksheet xmlns="${MAIN}"><sheetData>${body}</sheetData></worksheet>`;
}

function workbook(sheets) {
  const tags = sheets.map((s, i) => `<sheet name="${escape(s.name)}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`).join('');
  const rels = sheets.map((s, i) => `<Relationship Id="rId${i + 1}" Type="${REL}/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`).join('');
  const entries = [
    { name: '[Content_Types].xml', data: XML + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>' },
    { name: 'xl/workbook.xml', data: XML + `<workbook xmlns="${MAIN}" xmlns:r="${REL}"><sheets>${tags}</sheets></workbook>` },
    { name: 'xl/_rels/workbook.xml.rels', data: XML + `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels}</Relationships>` }
  ];
  sheets.forEach((s, i) => entries.push({ name: `xl/worksheets/sheet${i + 1}.xml`, data: sheetXml(s.rows) }));
  return zip(entries);
}

const FILES = {
  'Б-имена-рядов.xlsx': [
    { name: 'План и факт', rows: [['', 'Янв', 'Фев', 'Мар', 'Апр'], ['План', 10, 12, 9, 14], ['Факт', 8, 13, 11, 12]] }
  ],
  'В-вертикальная-таблица.xlsx': [
    { name: 'Сотрудники', rows: [['Сотрудник', 'Заказы', 'Выполнено'], ['Матвей', 10, 8], ['Иван', 5, 5], ['Данил', 34, 30], ['Светлана', 17, 15], ['Алина', 42, 40]] }
  ],
  'Г-без-заголовков.xlsx': [
    { name: 'Заказы', rows: [['Матвей', 10], ['Иван', 5], ['Данил', 34], ['Светлана', 17], ['Алина', 42]] }
  ],
  'Д-только-числа.xlsx': [
    { name: 'Строка', rows: [[10, 5, 34, 17, 42]] },
    { name: 'Столбец', rows: [[3], [8], [1], [9]] }
  ],
  'Название-таблицы.xlsx': [
    { name: 'Март', rows: [['Отчёт за март', '', ''], ['Матвей', 'Иван', 'Данил'], [10, 5, 34]] }
  ],
  'Разные-листы.xlsx': [
    { name: 'Продажи', rows: [['Регион', 'Квартал 1', 'Квартал 2', 'Комментарий'], ['Север', 120, 135, 'рост'], ['Юг', 90, 80, 'спад'], ['Запад', 60, 75, 'рост']] },
    { name: 'Числа как текст', rows: [['Янв', 'Фев', 'Мар'], ['10', '3,5', '1 234']] },
    { name: 'Пропуски', rows: [['а', 'б', 'в', 'г'], [4, 'нет', 6, 8]] },
    { name: 'Только текст', rows: [['Имя', 'Должность'], ['Матвей', 'Аналитик'], ['Иван', 'Менеджер']] }
  ]
};

for (const [file, sheets] of Object.entries(FILES)) {
  fs.writeFileSync(path.join(__dirname, file), workbook(sheets));
  console.log('создан', file);
}
