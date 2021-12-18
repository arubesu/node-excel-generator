const excel = require('excel4node');

const fileName = 'Excel.xlsx';
const workbook = new excel.Workbook();
const worksheet = workbook.addWorksheet('Planilha 1');

const borderProps = {
  border: {
    top: {
      style: 'thin',
      color: '#E05508',
    },
    right: {
      style: 'thin',
      color: '#E05508',
    },
    bottom: {
      style: 'thin',
      color: '#E05508',
    },
    left: {
      style: 'thin',
      color: '#E05508',
    },
    diagonal: {
      style: 'thin',
      color: '#E05508',
    }
  },
}

const evenStyle = workbook.createStyle({
  fill: {
    type: 'pattern',
    patternType: 'lightUp',
    fgColor: '#E05508',
  },
  ...borderProps,
});

const oddStyle = workbook.createStyle({
  ...borderProps,
});

const execute = () => {
  const headers = getHeaders();

  const data = getData();

  setHeaders(headers);

  setData(data);

  workbook.write(fileName);
}

const getHeaders = () => {
  return ['Name', 'Email', 'Password', 'Domain', 'Number', 'Text'];
}

const getData = () => {
  return [
    ['Name1', 'Email1', 'Password1', 'Domain1', 12345678912, '12345678912'],
    ['Name2', 'Email2', 'Password2', 'Domain2', 12345678912, '12345678912'],
    ['Name3', 'Email3', 'Password3', 'Domain3', 12345678912, '12345678912'],
    ['Name4', 'Email4', 'Password4', 'Domain4', 12345678912, '12345678912'],
  ]
}

const setHeaders = (headers) => {
  const style = getHeaderStyle();

  for (let column = 0; column < headers.length; column++) {
    const row = 1;
    const nextColumn = column + 1;

    const element = headers[column];
    worksheet.cell(row, nextColumn).string(element).style(style);
  }
};

const getHeaderStyle = () => {
  return workbook.createStyle({
    font: {
      color: '#FFFFFF',
      size: 12,
      bold: true,
    },
    fill: {
      type: 'pattern',
      patternType: 'solid',
      fgColor: '#E05508',
    },
  });
}

const setData = (data) => {
  const columnLength = data[0].length;

  for (let row = 0; row < data.length; row++) {
    for (let column = 0; column < columnLength; column++) {
      const element = data[row][column];

      const nextRow = row + 2; // skip first row (header)
      const nextColumn = column + 1;

      const style = nextRow % 2 === 0 ? evenStyle : oddStyle;

      worksheet.cell(nextRow, nextColumn).string(String(element)).style(style);
    }
  }
};

const getStyle = () => {
  let style = worksheet.Style();

  return style;
}

module.exports = execute;
