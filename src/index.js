const express = require('express');
const bodyParser = require('body-parser');
const axios = require('axios');
const Excel = require('exceljs');

const app = express();

app.use(bodyParser.json());
app.get('/motoristas', async (req, res) => {
  let data = [];
  let letter = ['A','A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O'];
  let count = 0;
  let columns = [];
  let workbook = new Excel.Workbook();
  let worksheet = workbook.addWorksheet('TabName');

  await axios.get('xxxxxxxxxxxxxxxxxx', {
    headers: {'Authorization': 'Bearer eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJ1c3VfSUQiOiI0ZmFhYzlhOTczMmQ0IiwiaWF0IjoxNTk1MDA4NTg5LCJleHAiOjE1OTU2MTMzODl9.iSIbfwyjnIviG6u8VdzrD1aVMI-_P1Ohud6iNh9Jlts'} 
  })
  .then(result => {
    data = result.data.endpoint;

    data.map((item, index) => {
      Object.keys(item).map((objectItem, objectIndex) => {
        if(index == 0) {
          count += 1;
          columns.push({ header: objectItem, key: objectItem });
        }
      })
    });
    worksheet.columns = columns;
  })
  .catch(err => console.log(err));

  worksheet.columns.forEach(column => {
    column.width = column.header.length < 12 ? 12 : column.header.length
  })
  worksheet.getRow(1).font = {bold: true }

  data.forEach((e, i) => {
    worksheet.addRow(e);
  })

  worksheet.eachRow({ includeEmpty: false }, function (row, rowNumber) {
    for(i=0;i<=count;i++) {
      worksheet.getCell(`${letter[i]}1`).border = {
        top: {style: 'thin'},
        left: {style: 'thin'},
        bottom: {style: 'thin'},
        right: {style: 'thin'}
      }
    }
  });

  workbook.xlsx.writeFile('FileName.xlsx');

  res.json({ ok: true })
});

app.listen(8888, () => console.log('🔥 Server running'));