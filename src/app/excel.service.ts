import { Injectable } from '@angular/core';
import { Workbook } from 'exceljs';
import * as fs from 'file-saver';
// import { DatePipe } from '@angular/common';

@Injectable({
  providedIn: 'root'
})
export class ExcelService {

  constructor() { }
  datarow;
  data = [
    {
      name: 'Test 1',
      age: 13,
      average: 8.2,
      approved: true,
      description: "using 'Content here, content here' "
    },
    {
      name: 'Test 2',
      age: 11,
      average: 8.2,
      approved: true,
      description: "using 'Content here, content here' "
    },
    {
      name: 'Test 3',
      age: 10,
      average: 8.2,
      approved: true,
      description: "using 'Content here, content here' "
    }
  ];

  generateExcel()
  {
    const title = 'Report';
    const header = Object.keys(this.data[0]);  // extracts the key from json key-value pair


    let workbook = new Workbook();
    let worksheet = workbook.addWorksheet('Sheet 1');  //  add the sheet to workbook


    let titleRow = worksheet.addRow([title]);   // add title to sheet

    worksheet.addRow([]);   // adds empty row

    let headerRow = worksheet.addRow(header);   // adds the column names
    headerRow.font = {bold : true};  // makes the column heads bold


	  // writing the text to buffer
    for (let i=0; i< this.data.length; i++)  {
      this.datarow = worksheet.addRow(Object.values(this.data[i]));
      this.datarow.font = {bold: false}; //  for making the text not bold while writing
    }


    worksheet.getColumn(5).width = 35;   // for increasing the size of a column. 5 is column number
    worksheet.addRow([]);


    // For saving the xlsx file
    workbook.xlsx.writeBuffer().then((data) => {
      let blob = new Blob([data], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      fs.saveAs(blob, 'CarData11.xlsx');
    });

  }
}
