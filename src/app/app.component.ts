import { Component } from '@angular/core';

import { Workbook } from 'exceljs';
import * as fs from 'file-saver';
import { ExcelService } from './excel.service';

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.css']
})
export class AppComponent {
  title = 'Angular-csv';

  constructor(private excelService: ExcelService){
  }

  generateExcel() {
    this.excelService.generateExcel();
    // console.log(Object.keys(this.data[0]));
  }
}
