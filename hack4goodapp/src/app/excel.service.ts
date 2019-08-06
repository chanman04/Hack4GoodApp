import { Injectable } from '@angular/core';
import * as XLSX from 'xlsx';
import * as FileSaver from 'file-saver';


const EXCEL_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheethtml.sheet;charset=UTF-8';
const EXCEL_EXTENSION = '.xlsx'


@Injectable({
  providedIn: 'root'
})
export class ExcelService {

  constructor() { }

public exportExcel(json: any[], excelFileName: string):void {

const worksheet: XLSX.WorkSheet = XLSX.utils.json_to_sheet(json);
console.log('worksheet', worksheet);
const workbook: XLSX.WorkBook = {Sheets: {'data': worksheet}, SheetNames: ['data']};
const excelBuffer: any = XLSX.write(workbook, { bookType: 'xlsx', type: 'array'});
this.saveAsExcel({ buffer: excelBuffer, fileName: excelFileName });
  }

private saveAsExcel({ buffer, fileName }: { buffer: any; fileName: string; }): void{
  const data: Blob = new Blob([buffer], {
    type: EXCEL_TYPE
  });
 FileSaver.saveAs(data, fileName + '_export_' + new Date().getTime() + EXCEL_EXTENSION);
  }


}

