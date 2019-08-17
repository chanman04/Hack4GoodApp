import { Component, OnInit, ElementRef } from '@angular/core';
import { FormBuilder, FormGroup, Validators, FormControl } from '@angular/forms';
import * as XLSX from 'xlsx';

type AOA = any[][];

@Component({
  selector: 'app-importer',
  templateUrl: './importer.component.html',
  styleUrls: ['./importer.component.css']
})
export class ImporterComponent implements OnInit {

  constructor(private el: ElementRef, private _formBuilder: FormBuilder) { }
  isMaxSelect = false;
  firstFormGroup: FormGroup;
  secondFormGroup: FormGroup;
  currentPage = 0;
  isEmptyDrop = true;
  isExcelDrop = true;
  isRadioChecked = false;
  refExcelData: Array<any>;
  excelFirstRow = [];
  excelDataEncodeToJson;
  excelTransformNum = [];

  sheetJsExcelName = 'null.xlsx';

  sheetCellRange;
  sheetMaxRow;
  localwSheet;
  localWorkbook;
  
  sheetNameForTab: Array<string> = ['excel tab 1', 'excel tab 2'];
  totalPage = this.sheetNameForTab.length;
  selectDefault;
  sheetBufferRender;

  origExcelData: AOA = [
    ['Data: 2018/10/26'],
    ['Data: 2018/10/26'],
    ['Data: 2018/10/26'],
  ];


onClickRadioExcel(){
  if(this.localWorkbook == undefined){
    throw new Error('The value is undefined');
    return;
  }
}

inputExcelOnClick(evt){
  const target: HTMLInputElement = evt.target;
  if(target.files.length == 0){
    throw new Error('You did not upload a Excel file')
  }
  if(target.files.length > 1){
    throw new Error('You uploaded multiple Excel, please upload one at a time')
  }

  this.sheetJsExcelName = evt.target.files.item(0).name;
  const reader: FileReader = new FileReader();
  this.readExcel(reader);
  reader.readAsArrayBuffer(target.files[0]);
  this.sheetBufferRender = target.files[0];
  this.isEmptyDrop = false;
  this.isExcelDrop = false;

}

dropExcelOnChance(targetInput: Array<File>) {
  this.sheetJsExcelName = targetInput[0].name;
  if(targetInput.length !== 1) {
    throw new Error('Sorry we cannot accept multiple files')
  }

  const reader: FileReader = new FileReader();
  this.readExcel(reader);
  reader.readAsArrayBuffer(targetInput[0]);
  this.sheetBufferRender = targetInput[0];
  this.isEmptyDrop = false;
  this.isExcelDrop = true;

}

//transform function
transform(value){
return (value >= 26 ? this.transform(((value / 26) >> 0)-1):'') + 'ABCDEFGHIJKLMNOPQRSTUVWXYZ' [value % 26 >> 0];
}

readExcel(reader, index = 0){

  this.origExcelData = [];
  reader.onload = (e: any) => {
    const data: string = e.target.result;
    const wBook: XLSX.WorkBook = XLSX.read(data, {type: 'array'});
    this.localWorkbook = wBook;
    const wsname: string = wBook.SheetNames[index];
    this.sheetNameForTab = wBook.SheetNames;
    this.totalPage = this.sheetNameForTab.length;
    this.selectDefault = this.sheetNameForTab[index];
    const wSheet: XLSX.WorkSheet = wBook.Sheets[wsname];
    this.localwSheet = wSheet;
    this.sheetCellRange = XLSX.utils.decode_range(wSheet['!ref']);
    this.sheetMaxRow = this.sheetCellRange.e.r;
    this.origExcelData = <AOA>XLSX.utils.sheet_to_json(wSheet, {
      header: 1,
      range: wSheet['!ref'],
      raw: true,
    });

    this.refExcelData = this.origExcelData.slice(1).map(value =>
    Object.assign([], value));
    this.excelTransformNum = [];
    for(let idx = 0; idx <= this.sheetCellRange.e.c; idx++){
      this.excelTransformNum[idx] = this.transform(idx);
    }
    this.refExcelData.map(x => x.unshift('#'));
    this.excelTransformNum.unshift('order');
    this.excelDataEncodeToJson = this.refExcelData.slice(0).map(item =>
    item.reduce((obj, val, i) =>{
      obj[this.excelTransformNum[i]] = val;
      return obj;
    })
      )

  }

}
//function for loading the tab
loadSheetOnTabClick(index: number){
  this.currentPage = index;
  if(this.localWorkbook == undefined) {
    throw new Error('Sorry load of the Excel sheet failed')
  }
  const reader: FileReader = new FileReader();
  this.readExcel(reader, index);
  reader.readAsArrayBuffer(this.sheetBufferRender);

}


  ngOnInit() {
    this.firstFormGroup = this._formBuilder.group({
      firstCtrl: ['', Validators.required],
    });
    this.secondFormGroup = this._formBuilder.group({
      secondCtrl: ['', Validators.required],
    });
  }
}
