import { Component } from '@angular/core';
import {ExcelService} from './excel.service';
@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.css']
})
export class AppComponent {
  title = 'hack4goodapp';

  data: any = [{
    superhero: 'Iron Man',
    villain: 'Mandarin'
  },
  {
    superhero: 'Incredible Hulk',
    villain: 'The Leader'
  
  },
  {
    superhero: 'The Black Panther',
    villain: 'The KKK'
}];

constructor (private excelService:ExcelService) {
  
  }
exportExcel():void {
this.excelService.exportExcel(this.data, 'sample');
  }
}
