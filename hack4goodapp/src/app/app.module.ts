import { BrowserModule } from '@angular/platform-browser';
import { NgModule } from '@angular/core';
import { DraganddropDirective } from './draganddrop.directive';

import { AppComponent } from './app.component';
import { ExcelService } from './excel.service';
import { ImporterComponent } from './importer/importer.component';
import { MaterialModule} from '@app/material.module';
import { FormsModule } from '@angular/forms';

@NgModule({
  declarations: [
    AppComponent,
    DraganddropDirective,
    ImporterComponent
  ],
  imports: [
    BrowserModule, MaterialModule, FormsModule
  ],
  providers: [ExcelService],
  bootstrap: [AppComponent]
})
export class AppModule { }
