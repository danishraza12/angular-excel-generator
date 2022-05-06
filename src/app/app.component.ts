import { Component, ElementRef, ViewChild } from '@angular/core';
import { ExportExcelService } from './services/export-excel.service';
import { Json2htmlService } from './services/json2html.service';
import { Emp_Performance } from './empPerformance';

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.css']
})
export class AppComponent {
  title = 'ExcelTest';

  @ViewChild('mytable') mytable: ElementRef | undefined;
  dataForExcel: any;

  empPerformance: any = Emp_Performance;

  constructor(public ete: ExportExcelService, private json2htmlService: Json2htmlService) { }

  json2HTML() {
    for (let prop in this.empPerformance.body) {
      this.json2htmlService.json2Html(this.empPerformance.body[prop], this.mytable);
    }
  }
  
  exportToExcel() {
    let headers = []
    this.dataForExcel = []

    for (let prop in this.empPerformance?.body) {    
      let obj = this.empPerformance.body[prop];
      let arr:any = []

      obj.forEach((row: any) => {
        arr.push(Object.values(row))
      })
      
      for (let i=0;i<arr.length;i++) {
        if (this.dataForExcel?.[prop]) { // if prop exists so just push
          this.dataForExcel[prop].push(arr[i])
        }
        else {
          // Create dynamic property if prop does not exist
          this.dataForExcel = { ...this.dataForExcel, [prop]: [] };
          this.dataForExcel[prop].push(arr[i]);
        }
      }

      headers.push(Object.keys(this.empPerformance?.body?.[prop]?.[0]));
    }

    let reportData = {
      title: 'Employee Sales Report - Jan 2020',
      data: this.dataForExcel,
      headers
    }

    this.ete.exportExcel(reportData);
  }

}
