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
 

  // empPerformance = [
  //   { ID: 10011, NAME: "A", DEPARTMENT: "Sales", MONTH: "Jan", YEAR: 2020, SALES: 132412, CHANGE: 12, LEADS: 35 },
  //   { ID: 10012, NAME: "A", DEPARTMENT: "Sales", MONTH: "Feb", YEAR: 2020, SALES: 232324, CHANGE: 2, LEADS: 443 },
  //   { ID: 10013, NAME: "A", DEPARTMENT: "Sales", MONTH: "Mar", YEAR: 2020, SALES: 542234, CHANGE: 45, LEADS: 345 },
  //   { ID: 10014, NAME: "A", DEPARTMENT: "Sales", MONTH: "Apr", YEAR: 2020, SALES: 223335, CHANGE: 32, LEADS: 234 },
  //   { ID: 10015, NAME: "A", DEPARTMENT: "Sales", MONTH: "May", YEAR: 2020, SALES: 455535, CHANGE: 21, LEADS: 12 },
  // ];

  constructor(public ete: ExportExcelService, private json2htmlService: Json2htmlService) { }

  json2HTML() {
    this.json2htmlService.json2Html(this.empPerformance.body.table_2, this.mytable);
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
        if (this.dataForExcel?.[prop]) {
          this.dataForExcel[prop].push(arr[i])
        }
        else {
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
