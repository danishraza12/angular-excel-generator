import { Component } from '@angular/core';
import { ExportExcelService } from './services/export-excel.service';

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.css']
})
export class AppComponent {
  title = 'ExcelTest';

  dataForExcel: any = [];

  empPerformance = [
    { ID: 10011, NAME: "ABC", DEPARTMENT: "Sales", MONTH: "Jan", YEAR: 2020, SALES: 132412, CHANGE: 12, LEADS: 35, NEWCOLUMN: "Test Column 1" },
    { ID: 10012, NAME: "ABC", DEPARTMENT: "Sales", MONTH: "Feb", YEAR: 2020, SALES: 232324, CHANGE: 2, LEADS: 443, NEWCOLUMN: "Test Column 2" },
    { ID: 10013, NAME: "ABC", DEPARTMENT: "Sales", MONTH: "Mar", YEAR: 2020, SALES: 542234, CHANGE: 45, LEADS: 345, NEWCOLUMN: "Test Column 3" },
    { ID: 10014, NAME: "ABC", DEPARTMENT: "Sales", MONTH: "Apr", YEAR: 2020, SALES: 223335, CHANGE: 32, LEADS: 234, NEWCOLUMN: "Test Column 4" },
    { ID: 10015, NAME: "ABC", DEPARTMENT: "Sales", MONTH: "May", YEAR: 2020, SALES: 455535, CHANGE: 21, LEADS: 12, NEWCOLUMN: "Test Column 5" },
  ];

  constructor(public ete: ExportExcelService) { }

  exportToExcel() {

    this.empPerformance.forEach((row: any) => {
      this.dataForExcel.push(Object.values(row))
    })

    let reportData = {
      title: 'Employee Sales Report - Jan 2020',
      data: this.dataForExcel,
      headers: Object.keys(this.empPerformance[0])
    }

    this.ete.exportExcel(reportData);
  }
}
