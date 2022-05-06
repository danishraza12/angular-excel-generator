import { Component } from '@angular/core';
import { ExportExcelService } from './services/export-excel.service';

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.css']
})
export class AppComponent {
  title = 'ExcelTest';

  dataForExcel: any;

  empPerformance: any = {
    body: {
       table_1: [
          {
             "employeeindex": 0,
             "clientindex": "0",
             "clientname": "",
             "buname": "",
             "employeename": "",
             "gradename": "",
             "designationname": "",
             "employmenttypedesc": ""
          },
          {
             "employeeindex": 300000001,
             "clientindex": "300000001",
             "clientname": "HRSG BPO",
             "buname": "HRSG BPO",
             "employeename": "Umair Ahmed Bukhari",
             "gradename": "",
             "designationname": "",
             "employmenttypedesc": ""
          },
          {
             "employeeindex": 300000002,
             "clientindex": "300000001",
             "clientname": "HRSG BPO",
             "buname": "HRSG BPO",
             "employeename": "Umair Ahmed",
             "gradename": "",
             "designationname": "",
             "employmenttypedesc": ""
          },
          {
             "employeeindex": 300000003,
             "clientindex": "300000001",
             "clientname": "HRSG BPO",
             "buname": "HRSG BPO",
             "employeename": "Yawer Sheikh",
             "gradename": "",
             "designationname": "",
             "employmenttypedesc": ""
          },
          {
             "employeeindex": 300000004,
             "clientindex": "0",
             "clientname": "",
             "buname": "",
             "employeename": "Umair Ahmed Bukhari 10",
             "gradename": "",
             "designationname": "",
             "employmenttypedesc": ""
          },
          {
             "employeeindex": 300000005,
             "clientindex": "0",
             "clientname": "",
             "buname": "",
             "employeename": "Nabeel Khan ",
             "gradename": "",
             "designationname": "",
             "employmenttypedesc": ""
          },
          {
             "employeeindex": 300000006,
             "clientindex": "300000001",
             "clientname": "HRSG BPO",
             "buname": "HRSG BPO",
             "employeename": "Nabeel Khan ",
             "gradename": "",
             "designationname": "",
             "employmenttypedesc": ""
          },
          {
             "employeeindex": 300000007,
             "clientindex": "0",
             "clientname": "",
             "buname": "",
             "employeename": "",
             "gradename": "-",
             "designationname": "-",
             "employmenttypedesc": ""
          },
          {
             "employeeindex": 300000008,
             "clientindex": "0",
             "clientname": "",
             "buname": "",
             "employeename": "",
             "gradename": "-",
             "designationname": "-",
             "employmenttypedesc": ""
          },
          {
             "employeeindex": 300000009,
             "clientindex": "0",
             "clientname": "",
             "buname": "",
             "employeename": "",
             "gradename": "-",
             "designationname": "-",
             "employmenttypedesc": ""
          },
          {
             "employeeindex": 300000010,
             "clientindex": "0",
             "clientname": "",
             "buname": "",
             "employeename": "Test JSON 1",
             "gradename": "-",
             "designationname": "-",
             "employmenttypedesc": ""
          },
          {
             "employeeindex": 300000011,
             "clientindex": "0",
             "clientname": "",
             "buname": "",
             "employeename": "Uzair Sabir",
             "gradename": "-",
             "designationname": "-",
             "employmenttypedesc": ""
          },
          {
             "employeeindex": 300000012,
             "clientindex": "0",
             "clientname": "",
             "buname": "",
             "employeename": "",
             "gradename": "-",
             "designationname": "-",
             "employmenttypedesc": ""
          },
          {
             "employeeindex": 300000013,
             "clientindex": "0",
             "clientname": "",
             "buname": "",
             "employeename": "Umair Ahmed Bukhari",
             "gradename": "",
             "designationname": "",
             "employmenttypedesc": ""
          },
          {
             "employeeindex": 300000014,
             "clientindex": "0",
             "clientname": "",
             "buname": "",
             "employeename": "Umair Ahmed Bukhari",
             "gradename": "",
             "designationname": "",
             "employmenttypedesc": ""
          },
          {
             "employeeindex": 300000015,
             "clientindex": "0",
             "clientname": "",
             "buname": "",
             "employeename": "Umair Ahmed Bukhari",
             "gradename": "",
             "designationname": "",
             "employmenttypedesc": ""
          },
          {
             "employeeindex": 300000016,
             "clientindex": "0",
             "clientname": "",
             "buname": "",
             "employeename": "Uzair sabir",
             "gradename": "",
             "designationname": "",
             "employmenttypedesc": ""
          },
          {
             "employeeindex": 300000017,
             "clientindex": "0",
             "clientname": "",
             "buname": "",
             "employeename": "TestUzair",
             "gradename": "-",
             "designationname": "-",
             "employmenttypedesc": ""
          },
          {
             "employeeindex": 300000018,
             "clientindex": "0",
             "clientname": "",
             "buname": "",
             "employeename": "",
             "gradename": "-",
             "designationname": "-",
             "employmenttypedesc": ""
          },
          {
             "employeeindex": 300000019,
             "clientindex": "0",
             "clientname": "",
             "buname": "",
             "employeename": "Umair Ahmed Bukhari",
             "gradename": "",
             "designationname": "",
             "employmenttypedesc": ""
          },
          {
             "employeeindex": 300000020,
             "clientindex": "0",
             "clientname": "",
             "buname": "",
             "employeename": "",
             "gradename": "-",
             "designationname": "-",
             "employmenttypedesc": ""
          },
          {
             "employeeindex": 300000021,
             "clientindex": "0",
             "clientname": "",
             "buname": "",
             "employeename": "",
             "gradename": "-",
             "designationname": "-",
             "employmenttypedesc": ""
          },
          {
             "employeeindex": 300000022,
             "clientindex": "0",
             "clientname": "",
             "buname": "",
             "employeename": "",
             "gradename": "-",
             "designationname": "-",
             "employmenttypedesc": ""
          },
          {
             "employeeindex": 300000023,
             "clientindex": "0",
             "clientname": "",
             "buname": "",
             "employeename": "",
             "gradename": "-",
             "designationname": "-",
             "employmenttypedesc": ""
          },
          {
             "employeeindex": 300000024,
             "clientindex": "0",
             "clientname": "",
             "buname": "",
             "employeename": "Test123abc",
             "gradename": "",
             "designationname": "",
             "employmenttypedesc": ""
          },
          {
             "employeeindex": 300000025,
             "clientindex": "0",
             "clientname": "",
             "buname": "",
             "employeename": "",
             "gradename": "-",
             "designationname": "-",
             "employmenttypedesc": ""
          },
          {
             "employeeindex": 300000026,
             "clientindex": "0",
             "clientname": "",
             "buname": "",
             "employeename": "",
             "gradename": "-",
             "designationname": "-",
             "employmenttypedesc": ""
          },
          {
             "employeeindex": 300000027,
             "clientindex": "0",
             "clientname": "",
             "buname": "",
             "employeename": "",
             "gradename": "-",
             "designationname": "-",
             "employmenttypedesc": ""
          },
          {
             "employeeindex": 300000028,
             "clientindex": "0",
             "clientname": "",
             "buname": "",
             "employeename": "",
             "gradename": "-",
             "designationname": "-",
             "employmenttypedesc": ""
          },
          {
             "employeeindex": 300000029,
             "clientindex": "0",
             "clientname": "",
             "buname": "",
             "employeename": "",
             "gradename": "-",
             "designationname": "-",
             "employmenttypedesc": ""
          },
          {
             "employeeindex": 300000030,
             "clientindex": "0",
             "clientname": "",
             "buname": "",
             "employeename": "Test abc khan",
             "gradename": "",
             "designationname": "",
             "employmenttypedesc": ""
          },
          {
            "employeeindex": 300000031,
             "clientindex": "0",
             "clientname": "",
             "buname": "",
             "employeename": "Test abc khan",
             "gradename": "",
             "designationname": "",
             "employmenttypedesc": ""
          },
          {
             "employeeindex": 300000032,
             "clientindex": "0",
             "clientname": "",
             "buname": "",
             "employeename": "Test abc khan",
             "gradename": "",
             "designationname": "",
             "employmenttypedesc": ""
          },
          {
             "employeeindex": 300000033,
             "clientindex": "0",
             "clientname": "",
             "buname": "",
             "employeename": "Test abc khan",
             "gradename": "",
             "designationname": "",
             "employmenttypedesc": ""
          },
          {
             "employeeindex": 300000034,
             "clientindex": "0",
             "clientname": "",
             "buname": "",
             "employeename": "Umair A Bukhari",
             "gradename": "",
             "designationname": "",
             "employmenttypedesc": ""
          },
          {
             "employeeindex": 300000035,
             "clientindex": "0",
             "clientname": "",
             "buname": "",
             "employeename": "",
             "gradename": "-",
             "designationname": "-",
             "employmenttypedesc": ""
          },
          {
             "employeeindex": 300000036,
             "clientindex": "0",
             "clientname": "",
             "buname": "",
             "employeename": "",
             "gradename": "-",
             "designationname": "-",
             "employmenttypedesc": ""
          },
          {
             "employeeindex": 300000037,
             "clientindex": "0",
             "clientname": "",
             "buname": "",
             "employeename": "Umair  Bukhari",
             "gradename": "",
             "designationname": "",
             "employmenttypedesc": ""
          },
          {
             "employeeindex": 300000038,
             "clientindex": "0",
             "clientname": "",
             "buname": "",
             "employeename": "",
             "gradename": "-",
             "designationname": "-",
             "employmenttypedesc": ""
          },
          {
             "employeeindex": 300000039,
             "clientindex": "0",
             "clientname": "",
             "buname": "",
             "employeename": "",
             "gradename": "-",
             "designationname": "-",
             "employmenttypedesc": ""
          },
          {
             "employeeindex": 300000040,
             "clientindex": "0",
             "clientname": "",
             "buname": "",
             "employeename": "",
             "gradename": "-",
             "designationname": "-",
             "employmenttypedesc": ""
          },
          {
             "employeeindex": 300000041,
             "clientindex": "0",
             "clientname": "",
             "buname": "",
             "employeename": "",
             "gradename": "-",
             "designationname": "-",
             "employmenttypedesc": ""
          },
          {
             "employeeindex": 300000042,
             "clientindex": "0",
             "clientname": "",
             "buname": "",
             "employeename": "",
             "gradename": "-",
             "designationname": "-",
             "employmenttypedesc": ""
          },
          {
             "employeeindex": 300000043,
             "clientindex": "0",
             "clientname": "",
             "buname": "",
             "employeename": "",
             "gradename": "-",
             "designationname": "-",
             "employmenttypedesc": ""
          },
          {
             "employeeindex": 300000044,
             "clientindex": "0",
             "clientname": "",
             "buname": "",
             "employeename": "",
             "gradename": "-",
             "designationname": "-",
             "employmenttypedesc": ""
          },
          {
             "employeeindex": 300000045,
             "clientindex": "0",
             "clientname": "",
             "buname": "",
             "employeename": "",
             "gradename": "-",
             "designationname": "-",
             "employmenttypedesc": ""
          },
          {
             "employeeindex": 300000046,
             "clientindex": "0",
             "clientname": "",
             "buname": "",
             "employeename": "",
             "gradename": "-",
             "designationname": "-",
             "employmenttypedesc": ""
          },
          {
             "employeeindex": 300000047,
             "clientindex": "0",
             "clientname": "",
             "buname": "",
             "employeename": "",
             "gradename": "-",
             "designationname": "-",
             "employmenttypedesc": ""
          },
          {
             "employeeindex": 300000048,
             "clientindex": "0",
             "clientname": "",
             "buname": "",
             "employeename": "",
             "gradename": "-",
             "designationname": "-",
             "employmenttypedesc": ""
          },
          {
             "employeeindex": 300000049,
             "clientindex": "0",
             "clientname": "",
             "buname": "",
             "employeename": "",
             "gradename": "-",
             "designationname": "-",
             "employmenttypedesc": ""
          },
          {
             "employeeindex": 300000050,
             "clientindex": "0",
             "clientname": "",
             "buname": "",
             "employeename": "",
             "gradename": "-",
             "designationname": "-",
             "employmenttypedesc": ""
          },
          {
             "employeeindex": 300000051,
             "clientindex": "0",
             "clientname": "",
             "buname": "",
             "employeename": "",
             "gradename": "-",
             "designationname": "-",
             "employmenttypedesc": ""
          },
          {
             "employeeindex": 300000052,
             "clientindex": "0",
             "clientname": "",
             "buname": "",
             "employeename": "",
             "gradename": "-",
             "designationname": "-",
             "employmenttypedesc": ""
          },
          {
             "employeeindex": 300000053,
             "clientindex": "0",
             "clientname": "",
             "buname": "",
             "employeename": "",
             "gradename": "-",
             "designationname": "-",
             "employmenttypedesc": ""
          }
       ],
      //  table_2: [
      //     {
      //        "employeeindex": 0,
      //        "clientindex": "0",
      //        "clientname": "",
      //        "buname": "",
      //        "employeename": "",
      //        "gradename": null,
      //        "designationname": "",
      //        "atId": "0",
      //        "approvalGroupEmploymentType": null,
      //        "clientBranchName": ""
      //     },
      //     {
      //        "employeeindex": 300000001,
      //        "clientindex": "300000001",
      //        "clientname": "HRSG BPO",
      //        "buname": "HRSG BPO",
      //        "employeename": "Umair Ahmed Bukhari",
      //        "gradename": null,
      //        "designationname": "",
      //        "atId": "0",
      //        "approvalGroupEmploymentType": null,
      //        "clientBranchName": ""
      //     }
      //  ] 
    },
    "header": {
       "code": 0,
       "message": "Success"
    }
  }
 

  // empPerformance = [
  //   { ID: 10011, NAME: "A", DEPARTMENT: "Sales", MONTH: "Jan", YEAR: 2020, SALES: 132412, CHANGE: 12, LEADS: 35 },
  //   { ID: 10012, NAME: "A", DEPARTMENT: "Sales", MONTH: "Feb", YEAR: 2020, SALES: 232324, CHANGE: 2, LEADS: 443 },
  //   { ID: 10013, NAME: "A", DEPARTMENT: "Sales", MONTH: "Mar", YEAR: 2020, SALES: 542234, CHANGE: 45, LEADS: 345 },
  //   { ID: 10014, NAME: "A", DEPARTMENT: "Sales", MONTH: "Apr", YEAR: 2020, SALES: 223335, CHANGE: 32, LEADS: 234 },
  //   { ID: 10015, NAME: "A", DEPARTMENT: "Sales", MONTH: "May", YEAR: 2020, SALES: 455535, CHANGE: 21, LEADS: 12 },
  // ];

  constructor(public ete: ExportExcelService) { }

  // exportToExcel() {

  //   this.empPerformance.forEach((row: any) => {
  //     this.dataForExcel.push(Object.values(row))
  //   })

  //   let reportData = {
  //     title: 'Employee Sales Report - Jan 2020',
  //     data: this.dataForExcel,
  //     headers: Object.keys(this.empPerformance[0])
  //   }

  //   this.ete.exportExcel(reportData);
  // }

  exportToExcel() {
    let headers = []
    this.dataForExcel = []

    for (let prop in this.empPerformance?.body) {    
      let obj = this.empPerformance.body[prop];
      let arr:any = []

      obj.forEach((row: any) => {
        // console.log(row)
        arr.push(Object.values(row))
      })
      
      // console.log(arr)
      for (let i=0;i<arr.length;i++) {
        if (this.dataForExcel?.[prop]) {
          console.log("inside if")
          console.log(arr[i])
          this.dataForExcel[prop].push(arr[i])
        }
        else {
          console.log("inside else")
          console.log(arr[i])
          this.dataForExcel = { ...this.dataForExcel, [prop]: [] };
          this.dataForExcel[prop].push(arr[i]);
        }
      }
      console.log(this.dataForExcel)

      headers.push(Object.keys(this.empPerformance?.body?.[prop]?.[0]));
    }
    // console.log(this.dataForExcel);
    // console.log(headers);

    let reportData = {
      title: 'Employee Sales Report - Jan 2020',
      data: this.dataForExcel,
      headers
    }

    this.ete.exportExcel(reportData);
  }

}
