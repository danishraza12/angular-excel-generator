import { Injectable } from '@angular/core';
import { Workbook } from 'exceljs';
import * as fs from 'file-saver';
import { DatePipe } from '@angular/common';

@Injectable({
  providedIn: 'root'
})
export class ExportExcelService {

  constructor(private datePipe: DatePipe) { }

  exportExcel(excelData: any) {

    //Title, Header & Data
    const title = excelData.title;
    const header = excelData.headers
    const data = excelData.data;

    //Create a workbook with a worksheet
    let workbook = new Workbook();
    let worksheet = workbook.addWorksheet('Sales Data');


    //Add Row and formatting
    worksheet.mergeCells('C1', 'F4');
    let titleRow = worksheet.getCell('C1');
    titleRow.value = title
    titleRow.font = {
      name: 'Calibri',
      size: 16,
      underline: 'single',
      bold: true,
      color: { argb: '0085A3' }
    }
    titleRow.alignment = { vertical: 'middle', horizontal: 'center' }

    // Date
    worksheet.mergeCells('A1:B1');
    let d = new Date();
    let fromDate = d.getDate() + '/' + d.getMonth() + '/' + d.getFullYear();
    let toDate = d.getDate() + '/' + d.getMonth() + '/' + d.getFullYear();
    let dateCell = worksheet.getCell('A1');
    
    dateCell.value = "Report For: "  + fromDate + " - " + toDate;
    dateCell.font = {
      name: 'Calibri',
      size: 12,
      bold: true
    }
    dateCell.alignment = { vertical: 'middle', horizontal: 'center' }
    dateCell.border = {
      top: {style:'thin', color: {argb:'000000'}},
      left: {style:'thin', color: {argb:'000000'}},
      bottom: {style:'thin', color: {argb:'000000'}},
      right: {style:'thin', color: {argb:'000000'}}
    };

    //Add Image
    // let myLogoImage = workbook.addImage({
    //   base64: logo.imgBase64,
    //   extension: 'png',
    // });
    // worksheet.mergeCells('A1:B4');
    // worksheet.addImage(myLogoImage, 'A1:B4');

    //Blank Row 
    worksheet.addRow([]);

    let i=0;
    for (let prop in data) {
      //Adding Header Row
      let headerRow = worksheet.addRow(header[i]);
    
      headerRow.eachCell((cell, number) => {
        cell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: '4167B8' },
          bgColor: { argb: '' }
        }
        cell.font = {
          bold: true,
          color: { argb: 'FFFFFF' },
          size: 12
        }
        cell.alignment = { vertical: 'middle', horizontal: 'center' }
        cell.border = {
          top: {style:'thin', color: {argb:'000000'}},
          left: {style:'thin', color: {argb:'000000'}},
          bottom: {style:'thin', color: {argb:'000000'}},
          right: {style:'thin', color: {argb:'000000'}}
        };

        worksheet.getColumn(number).width = 17
      })

      // console.log(data[prop])
      // Adding Data with Conditional Formatting
      data[prop].forEach((d:any, i:number) => {
        let row = worksheet.addRow(d);
        row.alignment = { vertical: 'middle', horizontal: 'center' }

        let dCell = row.getCell(6);
        let color = 'FF99FF99';
        // if (dCell?.value! < 200000) {
        //   color = 'FF9999'
        // }

        dCell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: color }
        }
      });

      // worksheet.getColumn(3).width = 20;
      worksheet.addRow([]);
      i++;
    }

    // worksheet.getColumn(3).width = 20;
    // worksheet.addRow([]);

    //Footer Row
    let date = this.datePipe.transform(Date.now(), 'MMM d, y, h:mm:ss a');
    // let time = d.getHours() + ":" + d.getMinutes() + ":" + d.getMinutes();
    let footerRow = worksheet.addRow(['Employee Sales Report generated on: ' + date]);
    footerRow.getCell(1).fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: '228B22' }
    };
    footerRow.getCell(1).font = {
      bold: true,
      color: { argb: 'FFFFFF' },
      size: 12
    }
    footerRow.getCell(1).border = {
      top: {style:'thin', color: {argb:'000000'}},
      left: {style:'thin', color: {argb:'000000'}},
      bottom: {style:'thin', color: {argb:'000000'}},
      right: {style:'thin', color: {argb:'000000'}}
    };

    //Merge Cells
    worksheet.mergeCells(`A${footerRow.number}:D${footerRow.number}`);

    //Generate & Save Excel File
    workbook.xlsx.writeBuffer().then((data) => {
      let blob = new Blob([data], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      fs.saveAs(blob, title + '.xlsx');
    })

  }

  // exportExcel(excelData: any) {

  //   //Title, Header & Data
  //   const title = excelData.title;
  //   const header = excelData.headers
  //   const data = excelData.data;

  //   //Create a workbook with a worksheet
  //   let workbook = new Workbook();
  //   let worksheet = workbook.addWorksheet('Sales Data');


  //   //Add Row and formatting
  //   worksheet.mergeCells('C1', 'F4');
  //   let titleRow = worksheet.getCell('C1');
  //   titleRow.value = title
  //   titleRow.font = {
  //     name: 'Calibri',
  //     size: 16,
  //     underline: 'single',
  //     bold: true,
  //     color: { argb: '0085A3' }
  //   }
  //   titleRow.alignment = { vertical: 'middle', horizontal: 'center' }

  //   // Date
  //   worksheet.mergeCells('A1:B1');
  //   let d = new Date();
  //   let fromDate = d.getDate() + '/' + d.getMonth() + '/' + d.getFullYear();
  //   let toDate = d.getDate() + '/' + d.getMonth() + '/' + d.getFullYear();
  //   let dateCell = worksheet.getCell('A1');
    
  //   dateCell.value = "Report For: "  + fromDate + " - " + toDate;
  //   dateCell.font = {
  //     name: 'Calibri',
  //     size: 12,
  //     bold: true
  //   }
  //   dateCell.alignment = { vertical: 'middle', horizontal: 'center' }
  //   dateCell.border = {
  //     top: {style:'thin', color: {argb:'000000'}},
  //     left: {style:'thin', color: {argb:'000000'}},
  //     bottom: {style:'thin', color: {argb:'000000'}},
  //     right: {style:'thin', color: {argb:'000000'}}
  //   };

  //   //Add Image
  //   // let myLogoImage = workbook.addImage({
  //   //   base64: logo.imgBase64,
  //   //   extension: 'png',
  //   // });
  //   // worksheet.mergeCells('A1:B4');
  //   // worksheet.addImage(myLogoImage, 'A1:B4');

  //   //Blank Row 
  //   worksheet.addRow([]);

  //   //Adding Header Row
  //   let headerRow = worksheet.addRow(header);
  
  //   headerRow.eachCell((cell, number) => {
  //     cell.fill = {
  //       type: 'pattern',
  //       pattern: 'solid',
  //       fgColor: { argb: '4167B8' },
  //       bgColor: { argb: '' }
  //     }
  //     cell.font = {
  //       bold: true,
  //       color: { argb: 'FFFFFF' },
  //       size: 12
  //     }
  //     cell.alignment = { vertical: 'middle', horizontal: 'center' }
  //     cell.border = {
  //       top: {style:'thin', color: {argb:'000000'}},
  //       left: {style:'thin', color: {argb:'000000'}},
  //       bottom: {style:'thin', color: {argb:'000000'}},
  //       right: {style:'thin', color: {argb:'000000'}}
  //     };

  //     worksheet.getColumn(number).width = 17
  //   })

  //   // Adding Data with Conditional Formatting
  //   data.forEach((d:any, i:number) => {
  //     let row = worksheet.addRow(d);
  //     row.alignment = { vertical: 'middle', horizontal: 'center' }

  //     let dCell = row.getCell(6);
  //     let color = 'FF99FF99';
  //     if (dCell?.value! < 200000) {
  //       color = 'FF9999'
  //     }

  //     dCell.fill = {
  //       type: 'pattern',
  //       pattern: 'solid',
  //       fgColor: { argb: color }
  //     }
  //   });

  //   // worksheet.getColumn(3).width = 20;
  //   worksheet.addRow([]);

  //   //Footer Row
  //   let date = this.datePipe.transform(Date.now(), 'MMM d, y, h:mm:ss a');
  //   // let time = d.getHours() + ":" + d.getMinutes() + ":" + d.getMinutes();
  //   let footerRow = worksheet.addRow(['Employee Sales Report Generated on: ' + date]);
  //   footerRow.getCell(1).fill = {
  //     type: 'pattern',
  //     pattern: 'solid',
  //     fgColor: { argb: '228B22' }
  //   };
  //   footerRow.getCell(1).font = {
  //     bold: true,
  //     color: { argb: 'FFFFFF' },
  //     size: 12
  //   }
  //   footerRow.getCell(1).border = {
  //     top: {style:'thin', color: {argb:'000000'}},
  //     left: {style:'thin', color: {argb:'000000'}},
  //     bottom: {style:'thin', color: {argb:'000000'}},
  //     right: {style:'thin', color: {argb:'000000'}}
  //   };

  //   //Merge Cells
  //   worksheet.mergeCells(`A${footerRow.number}:D${footerRow.number}`);

  //   //Generate & Save Excel File
  //   workbook.xlsx.writeBuffer().then((data) => {
  //     let blob = new Blob([data], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  //     fs.saveAs(blob, title + '.xlsx');
  //   })

  // }
}
