import { Injectable } from '@angular/core';

@Injectable({
  providedIn: 'root'
})
export class Json2htmlService {

  constructor() { }

  json2Html(data: any, mytable: any) {
    let cols = [];
             
    for (let i = 0; i < data.length; i++) {
      for (let k in data[i]) {
        if (cols.indexOf(k) === -1) {
          // Push all keys to the array
          cols.push(k);
        }
      }
    }

    // Create a table element
    var table = document.createElement("table");
    table.setAttribute('style', "border-collapse: collapse;")
             
    // Create table row tr element of a table
    var tr = table.insertRow(-1);
     
    for (var i = 0; i < cols.length; i++) {
      // Create the table header th element
      var theader = document.createElement("th");
      theader.setAttribute('style', "border: 1px solid black;");
      theader.setAttribute('style', theader.getAttribute('style')+'; color: blue');

      theader.innerHTML = cols[i];
      // Append columnName to the table row
      tr.appendChild(theader);
    }
     
    // Adding the data to the table
    for (var i = 0; i < data.length; i++) {
      // Create a new row
      let trow = table.insertRow(-1);
      for (var j = 0; j < cols.length; j++) {
        var cell = trow.insertCell(-1);
        cell.setAttribute('style', "border: 1px solid black;");
        // Inserting the cell at particular place
        cell.innerHTML = data[i][cols[j]];
      }
    }

    // mytable.nativeElement.innerHTML = "";
    mytable.nativeElement.appendChild(table);
    
    let br = document.createElement("br");
    mytable.nativeElement.appendChild(br);
  }
}
