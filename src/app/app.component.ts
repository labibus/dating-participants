import { Component } from '@angular/core';
import * as XLSX from 'xlsx';
import * as FileSaver from 'file-saver';

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.css']
})
export class AppComponent {
  title = 'dating';
originaldating : any;
dating : any;
bdd : any;
arr : any;
notExisitingUsers=[];
exisitingUsers=[];
duplicateArray = [];

constructor()
{
  //excel to json https://beautifytools.com/excel-to-json-converter.php
  this.bddJson();
  
}


public checkIfExists()
{
  this.dating.forEach(element => {
    
    const membre = this.bdd.find(e=>e.CIN ===element.CIN || e.Mail === element.Mail || e.Tel === element.Num);
    if(membre == undefined)
    {
      this.notExisitingUsers.push(element);        
    }else{
      element["BDDID"] = membre["ID"];
      this.exisitingUsers.push(element);
    }      
  });
  //json to excel https://json-csv.com/

  }

  public removeDuplicates(originalArray, prop) {
    var newArray = [];
    var lookupObject  = {};
    
    

    for(var i in originalArray) {
       lookupObject[originalArray[i][prop]] = originalArray[i];
    }

    for(i in lookupObject) {
        newArray.push(lookupObject[i]);
    }
     this.dating = newArray;
}

public getDuplicated()
{
  for (let index = 1; index < this.originaldating.length+1; index++) {
    const membre = this.dating.find(e => e.ID === index.toString());
    if(membre == undefined)
    {
      this.duplicateArray.push(this.originaldating.find(e=> e.ID === index.toString()));
    }
  }
}

  public datingJson() {
    fetch('../assets/dating.json').then(res => res.json())
      .then(json => {
        this.originaldating = json["Form Responses 1"];
        this.dating = json["Form Responses 1"];
        this.removeDuplicates(this.dating, "CIN");
        this.removeDuplicates(this.dating, "Mail");
        this.removeDuplicates(this.dating, "Num");
        this.getDuplicated();
        this.checkIfExists();
        console.log(JSON.stringify(this.duplicateArray));//removed duplicated
        console.log(JSON.stringify(this.notExisitingUsers));//not exisiting
        console.log(JSON.stringify(this.exisitingUsers));//exisiting
      });
  }

  public bddJson() {
    fetch('../assets/full bdd.json').then(res => res.json())
      .then(json => {
        this.bdd = json["Feuille 1"];
    this.datingJson();

      });
  }

  downloadExisitingUsers()
  {
    this.exportAsExcelFile(this.exisitingUsers, 'exisitingUsers');
    
  }
  downloadNotExisitingUsers()
  {
    this.exportAsExcelFile(this.notExisitingUsers, 'notExisitingUsers');
    
  }
  downloadDuplicate()
  {
    this.exportAsExcelFile(this.duplicateArray, 'duplicateArray');
    
  }
  
  //#region  excel
  exportAsXLSX():void {
    this.exportAsExcelFile(this.duplicateArray, 'duplicateArray');
    this.exportAsExcelFile(this.notExisitingUsers, 'notExisitingUsers');
    this.exportAsExcelFile(this.exisitingUsers, 'exisitingUsers');
  }
  organise(arr) {
    var headers = [], // an Array to let us lookup indicies by group
      objs = [],    // the Object we want to create
      i, j;
    for (i = 0; i < arr.length; ++i) {
      j = headers.indexOf(arr[i].id); // lookup
      if (j === -1) { // this entry does not exist yet, init
        j = headers.length;
        headers[j] = arr[i].id;
        objs[j] = {};
        objs[j].id = arr[i].id;
        objs[j].data = [];
      }
      objs[j].data.push( // create clone
        {
          case_worked: arr[i].case_worked,
          note: arr[i].note, id: arr[i].id
        }
      );
    }
    return objs;
  }

  
 EXCEL_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=UTF-8';
 EXCEL_EXTENSION = '.xlsx';
  public exportAsExcelFile(json: any[], excelFileName: string): void {
    
    const worksheet: XLSX.WorkSheet = XLSX.utils.json_to_sheet(json);
    console.log('worksheet',worksheet);
    const workbook: XLSX.WorkBook = { Sheets: { 'data': worksheet }, SheetNames: ['data'] };
    const excelBuffer: any = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });
    //const excelBuffer: any = XLSX.write(workbook, { bookType: 'xlsx', type: 'buffer' });
    this.saveAsExcelFile(excelBuffer, excelFileName);
  }

  private saveAsExcelFile(buffer: any, fileName: string): void {
    const data: Blob = new Blob([buffer], {
      type: this.EXCEL_TYPE
    });
    FileSaver.saveAs(data, fileName + this.EXCEL_EXTENSION);
  }
  //#endregion

}
