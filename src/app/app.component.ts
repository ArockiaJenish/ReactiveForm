import { Component } from '@angular/core';
import { FormControl, FormGroup } from '@angular/forms';
import * as XLSX from 'xlsx';
import * as moment from 'moment';

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.css']
})
export class AppComponent {
  title = 'ReactiveForm';

  valid: boolean = false;
  f1: boolean = true;
  f2: boolean = true;

  fileForm = new FormGroup({
    firstFile: new FormControl(''),
    secondFile: new FormControl(''),
  });

  check(event: any) {
    console.log(this.fileForm);
    let file1 = this.fileForm.value.firstFile;
    let file2 = this.fileForm.value.secondFile;
    console.log(event.target.files[0]);

    if (file1?.endsWith("xls") && file2?.endsWith("xls"))
      this.valid = true;
    else
      this.valid = false;
  }

  check1(event: any) {
    console.log(event);
    let file;
    if (event.target.files.length > 0) {
      file = event.target.files[0].name;
      if (file.endsWith("xlsx"))
        this.f1 = false;
      else
        this.f1 = true;
    }
  }

  check2(event: any) {
    console.log(event);
    let file;
    if (event.target.files.length > 0) {
      file = event.target.files[0].name;
      if (file.endsWith("xlsx"))
        this.f2 = false;
      else
        this.f2 = true;
    }
  }

  check3(event: any) {
    /* wire up file reader */
    const target: DataTransfer = <DataTransfer>(event.target);
    if (target.files.length !== 1) {
      throw new Error('Cannot use multiple files');
    }
    const file: File = event.target.files[0];
    const reader: FileReader = new FileReader();
    reader.readAsBinaryString(target.files[0]); //we can give the file directly here.
    reader.onload = (e: any) => {

      /* create workbook */
      const binarystr: string = e.target.result;
      //console.log(binarystr);
      const wb: XLSX.WorkBook = XLSX.read(binarystr, { type: 'binary' });
      console.log(wb);

      /* selected the first sheet */
      const wsname: string = wb.SheetNames[0];
      console.log(wsname);
      const ws: XLSX.WorkSheet = wb.Sheets[wsname];
      console.log(ws);

      /* save data */
      const data = XLSX.utils.sheet_to_json(ws, { header: 1 }); // to get 2d array pass 2nd parameter as object {header: 1}
      console.log(data); // Data will be logged in array format containing objects
    };
  }

  backgroundColor() {

    var xls = new ActiveXObject("Excel.Application");
    xls.visible = true;
    xls.DisplayAlerts = false;
    var wb = xls.Workbooks.Open("assets/In.xls");
    xls.Range("A1", "B1").Interior.ColorIndex = 37;
    xls.Range("C1", "D1").Interior.ColorIndex = 37;
    xls.Range("A1:D1").Merge();
    wb.SaveAs("assets/In.xls");
    xls.Quit();

    // var Excel = new ActiveXObject("Excel.Application");
    // var fso = new ActiveXObject("Scripting.FileSystemObject");
    // var checkFile = fso.FileExists("C:\\Report.xlsx");
    // if (checkFile) {
    //   fso.DeleteFile("C:\\Report.xlsx", true);
    // }
    // var ExcelSheet = new ActiveXObject("Excel.Sheet");
    // ExcelSheet.ActiveSheet.Range("A1", "G1").Merge();
    // ExcelSheet.ActiveSheet.Range("A1").value = "ECN REPORT";
    // ExcelSheet.ActiveSheet.Range("A1").Font.Bold = true;
    // ExcelSheet.ActiveSheet.Range("A1").Font.Size = 24;
    // ExcelSheet.ActiveSheet.Range("A1").HorizontalAlignment = -4108;
    // ExcelSheet.ActiveSheet.Range("A1", "F1").Interior.ColorIndex = 2;
    // // xlEdgeLeft,xlEdgeTop,xlEdgeBottom,xlEdgeRight all set to continuous  borders
    // ExcelSheet.ActiveSheet.Range("A1", "F1").Borders(7).LineStyle = 1;
    // ExcelSheet.ActiveSheet.Range("A1", "F1").Borders(8).LineStyle = 1;
    // ExcelSheet.ActiveSheet.Range("A1", "F1").Borders(9).LineStyle = 1;
    // ExcelSheet.ActiveSheet.Range("A1", "F1").Borders(10).LineStyle = 1;
  }

}
