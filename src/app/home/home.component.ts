import { Component, OnInit } from '@angular/core';
// import { de, ro } from 'date-fns/locale';
import * as XLSX from 'xlsx';
import * as moment from 'moment';
import { FormControl, FormGroup } from '@angular/forms';

@Component({
  selector: 'app-home',
  templateUrl: './home.component.html',
  styleUrls: ['./home.component.css']
})
export class HomeComponent implements OnInit {

  public f1: boolean = true;
  public f2: boolean = true;

  public workbook: any;
  public worksheet: any;
  public rowMajor: any = [];
  public Data: any = [];
  public newData: any = [];
  public downloadSheet: any = [];
  public newDataHeader: any = [];
  public avgCal = 0;
  public sheetName: string = "";
  public isCalcualted: boolean = false;

  constructor() { }

  valid: boolean = false;

  fileForm = new FormGroup({
    firstFile: new FormControl(''),
    secondFile: new FormControl(''),
  });

  //event for in-file
  uploadInFile(event: any) {
    let file: File = event.target.files[0];
    console.log(file);
    // console.log(filesData[0]);
    // console.log(event.target.files.length);
    // const file:File = event.target.files[0];
    if (event.target.files.length == 1) {
      // this.readXlsx(event.taret.file[0].name,"In");
      // console.log(typeof(event.target.files[0]));
      this.f1 = false;
    } else {
      this.f1 = true;
    }
  }
  //event for out-file
  uploadOutFile(event: any) {
    //const file:File = event.target.files[0];
    //console.log(event.target.files);
    if (event.target.files.length == 1) {
      // this.readXlsx(event.taret.file[0].name,"Out");
      this.f2 = false;
      console.log(typeof (event.target.files[0]));
    } else {
      this.f2 = true;
    }
  }

  // calculateData(){
  // this.readXlsx('/assets/In.xlsx','In');
  // this.readXlsx('/assets/Out.xlsx','Out');
  // this.calculateHr();
  // }

  ngOnInit(): void {
    // this.readXlsx('/assets/In.xls','In');
    // this.readXlsx('/assets/Out.xls','Out');
    //this.calculateHr();
  }
  // readXlsx('','In');
  // readXlsx('/assets/Out.xlsx','Out');

  file1(event: any){
    let file: File = event.target.files[0];
    this.readXlsx(file, 'In');
  }

  file2(event: any){
    let file: File = event.target.files[0];
    this.readXlsx(file, 'Out');
    // this.calculateHr();
  }

  // check3(event: any) {
  //   /* wire up file reader */
  //   const target: DataTransfer = <DataTransfer>(event.target);
  //   if (target.files.length !== 1) {
  //     throw new Error('Cannot use multiple files');
  //   }
  //   const file: File = event.target.files[0];
  //   const reader: FileReader = new FileReader();
  //   reader.readAsBinaryString(target.files[0]); //we can give the file directly here.
  //   reader.onload = (e: any) => {

  //     /* create workbook */
  //     const binarystr: string = e.target.result;
  //     //console.log(binarystr);
  //     const wb: XLSX.WorkBook = XLSX.read(binarystr, { type: 'binary' });
  //     console.log(wb);

  //     /* selected the first sheet */
  //     const wsname: string = wb.SheetNames[0];
  //     console.log(wsname);
  //     const ws: XLSX.WorkSheet = wb.Sheets[wsname];
  //     console.log(ws);

  //     /* save data */
  //     const data = XLSX.utils.sheet_to_json(ws, { header: 1 }); // to get 2d array pass 2nd parameter as object {header: 1}
  //     console.log(data); // Data will be logged in array format containing objects
  //   };
  // }

  async readXlsx(file: File, Type: string) {

    

    this.workbook = {};
    this.worksheet = {};
    this.rowMajor = [];
    const reader: FileReader = new FileReader();
    reader.readAsBinaryString(file); //we can give the file directly here.
    reader.onload = (e: any) => {

      /* create workbook */
      const binarystr: string = e.target.result;
      //console.log(binarystr);
      const wb: XLSX.WorkBook = XLSX.read(binarystr, { type: 'binary' });
      //console.log(wb);

      /* selected the first sheet */
      const wsname: string = wb.SheetNames[0];
      console.log(wsname);
      const ws: XLSX.WorkSheet = wb.Sheets[wsname];
      //console.log(ws);

      /* save data */
      const data = XLSX.utils.sheet_to_json(ws, { header: 1 }); // to get 2d array pass 2nd parameter as object {header: 1}
      console.log(data); // Data will be logged in array format containing objects
      this.rowMajor = data;
    };

    // this.workbook = await fetch(fileName)
    //   .then(resp => resp.arrayBuffer())
    //   .then(buff => XLSX.read(buff))
    //   .catch(err => console.error(err))
    // console.log(this.workbook);

    // var Sheet = this.workbook.SheetNames;
    // console.log(Sheet);
    // this.worksheet = this.workbook.Sheets[Sheet[0]];
    // this.rowMajor = XLSX.utils.sheet_to_json(this.worksheet, { header: 1 });
    // console.log(this.rowMajor);

    if (Type === 'In') {
      this.rowMajor.map((r: any, i: number) => {
        if (r.length > 0) {
          if (i == 3) {
            let data = r.filter((r: any) => (r != undefined && r != "Su" && r != "Sa" && r != "" && r != null));
            this.avgCal = data.length;
            console.log(this.avgCal);
          }
          if (i >= 4) {
            let ps = this.toObject(this.rowMajor[2], r, 'In');
            this.Data.push(ps);
          }
        }
      });
    }

    if (Type === 'Out') {
      // this.newData = [];
      this.rowMajor.map((r: any, i: number) => {
        if (r.length > 0) {
          if (i >= 4) {
            let ps = this.toObject(this.rowMajor[2], r, 'Out');
            this.Data.map((d: any) => {
              if (d.No == ps.No) {
                let newObjectData: any = {};
                let mappingData = Object.keys(d);
                this.newDataHeader = mappingData;
                mappingData.map((k: any) => {
                  if (k !== "No" && k !== "Name") {
                    newObjectData[k] = {
                      "In": d[k].In,
                      "Out": ps[k].Out
                    };
                  } else {
                    newObjectData[k] = d[k];
                  }
                });
                this.newData.push(newObjectData);
              }
            })
          }
        }
      });
    }
    // console.log(this.newData);
  }

  toObject(key: any, value: any, type: string) {
    let rv: any = {};
    for (var i = 0; i < key.length; ++i) {
      if (i >= 2) {
        let valueArry: any[] = (typeof value[i] !== 'number') ? ((value[i]) ? value[i].split('\n') : []) : (moment(value[i]).format("hh:mm"));
        let InOut = {
          [type]: valueArry.filter(r => r !== "")
        }
        rv[key[i]] = valueArry.splice(valueArry.length - 1);
        rv[key[i]] = InOut;
      } else {
        rv[key[i].replace(".", "")] = (value[i]) ? value[i] : "";
      }
    }
    return rv;
  }

  calculateHr(event: any) {
    event.preventDefault();
    this.newData.map((r: any) => {
      let mappingData = Object.keys(r);
      r.AvarageTime = 0;
      r.TotalWorkingDays = 0;
      r.TotalWorkingHours = "0";
      let TotalWorkingHoursDurations = moment.duration();
      mappingData.map((dat: any) => {
        let duration: any;
        if (dat !== "No" && dat !== "Name") {
          let startTime = (r[dat].In.length > 0) ? r[dat].In[0] : "00:00";
          let endTime = (r[dat].Out.length > 0) ? r[dat].Out[r[dat].Out.length - 1] : "00:00";
          startTime = moment(r[dat].In[0], 'HH:mm');
          endTime = moment(r[dat].Out[r[dat].Out.length - 1], 'HH:mm');
          duration = moment.duration(endTime.diff(startTime));
          r[dat].totalHr = moment.utc(duration.asMilliseconds()).format("HH:mm");
          r[dat].totalHr = (r[dat].totalHr !== 'Invalid date') ? r[dat].totalHr : "00:00";
          r[dat].totalHrBackground = (r[dat].totalHr >= "09:00") ? "table-success" : ((r[dat].totalHr >= "08:00") ? "table-warning" : ((r[dat].totalHr == "00:00") ? "table-light" : "table-danger"));
          r.AvarageTime = (duration.isValid()) ? (r.AvarageTime + duration.asMilliseconds()) : r.AvarageTime;
          r.TotalWorkingDays = (duration.isValid()) ? (r.TotalWorkingDays + 1) : r.TotalWorkingDays;
          TotalWorkingHoursDurations.add(moment.duration(`${r[dat].totalHr}:00`));
        }
      });
      r.TotalWorkingHours = `${(TotalWorkingHoursDurations.asHours()).toString().split(".")[0]}:${TotalWorkingHoursDurations.minutes()}`;
      r.AvarageTime = ((r.AvarageTime) / r.TotalWorkingDays);
      r.AvarageTime = moment.utc(r.AvarageTime).format("HH:mm");
      r.AvarageTime = (r.AvarageTime !== 'Invalid date') ? r.AvarageTime : "00:00";
      r.OverAllTimeBackground = (r.AvarageTime >= "09:00") ? "table-success" : ((r.AvarageTime >= "08:00") ? "table-warning" : ((r.AvarageTime == "00:00") ? "table-light" : "table-danger"));
      return r;
    });
    this.newDataHeader.push('TotalWorkingDays');
    this.newDataHeader.push('TotalWorkingHours');
    this.newDataHeader.push('AvarageTime');
    this.changeIndex("Name");
    this.changeIndex("No");
    console.log('newData---->')
    console.log(this.newData);
    this.isCalcualted = true;
  }

  changeIndex(item: any) {
    const fromIndex = this.newDataHeader.indexOf(item); // 👉️ 0
    const toIndex = 0;

    const element = this.newDataHeader.splice(fromIndex, 1)[0];
    console.log(element); // ['css']

    this.newDataHeader.splice(toIndex, 0, element);
    console.log(this.newDataHeader);
  }

  changeIndexDownlaod(item: any) {
    const fromIndex = this.downloadSheet.indexOf(item); // 👉️ 0
    const toIndex = 0;

    const element = this.downloadSheet.splice(fromIndex, 1)[0];
    console.log(element); // ['css']

    this.downloadSheet.splice(toIndex, 0, element);
    console.log(this.downloadSheet);
  }

  downloadExcel() {
    this.newData.map((row: any) => {
      let data: any = {};
      this.newDataHeader.map((header: any) => {
        if (header !== "No" && header !== "Name" && header !== "AvarageTime" && header !== "TotalWorkingHours" && header !== "TotalWorkingDays") {
          data[header] = row[header].totalHr;
        } else {
          data[header] = row[header];
        }
      });
      this.downloadSheet.push(data);
    });
    console.log(this.downloadSheet);
    const worksheet = XLSX.utils.json_to_sheet(this.downloadSheet);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "TimeSheet");
    XLSX.writeFile(workbook, "23Jul2022-21Aug2022.xlsx");

  }

}

