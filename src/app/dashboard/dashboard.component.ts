import { Component } from '@angular/core';
import { FormGroup, FormBuilder, Validators } from '@angular/forms';
import { LoginServiceService } from '../Service/login-service.service';
import * as XLSX from 'xlsx'; 
import { saveAs } from 'file-saver';
import { Router } from '@angular/router';
import { HttpClient } from '@angular/common/http';


@Component({
  selector: 'app-dashboard',
  templateUrl: './dashboard.component.html',
  styleUrls: ['./dashboard.component.css']
})
export class DashboardComponent {

  form: FormGroup;
  userName1:any
  constructor(private formBuilder: FormBuilder,
    private excelService: LoginServiceService,
    private route :Router,
    private http:HttpClient) {

    this.LoadExcelFile();
    console.log("a",this.ExcelDataStore);
    
  this.userName1=  localStorage.getItem('user')
    this.form = this.formBuilder.group({
      num1: ['', Validators.required],
      num2: ['', Validators.required],
      sum: [{ value: '', disabled: true }]

    });
  }
  
sum:any
calculateSum(): void {
  const num1 = this.form.get('num1')?.value;
  const num2 = this.form.get('num2')?.value;

  if (num1 !== null && num2 !== null) {
    this.sum = num1 + num2;
    this.form.get('sum')?.setValue(this.sum);
  } else {
    this.form.get('sum')?.setValue('');
  }
}
userName:any
pass:any


onSubmit(): void {
  if (this.form.valid) {
    const formData = {
      num1: this.form.value.num1,
      num2: this.form.value.num2,
    };
    
    const num1 = this.form.get('num1')?.value;
    const num2 = this.form.get('num2')?.value;

    this.userName = localStorage.getItem('user');
    this.pass = localStorage.getItem('user');

    console.log(this.userName);
    const addData = {
      userName: this.userName,
      Password: this.pass,
      num1,
      num2,
      result: num1 + num2,
    };
    
    console.log("addData", addData);

    this.readExcelFromAssets().then((originalData) => {
      this.updateExcel(addData, originalData);
    });
  }
}
 

  
  generateExcel(data: any) {

    
    console.log("dataFetched",this.ExcelDataStore);
    
    
    // Read existing data from the file
    this.readExcelFile().then(existingData => {

      // Remove existing data if it exists
      const newData = existingData.filter(item => item.userName !== data.userName);
  console.log(newData,"ddd");
  
      // Append new data to the existing data
      newData.push(data);

      console.log(newData,"dsa");
      
  
      // Write the combined data to the Excel file
      const worksheet: XLSX.WorkSheet = XLSX.utils.json_to_sheet(newData);
      const workbook: XLSX.WorkBook = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(workbook, worksheet, 'Sheet1');
      XLSX.writeFile(workbook, this.fileName);
    }).catch(error => {
      console.error('Error reading Excel file:', error);
    });
  }
  
  readExcelFile(): Promise<any[]> {
    // Read the existing data from the Excel file
    const fileInput = document.getElementById('assets/Calculation.xlsx') as HTMLInputElement;
    console.log(fileInput,"file");
    
    const file = fileInput?.files?.[0];
  
    if (file) {
      return new Promise<any[]>((resolve, reject) => {
        const reader = new FileReader();
  
        reader.onload = (e: any) => {
          const data = e.target.result;
          const workbook: XLSX.WorkBook = XLSX.read(data, { type: 'binary' });
          const jsonData = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]]);
          resolve(jsonData);
        };
  
        reader.onerror = (error) => {
          reject(error);
        };
  
        reader.readAsBinaryString(file);
      });
    } else {
      return Promise.resolve([]);
    }
  }
  
  
  fileName = 'assets/CRUD.xlsx';



  logOut()
  {
   this.route.navigateByUrl("/") 

  }






///////fetch
ExcelDataStore:any
LoadExcelFile() {

  this.http.get(this.fileName, { responseType: 'arraybuffer' })
    .subscribe((data: ArrayBuffer) => {
      const binaryData = new Uint8Array(data);
      const workbook = XLSX.read(binaryData, { type: 'array' });
      const sheetName = workbook.SheetNames[0];
      const sheet = workbook.Sheets[sheetName];

      console.log("sheet",sheet);
      

    
      const header = ["userName","Password", "A", "B","Result",];

      
      this.ExcelDataStore = XLSX.utils.sheet_to_json(sheet, { header });
      console.log(this.ExcelDataStore,"dd");
      
    });
}




////////////////////////////////////////////////////////////////////


updateExcel(updatedData: any, originalData: any[]): void {
  const workbook: XLSX.WorkBook = XLSX.utils.book_new();
  const sheet: XLSX.WorkSheet = XLSX.utils.json_to_sheet(originalData);

  // Update the sheet with the new data
  Object.keys(updatedData).forEach((key, index) => {
    const rowIndex = index + 2; // Assuming the header is in the first row
    const cellAddress = XLSX.utils.encode_cell({ r: rowIndex, c: 0 });
    sheet[cellAddress] = { t: 's', v: updatedData[key] };
  });

  XLSX.utils.book_append_sheet(workbook, sheet, 'Sheet1');
  XLSX.writeFile(workbook, 'updated_file.xlsx');
}



















newFile = 'assets/CRUD.xlsx';

readExcelFromAssets(): Promise<any[]> {
  const filePath = this.newFile; // Remove the `assets/` prefix
  return this.http.get(filePath, { responseType: 'blob' })
    .toPromise()
    .then((blob: Blob | undefined) => {
      if (blob) {
        return this.readExcel(blob);
      } else {
        console.error('Error: Blob is undefined');
        return [];
      }
    })
    .catch(error => {
      console.error('Error reading Excel file:', error);
      return [];
    });
}



private readExcel(file: Blob): Promise<any[]> {
  const reader: FileReader = new FileReader();
  let result: any[];

  return new Promise((resolve, reject) => {
    reader.onload = (e: any) => {
      const data: string = e.target.result;
      const workbook: XLSX.WorkBook = XLSX.read(data, { type: 'binary' });
      const sheetName: string = workbook.SheetNames[0];
      const worksheet: XLSX.WorkSheet = workbook.Sheets[sheetName];
      
      // Set header: 1 to treat the first row as headers
      result = XLSX.utils.sheet_to_json(worksheet, { raw: true, header: 1 }) as any[];
      resolve(result);
    };

    reader.onerror = (error) => {
      console.error('Error reading Excel file:', error);
      reject(error);
    };

    reader.readAsBinaryString(file);
  });
}










}




