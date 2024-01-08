import { HttpClient } from '@angular/common/http';
import { Component } from '@angular/core';
import { FormGroup, FormBuilder, Validators } from '@angular/forms';
import { Router } from '@angular/router';
import { LoginServiceService } from '../Service/login-service.service';
import * as XLSX from 'xlsx'; 


@Component({
  selector: 'app-update-excel',
  templateUrl: './update-excel.component.html',
  styleUrls: ['./update-excel.component.css']
})
export class UpdateExcelComponent {



  form: FormGroup;
  userName1:any
  previous:any
  constructor(private formBuilder: FormBuilder,
    private excelService: LoginServiceService,
    private route :Router,
    private http:HttpClient) {

  this.previous=localStorage.getItem('val')

    
  this.userName1=  localStorage.getItem('user')
    this.form = this.formBuilder.group({
      num1: ['', Validators.required],
      num2: ['', Validators.required],
      sum: [{ value: '', disabled: true }]

    });
  }


  formLoad()
  {
    
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
    
    const A = this.form.get('num1')?.value;
    const B = this.form.get('num2')?.value;

    this.userName = localStorage.getItem('user');
    this.pass = localStorage.getItem('user');

    console.log(this.userName);
    const addData = {
      userName: this.userName,
      Password: this.pass,
      A,
      B,
      Result: A + B,
    };
    
    console.log("addData", addData);
    console.log("addData", addData.A);
    console.log("addData", addData.Password);
    console.log("addData", addData.userName);
   this.loadExcelFile(addData);



  }
}




logOut()
{

}


excelFileName = 'assets/FinalExcel.xlsx';
ExcelDataStore:any; 
newData:any





loadExcelFile(data: any): void {
  this.newData = data;

  // Load the existing workbook
  this.http.get(this.excelFileName ?? '', { responseType: 'arraybuffer' })
    .subscribe((existingData: ArrayBuffer) => {
      const existingBinaryData = new Uint8Array(existingData);
      const workbook = XLSX.read(existingBinaryData, { type: 'array' });

      // Assuming the first sheet is the target sheet in the existing workbook
      const existingSheetName = workbook.SheetNames[0];
     
      const existingSheet = workbook.Sheets[existingSheetName];
      console.log("existingSheetName",existingSheetName);
      console.log("existingSheet",existingSheet);

      const header = ["user_Name", "password", "A", "B", "Results"];

      // Convert the existing sheet data to JSON
      const existingSheetData = XLSX.utils.sheet_to_json(existingSheet, { header }) as { [x: string]: any }[];

      existingSheetData.forEach((row) => {
        console.log(row['user_Name'], "row username");

        if (row['user_Name'] === 'Akash') {
          row['password'] = this.newData.Password;
          row['A'] = this.newData.A;
          row['B'] = this.newData.B;
          row['Results'] = this.newData.A + this.newData.B;
        }



        const updatedSheet = XLSX.utils.json_to_sheet(existingSheetData, { skipHeader: true });

       
    

 // Create a new workbook and add the updated sheet to it
 const newWorkbook = XLSX.utils.book_new();
 console.log("a",newWorkbook);
 console.log("b",updatedSheet);
 console.log("c",existingSheetName);
 
 
            XLSX.utils.book_append_sheet(newWorkbook, updatedSheet, existingSheetName);

 // Write the new workbook to the Excel file
          XLSX.writeFile(newWorkbook, 'FinalExcel.xlsx');







      });




















    });
















}
 







}
 