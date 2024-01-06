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
  constructor(private formBuilder: FormBuilder,
    private excelService: LoginServiceService,
    private route :Router,
    private http:HttpClient) {

  

    
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
loadExcelFile(data:any): void {
  this.newData=data
  this.http.get(this.excelFileName, { responseType: 'arraybuffer' })
    .subscribe((data: ArrayBuffer) => {
      const binaryData = new Uint8Array(data);
      const workbook = XLSX.read(binaryData, { type: 'array' });
      const sheetName = workbook.SheetNames[0];
      const sheet = workbook.Sheets[sheetName];

      const header = ["user_Name", "password", "A", "B", "Results"];

      this.ExcelDataStore = XLSX.utils.sheet_to_json(sheet, { header });

      this.ExcelDataStore.forEach((row: { [x: string]: any; }) => {
        console.log(row['user_Name'],"dddd");
        
        if (row['user_Name'] === 'Akash') {
          row['password']=this.newData.Password
          row['A'] = this.newData.A
          row['B'] = this.newData.B
          row['Results'] =this.newData.A+ this.newData.B
        }
      });
      const newWorksheet: XLSX.WorkSheet = XLSX.utils.json_to_sheet(this.ExcelDataStore);

      // Step 2: Add the new worksheet to the existing workbook
      const workbookWithUpdatedData: XLSX.WorkBook = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(workbookWithUpdatedData, newWorksheet, 'Sheet1');
      
      // Step 3: Write the updated workbook to binary data
      const updatedBinaryData = XLSX.write(workbookWithUpdatedData, { bookType: 'xlsx', type: 'array' });
      
      // Step 4: Save the binary data back to the Excel file
      this.saveUpdatedExcelFile(updatedBinaryData);





      // // Update the sheet with modified data
      // XLSX.utils.sheet_add_json(sheet, this.ExcelDataStore, { header, skipHeader: true });
      // const updatedBinaryData = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });
      // console.log(updatedBinaryData,"updatedBinaryData");
      

      // // Save the modified data back to the same Excel sheet
      // this.saveUpdatedExcelFile(updatedBinaryData);
    });
}

// saveUpdatedExcelFile(data: any): void {
//   const blob = new Blob([data], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
//   const url = window.URL.createObjectURL(blob);
//   const a = document.createElement('a');
//   a.href = url;
//   a.download = 'assets/FinalExcel.xlsx';
//   document.body.appendChild(a);
//   a.click();
//   document.body.removeChild(a);
//   window.URL.revokeObjectURL(url);
// }

saveUpdatedExcelFile(data: any): void {

  const blob = new Blob([data], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  const url = window.URL.createObjectURL(blob);

  const a = document.createElement('a');
  document.body.appendChild(a);

  a.href = url;
  a.download = 'UpdatedExcelFile.xlsx';

  // Display a button for the user to click and replace the existing file
  const replaceButton = document.createElement('button');
  replaceButton.textContent = 'Replace Existing File';
  replaceButton.addEventListener('click', () => {
    // Remove the button and download the updated file
    document.body.removeChild(replaceButton);
    a.click();
    document.body.removeChild(a);
    window.URL.revokeObjectURL(url);

    // Inform the user that they need to manually replace the existing file
    alert('Please manually replace the existing file in the assets folder with the downloaded file.');
  });

  document.body.appendChild(replaceButton);
}




}
 