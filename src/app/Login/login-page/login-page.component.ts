import { HttpClient } from '@angular/common/http';
import { Component, OnInit } from '@angular/core';
import { FormControl, FormGroup, Validators } from '@angular/forms';
import { Router } from '@angular/router';
import { LoginServiceService } from 'src/app/Service/login-service.service';
import * as XLSX from 'xlsx'; 

@Component({
  selector: 'app-login-page',
  templateUrl: './login-page.component.html',
  styleUrls: ['./login-page.component.css']
})
export class LoginPageComponent implements OnInit {
  excelData: any ;
 


constructor(private http:HttpClient,private route :Router){
  this.LoadExcelFile();
}


  LoginForm=new FormGroup(
    {
      username:new FormControl('',[Validators.required]),
      Password:new FormControl('',[Validators.required]),
  
    }
  )
  
  get emailValid()
  {
   return this.LoginForm.get('username')
  }
  get PasswordValid()
  {
    return this.LoginForm.get('Password')
  }


 

  ngOnInit() {
    
  }



  excelFileName = 'assets/FinalExcel.xlsx';
  ExcelDataStore:any; 
  loginForm = { username: '', password: '' };
 

 
  LoadExcelFile() {

    this.http.get(this.excelFileName, { responseType: 'arraybuffer' })
      .subscribe((data: ArrayBuffer) => {
        const binaryData = new Uint8Array(data);
        const workbook = XLSX.read(binaryData, { type: 'array' });
        const sheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[sheetName];
 
      
        const header = [ "user_Name", "password","A","B","Results"];
 
        
        this.ExcelDataStore = XLSX.utils.sheet_to_json(sheet, { header }).slice(1);
        console.log(this.ExcelDataStore,"fetched json Data");
        console.log(this.ExcelDataStore[0].Results,"fetched json Data");
      });
  }
 
  onSubmit() {
    const username = this.loginForm?.username || '';
    const password = this.loginForm?.password || '';
 
    const matchingUser = this.ExcelDataStore.find((user: User) =>
      user.user_Name?.trim().toLowerCase() === username.trim().toLowerCase() &&
      user.password?.trim() === password.trim()

    );
 
    if (matchingUser) {
      localStorage.setItem('user',username)
      localStorage.setItem('pas',password)
     const value=this.ExcelDataStore[0].Results
      localStorage.setItem('val',value)
      
      alert('Login successful!');
      this.route.navigateByUrl('/dashBoard')

    } else {
      alert('Check user name and password');
    }
  }

  
  
  

}


 
interface User {

  user_Name: string;
  password: string;


}