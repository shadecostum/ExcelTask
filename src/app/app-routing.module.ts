import { NgModule } from '@angular/core';
import { RouterModule, Routes } from '@angular/router';
import { LoginPageComponent } from './Login/login-page/login-page.component';
import { DashboardComponent } from './dashboard/dashboard.component';
import { UpdateExcelComponent } from './update-excel/update-excel.component';

const routes: Routes = [
  {path:"",component:LoginPageComponent},
  {path:"dashBoard",component:UpdateExcelComponent}
];

@NgModule({
  imports: [RouterModule.forRoot(routes)],
  exports: [RouterModule]
})
export class AppRoutingModule { } 
