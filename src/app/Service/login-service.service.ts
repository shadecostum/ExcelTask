import { HttpClient } from '@angular/common/http';
import { Injectable } from '@angular/core';
import * as XLSX from 'xlsx'; 

@Injectable({
  providedIn: 'root'
})
export class LoginServiceService {

  private excelFile = 'assets/Calculation.xlsx'; // Adjust the path to your Excel file

  constructor(private http: HttpClient) {}

  readExcel(): Promise<any[]> {
    return this.http.get(this.excelFile, { responseType: 'arraybuffer' })
      .toPromise()
      .then((data: ArrayBuffer | undefined) => {
        const binaryData = new Uint8Array(data!); // Using the non-null assertion operator (!)
        const workbook = XLSX.read(binaryData, { type: 'array' });
        const sheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[sheetName];

        return XLSX.utils.sheet_to_json(sheet);
      });
  }

  writeExcel(data: any[]): void {
    const worksheet: XLSX.WorkSheet = XLSX.utils.json_to_sheet(data);
    const newWorkbook: XLSX.WorkBook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(newWorkbook, worksheet, 'Sheet1');
    XLSX.writeFile(newWorkbook, this.excelFile);
  }
}
