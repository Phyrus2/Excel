import { Component } from '@angular/core';
import { CommonModule } from '@angular/common';
import * as XLSX from 'xlsx';

@Component({
  selector: 'app-tampilan',
  standalone: true,
  imports: [CommonModule],
  templateUrl: './tampilan.component.html',
  styleUrls: ['./tampilan.component.css'],
})
export class TampilanComponent {
  excelData: { [category: string]: { Group: string; NamaTamu: string; Phone: string }[] } = {};

  onFileChange(event: Event): void {
    const target = event.target as HTMLInputElement;
    if (target?.files?.length) {
      const file = target.files[0];
      const reader = new FileReader();

      reader.onload = (e) => {
        const data = new Uint8Array(e.target?.result as ArrayBuffer);
        const workbook = XLSX.read(data, { type: 'array' });

        const sheetName = workbook.SheetNames[9]; // Sheet ke-14
        const sheet = workbook.Sheets[sheetName];

        const jsonData: any[] = XLSX.utils.sheet_to_json(sheet, {
          header: [
            'GROUP',
            'Pickup Time',
            'Nama Tamu',
            'ADDITIONAL INFO',
            'Booking Code',
            'Pax',
            'Pickup Point',
          ],
          range: 1,
          defval: '',
        });

        const categorizedData: { [category: string]: { Group: string; NamaTamu: string; Phone: string }[] } = {};
        jsonData.forEach((row) => {
          if (row['GROUP']) {
            const groupPrefix = row['GROUP'].trim();
            const pickupPoint = row['Pickup Point'] || '';
            const phoneMatch = pickupPoint.match(/Phone:\s*([+\d]+)/);
            const phone = phoneMatch ? phoneMatch[1] : '';

            // Exclude rows starting with "ENP"
            if (groupPrefix.startsWith('ENP')) {
              return; // Skip this row
            }

            let currentCategory = '';
            if (groupPrefix.startsWith('EMP')) {
              currentCategory = 'EAST & WEST SANUR MEETING POINT';
            } else if (groupPrefix.startsWith('ET')) {
              currentCategory = 'EAST TOUR FROM BALI';
            } else if (groupPrefix.startsWith('A') || groupPrefix.startsWith('NA')) {
              currentCategory = 'ALL INCLUSIVE FROM BALI';
            } else if (groupPrefix.startsWith('WMP')) {
              currentCategory = 'WEST SANUR MEETING POINT';
            } else if (groupPrefix.startsWith('W')) {
              currentCategory = 'WEST TOUR FROM BALI';
            } else if (groupPrefix.startsWith('E')) {
              currentCategory = 'EAST & WEST TOUR';
            } else {
              return; // Skip rows that don't match any of these categories
            }

            if (!categorizedData[currentCategory]) {
              categorizedData[currentCategory] = [];
            }

            categorizedData[currentCategory].push({
              Group: row['GROUP'],
              NamaTamu: row['Nama Tamu'],
              Phone: phone,
            });
          }
        });

        this.excelData = categorizedData;
        console.log('Excel Data:', this.excelData);
      };

      reader.readAsArrayBuffer(file);
    }
  }
}
