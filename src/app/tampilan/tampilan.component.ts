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

        const sheetName = workbook.SheetNames[13]; // Sheet ke-14
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
        let currentCategory = '';

        jsonData.forEach((row) => {
          if (row['GROUP']?.startsWith('*')) {
            currentCategory = row['GROUP']?.replace(/\*/g, '').trim();
            if (!categorizedData[currentCategory]) {
              categorizedData[currentCategory] = [];
            }
          } else if (row['GROUP'] && currentCategory) {
            const pickupPoint = row['Pickup Point'] || '';
            const phoneMatch = pickupPoint.match(/Phone:\s*([+\d]+)/);
            const phone = phoneMatch ? phoneMatch[1] : '';

            categorizedData[currentCategory].push({
              Group: row['GROUP'],
              NamaTamu: row['Nama Tamu'],
              Phone: phone,
            });
          }
        });

        this.excelData = categorizedData;
      };

      reader.readAsArrayBuffer(file);
    }
  }
}
