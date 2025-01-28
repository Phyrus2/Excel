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
  excelData: { [category: string]: { Group: string; NamaTamu: string; Phone: string,Pax: string, BookingCode: string  }[] } = {};

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

        const categorizedData: { [category: string]: { Group: string; NamaTamu: string; Phone: string, Pax: string, BookingCode: string }[] } = {};
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
              Pax: row['Pax'], // Ambil Pax
              BookingCode: row['Booking Code'], // Ambil Booking Code
            });
          }
        });

        this.excelData = categorizedData;
        console.log('Excel Data:', this.excelData);
      };

      reader.readAsArrayBuffer(file);
    }
  }

  sendMessage(row: { Group: string; NamaTamu: string; Phone: string; Pax: string; BookingCode: string }): void {
    const group = row.Group.trim();
    const isCategoryE = group.startsWith('E'); // Periksa apakah kategori dimulai dengan "E"
    
    let message = '';
  
    if (isCategoryE) {
      const bookingCode = row.BookingCode || 'Unknown Booking Code'; // Ambil Booking Code dari data
      const activityDate = this.getNextDayDate(); // Dapatkan tanggal besok
      const pax = row.Pax || '0'; // Ambil Pax dari data, gunakan 0 jika kosong
      
      message = `${row.Group},

      Dear ${row.NamaTamu},

      Greetings from Trip Gotik, Get Your Guide Local partner. We are excited to inform you that your booking for the Nusa Penida Trip with the following details is confirmed:

      * Booking Code: ${bookingCode}
      * Activity date: ${activityDate}
      * Total Person: ${pax}

      Please note that your pick-up time will be between 6:00-6:15 AM from ${bookingCode}. The driver will assist you with the check in process in Bali harbor. Please be informed that this is a group tour, and on rare occasions, some participants may not be punctual. However, rest assured that we will inform you in case of any delays when picking you up. Please don't worry, as you will still be picked up as scheduled

      For tomorrow we are scheduled depart at 07:30 AM from Sanur port. When you arrive in Nusa Penida, please be attentive and look for our team holding a white paper sign with your name on it. Your tour will be arranged by our team from this point onwards.

      To ensure your comfort throughout the trip, it is recommended that you wear comfortable clothing, walking shoes, sneakers, apply sunscreen, and bring sunglasses.  Feel free to bring your swimsuit. You will have the chance to go for a swim, especially when visiting Diamond Beach. 

      Nusa Penida is a relatively new destination that is not fully developed yet, giving you a glimpse of Bali as it was 30 years ago. Approximately 20% of the roads in Nusa Penida are still bumpy, and public facilities are limited. Due to the narrow roads, we may encounter some traffic jams while moving from one spot to another.

      Additionally, please bring some extra cash for restroom usage and lunch. The local restaurants offer a variety of food options, including Indonesian, Western, and Chinese cuisine.

      Upon your return to Sanur Harbor around 5:45- 6:00 PM, please make your way back to the ticket pick-up point. Your driver will be waiting there, ready to transport you back to your hotel.  

      If you have any questions or need further assistance regarding this booking, please feel free to contact us.

      Thank you,
      Karma`;
    } else {
      message = `Halo ${row.NamaTamu}, saya ingin menghubungi Anda melalui informasi dari file Excel.`;
    }
  
    const encodedMessage = encodeURIComponent(message.trim());
    const whatsappUrl = `https://wa.me/${row.Phone}?text=${encodedMessage}`;
    window.open(whatsappUrl, '_blank');
  }
  
  getNextDayDate(): string {
    const tomorrow = new Date();
    tomorrow.setDate(tomorrow.getDate() + 1);
    const options: Intl.DateTimeFormatOptions = { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' };
    return tomorrow.toLocaleDateString('id-ID', options); // Format dalam Bahasa Indonesia
  }
  
  
}
