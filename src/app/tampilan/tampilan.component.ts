import { Component } from '@angular/core';
import { CommonModule } from '@angular/common';
import * as XLSX from 'xlsx';
import Swal from 'sweetalert2';

@Component({
  selector: 'app-tampilan',
  standalone: true,
  imports: [CommonModule],
  templateUrl: './tampilan.component.html',
  styleUrls: ['./tampilan.component.css'],
})
export class TampilanComponent {
  excelData: { [category: string]: { Group: string; NamaTamu: string; Phone: string,Pax: string, BookingCode: string, Email: string, Location: string  }[] } = {};

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
        } as any);
  
        const categorizedData: { [category: string]: { Group: string; NamaTamu: string; Phone: string; Email: string; Location: string; Pax: string; BookingCode: string, }[] } = {};
        jsonData.forEach((row) => {
          if (row['GROUP']) {
            const groupPrefix = row['GROUP'].trim();
            const pickupPoint = row['Pickup Point'] || '';
  
            // Extract phone number
            const phoneMatch = pickupPoint.match(/Phone:\s*([+\d]+)/);
            const phone = phoneMatch ? phoneMatch[1] : '';
  
            // Extract email
            const emailMatch = pickupPoint.match(/([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})/);
            const email = emailMatch ? emailMatch[1] : '';
  
            // Extract location
            const locationMatch = pickupPoint.match(/📍(.*)/);
            const location = locationMatch ? locationMatch[1].trim() : '';
  
            // Exclude rows starting with "ENP"
            if (groupPrefix.startsWith('ENP')) {
              return; // Skip this row
            }
  
            let currentCategory = '';
            if (groupPrefix.startsWith('EMP.')) {
              currentCategory = 'EAST & WEST SANUR MEETING POINT';
            } else if (groupPrefix.startsWith('ET.')) {
              currentCategory = 'EAST TOUR FROM BALI';
            } else if (groupPrefix.startsWith('A.') || groupPrefix.startsWith('NA.')) {
              currentCategory = 'ALL INCLUSIVE FROM BALI';
            } else if (groupPrefix.startsWith('WMP.')) {
              currentCategory = 'WEST SANUR MEETING POINT';
            } else if (groupPrefix.startsWith('W.')) {
              currentCategory = 'WEST TOUR FROM BALI';
            } else if (groupPrefix.startsWith('E.')) {
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
              Email: email,
              Location: location,
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
  

  sendMessage(row: { Group: string; NamaTamu: string; Phone: string; Pax: string; BookingCode: string, Email: string, Location: string }): void {
    const group = row.Group.trim();
    const isCategoryE = group.startsWith('E'); // Periksa apakah kategori dimulai dengan "E"
    
    let message = '';
  
    if (isCategoryE) {
      const bookingCode = row.BookingCode || 'Unknown Booking Code'; // Ambil Booking Code dari data
      const activityDate = this.getNextDayDate(); // Dapatkan tanggal besok
      const pax = parseInt(row.Pax, 10) || 0; // Konversi Pax ke angka, default 0 jika tidak valid
      const paxLabel = pax === 1 ? 'person' : 'people';
      const namaTamu = row.NamaTamu.replace(/\s+/g, ' ').trim();
  
      message = `
${row.Group},
  
Dear ${namaTamu}
  
Greetings from Trip Gotik Get Your Guide Local Partner. We are excited to inform you that your booking for the **Nusa Penida Trip** with the following details is confirmed:
  
* **Booking Code**  : ${bookingCode}
* **Activity Date** : ${activityDate}
* **Total Person**  : ${pax} ${paxLabel}
  
Please note that your pick-up time will be between **6:00-6:15 AM** from ${row.Location}. The driver will assist you with the check-in process at Bali Harbor. Please be informed that this is a group tour, and on rare occasions, some participants may not be punctual. However, rest assured that we will inform you in case of any delays during pick-up. Please don't worry, as you will still be picked up as scheduled.
  
For tomorrow, we are scheduled to depart at **07:30 AM** from **Sanur Port**. When you arrive in **Nusa Penida**, please be attentive and look for our team holding a white paper sign with your name on it. Your tour will be arranged by our team from this point onwards.
  
To ensure your comfort throughout the trip, we recommend:
- Wearing comfortable clothing and walking shoes/sneakers.
- Applying sunscreen and bringing sunglasses.
- Packing a swimsuit (you’ll have the chance to swim, especially when visiting Diamond Beach).
  
Please note:
- **Nusa Penida** is a relatively new destination, giving you a glimpse of Bali as it was 30 years ago. Around 20% of the roads are still bumpy, and public facilities are limited.
- Due to narrow roads, we may encounter some traffic jams while moving from one spot to another.
- Bring some extra cash for restroom usage and lunch. Local restaurants offer a variety of cuisines, including Indonesian, Western, and Chinese dishes.
  
Upon your return to **Sanur Harbor** (around **5:45-6:00 PM**), please make your way back to the ticket pick-up point. Your driver will be waiting there, ready to transport you back to your hotel.
  
If you have any questions or need further assistance regarding this booking, please feel free to contact us.
  
Thank you,
Karma
      `.trim();
    } else {
      message = `Halo ${row.NamaTamu}, saya ingin menghubungi Anda melalui informasi dari file Excel.`.trim();
    }
  
    // Salin pesan ke clipboard
    navigator.clipboard.writeText(message)
      .then(() => {
        console.log('Pesan berhasil disalin ke clipboard');
        // Buka WhatsApp tanpa pesan
        const whatsappUrl = `https://wa.me/${row.Phone}`;
        window.open(whatsappUrl, '_blank');
      })
      .catch(err => {
        console.error('Gagal menyalin pesan ke clipboard: ', err);
      });
  }

  sendEmail(row: any): void {
    const emailLink = `mailto:${row.Email}?subject=Booking%20Information&body=Dear%20${row.NamaTamu},%20...`;
    window.open(emailLink, '_blank');
  }
  
  showActionModal(row: any): void {
    Swal.fire({
      title: 'Pilih Aksi',
      text: `Apa yang ingin Anda lakukan untuk ${row.NamaTamu}?`,
      icon: 'question',
      showCancelButton: true,
      confirmButtonText: row.Phone ? 'Kirim WA' : 'WA Tidak Tersedia',
      denyButtonText: row.Email ? 'Kirim Email' : 'Email Tidak Tersedia',
      showDenyButton: true,
      confirmButtonColor: row.Phone ? '#3085d6' : '#d6d6d6', // Enable/disable button color
      denyButtonColor: row.Email ? '#e74c3c' : '#d6d6d6', // Enable/disable button color,
      allowOutsideClick: false,
    }).then((result) => {
      if (result.isConfirmed && row.Phone) {
        this.sendMessage(row); // Call your WA function
      } else if (result.isDenied && row.Email) {
        this.sendEmail(row); // Call your email function
      }
    });
  }
  
  getNextDayDate(): string {
    const tomorrow = new Date();
    tomorrow.setDate(tomorrow.getDate() + 1);
    const options: Intl.DateTimeFormatOptions = { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' };
    return tomorrow.toLocaleDateString('id-ID', options); // Format dalam Bahasa Indonesia
  }
  
  
}
