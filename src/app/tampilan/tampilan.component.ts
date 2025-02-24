import { Component , OnInit} from '@angular/core';
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
export class TampilanComponent implements OnInit {
  excelData: { [category: string]: { Group: string; NamaTamu: string; Phone: string,Pax: string, BookingCode: string, Email: string, Location: string, PickupTime: string, AdditionalInfo: string  }[] } = {};
  bookingDate: string = '';

  ngOnInit() {
    this.askWhatsAppPreference(); // Selalu tanyakan pilihan saat website dibuka
}

  askWhatsAppPreference() {
      Swal.fire({
          title: 'Pilih Metode WhatsApp',
          text: 'Ingin mengirim pesan via WhatsApp Web atau Aplikasi?',
          icon: 'question',
          showCancelButton: true,
          confirmButtonText: 'WhatsApp Web',
          cancelButtonText: 'WhatsApp App',
      }).then((result) => {
          if (result.isConfirmed) {
              sessionStorage.setItem('whatsappPreference', 'web');
          } else {
              sessionStorage.setItem('whatsappPreference', 'app');
          }
      });
  }

  onFileChange(event: Event): void {
    const target = event.target as HTMLInputElement;
    if (target?.files?.length) {
      const file = target.files[0];
      const reader = new FileReader();
      this.bookingDate = this.getNextDayDate();
  
      reader.onload = (e) => {
        const data = new Uint8Array(e.target?.result as ArrayBuffer);
        const workbook = XLSX.read(data, { type: 'array' });
        const today = new Date();
        const tomorrow = today.getDate();
        const sheetName = workbook.SheetNames[tomorrow]; // Sheet ke-14
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
  
        const categorizedData: { [category: string]: { Group: string; NamaTamu: string; Phone: string; Email: string; Location: string; Pax: string; BookingCode: string, PickupTime: string, AdditionalInfo: string }[] } = {};
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
            } else if (groupPrefix.startsWith('W.')|| groupPrefix.startsWith('NW.')) {
              currentCategory = 'WEST TOUR FROM BALI';
            } else if (groupPrefix.startsWith('E.')|| groupPrefix.startsWith('PE.')) {
              currentCategory = 'EAST & WEST TOUR';
            } else if (groupPrefix.startsWith('SW.')|| groupPrefix.startsWith('SMP.')) {
              currentCategory = 'SNORKELING MANTA POINT & WEST COAST TOUR';
            }
            else if (groupPrefix.startsWith('IBH.')) {
              currentCategory = 'ISLAND BEACH HIGHLIGHTS SWIM & HIKE TOUR';
            }
            else {
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
              PickupTime: row['Pickup Time'],
              AdditionalInfo: row['ADDITIONAL INFO']
            });
          }
        });
  
        this.excelData = categorizedData;
        if (Object.keys(this.excelData).length === 0) {
          Swal.fire({
            icon: 'warning',
            title: 'Data Tour Tidak Ditemukan',
            text: 'Pastikan file yang diunggah berisi data yang sesuai.',
          });
          return;
        }      
        console.log('Excel Data:', this.excelData);
      };
  
      reader.readAsArrayBuffer(file);
    }
  }
  

  




  sendMessage(row: { Group: string; NamaTamu: string; Phone: string; Pax: string; BookingCode: string, Email: string, Location: string, PickupTime: string, AdditionalInfo: string }): void {
    const preference = sessionStorage.getItem('whatsappPreference'); // Ambil pilihan pengguna
    if (!preference) {
        console.error('WhatsApp preference not set.');
        return;
    }
    const group = row.Group.trim();
    const isCategoryE = group.startsWith('E.')|| group.startsWith('PE.');
    const isCategoryEMP = group.startsWith('EMP.');
    const isCategoryET = group.startsWith('ET.');
    const isCategoryAorNA = group.startsWith('A.') || group.startsWith('NA.');
    const isCategoryWMP = group.startsWith('WMP.');
    const isCategoryW = group.startsWith('W.') || group.startsWith('NW.'); // Periksa apakah kategori dimulai dengan "E"
    const isCategorySW = group.startsWith('SW.');
    const isCategoryIBH = group.startsWith('IBH.');
    const isCategorySMP = group.startsWith('SMP.');
    let message = '';
    const bookingCode = row.BookingCode || 'Unknown Booking Code'; // Ambil Booking Code dari data
      const activityDate = this.getNextDayDate(); // Dapatkan tanggal besok
      const pax = parseInt(row.Pax, 10) || 0; // Konversi Pax ke angka, default 0 jika tidak valid
      const paxLabel = pax === 1 ? 'person' : 'people';
      const namaTamu = row.NamaTamu.replace(/\s+/g, ' ').trim().replace(/\s*\(.*?\)\s*/g, '');
      const pickupTimeStr = String(row.PickupTime); // Contoh: "6.00"
      let [hours, minutes] = pickupTimeStr.split('.').map(Number); // Pisahkan jam dan menit

      minutes += 10; // Tambah 10 menit

      if (minutes >= 60) { 
          hours += Math.floor(minutes / 60); // Tambahkan ke jam jika lebih dari 60 menit
          minutes %= 60; // Sisakan menit yang tersisa
      }

      // Format hasil agar menit tetap dua digit
      const pickupTimeUpdated = `${hours}.${minutes.toString().padStart(2, '0')}`;
      const additionalInfo = row.AdditionalInfo || '';
      const polaroidMatch = additionalInfo.match(/Polaroid.*?\)/); // Ambil teks "Polaroid" sampai tanda tutup kurung
      const polaroidData = polaroidMatch ? polaroidMatch[0] : null;
      
    if (isCategoryE) {
      
  
      message = `
${row.Group},
  
Dear Mr./Mrs ${namaTamu}
  
Greetings from Trip Gotik Get Your Guide Local Partner. We are excited to inform you that your booking for the Nusa Penida Trip with the following details is confirmed:
  
* Booking Code  : ${bookingCode}
* Activity Date : ${activityDate}
* Total Person  : ${pax} ${paxLabel}
${polaroidData ? `* Add on       : ${polaroidData}\n` : ''}
Please note that your pick-up time will be between ${row.PickupTime} - ${pickupTimeUpdated} AM from ${row.Location}. The driver will assist you with the check in process in Bali harbor. Please be informed that this is a group tour, and on rare occasions, some participants may not be punctual. However, rest assured that we will inform you in case of any delays when picking you up. Please don't worry, as you will still be picked up as scheduled

For tomorrow we are scheduled depart at 07:30 AM from Sanur port. When you arrive in Nusa Penida, please be attentive and look for our team holding a white paper sign with your name on it. Your tour will be arranged by our team from this point onwards.

To ensure your comfort throughout the trip, it is recommended that you wear comfortable clothing, walking shoes, sneakers, apply sunscreen, and bring sunglasses.  Feel free to bring your swimsuit. You will have the chance to go for a swim, especially when visiting Diamond Beach. 

Nusa Penida is a relatively new destination that is not fully developed yet, giving you a glimpse of Bali as it was 30 years ago. Approximately 20% of the roads in Nusa Penida are still bumpy, and public facilities are limited. Due to the narrow roads, we may encounter some traffic jams while moving from one spot to another.

Additionally, please bring some extra cash for restroom usage and lunch. The local restaurants offer a variety of food options, including Indonesian, Western, and Chinese cuisine.

Upon your return to Sanur Harbor around 5:45- 6:00 PM, please make your way back to the ticket pick-up point. Your driver will be waiting there, ready to transport you back to your hotel.  

If you have any questions or need further assistance regarding this booking, please feel free to contact us.

Thank you,
Karma
      `.trim();
    } 

    else if (isCategoryEMP) {
     
  
      if (row.AdditionalInfo.toLowerCase() === "nusa team") {
        message = `
${row.Group},
          
Dear Mr./Mrs ${namaTamu}
          
Greetings from Trip Gotik Get Your Guide Local Partner. We are excited to inform you that your booking for the Nusa Penida Trip with the following details is confirmed:
          
* Booking Code  : ${bookingCode}
* Activity Date : ${activityDate}
* Total Person  : ${pax} ${paxLabel}
${polaroidData ? `* Add on       : ${polaroidData}\n` : ''}
Please note that you must arrive at Sanur Matahari Terbit Harbor at 8:00 AM for 8:30 AM boat departure. Our leader, Mr. Galung, can be reached at +62 813-5312-3400 and will assist you with the check-in process. If you have any trouble communicating with our leader, please head to the WIJAYA BUYUK FAST BOAT COUNTER located next to COCO MART EXPRESS for further assistance.

During the tour, you will be accompanied by a tour leader. When you arrive in Nusa Penida, please be attentive and look for our team holding a white paper sign with your name on it. Your tour will be arranged by our team from this point onwards.

To ensure your comfort throughout the trip, it is recommended that you wear comfortable clothing, Swimming suit, towel, walking shoes or flip-flops, apply sunscreen, and bring sunglasses. Nusa Penida is a relatively new destination that is not fully developed yet, giving you a glimpse of Bali as it was 30 years ago. Approximately 20% of the roads in Nusa Penida are still bumpy, and public facilities are limited. Due to the narrow roads, we may encounter some traffic jams while moving from one spot to another.

Enjoy some leisure time at the beautiful Crystal Bay Beach! Relax on the sand and take a refreshing swim. Remember to bring your swimsuit. Paid changing rooms and showers are available for your convenience

Additionally, please bring some extra cash for restroom usage, shower facilities, and lunch. The local restaurants offer a variety of food options, including Indonesian, Western, and Chinese cuisine.

If you have any questions or need further assistance regarding this booking, please feel free to contact us.

Thank you,
Karma

Coco Express Pantai Matahari Terbit:https://maps.app.goo.gl/ifDWG2fmto3LEMCF8 
      `.trim();
      }
            
      
      else{
        message = `
${row.Group},
  
Dear Mr./Mrs ${namaTamu}
  
Greetings from Trip Gotik Get Your Guide Local Partner. We are excited to inform you that your booking for the Nusa Penida Trip with the following details is confirmed:
  
* Booking Code  : ${bookingCode}
* Activity Date : ${activityDate}
* Total Person  : ${pax} ${paxLabel}
${polaroidData ? `* Add on       : ${polaroidData}\n` : ''}
Please note that you must arrive in THE ANGKAL FAST BOAT office, which is located directly next to CK Mart Matahari Terbit (address provided in the link below), at Sanur Matahari Terbit Harbor at 7:00 AM for the 7:30 AM boat departure. Should you encounter any difficulties with the timing or have trouble locating the office, kindly contact this number for immediate assistance.

During the tour, you will be accompanied by a tour leader. When you arrive in Nusa Penida, please be attentive and look for our team holding a white paper sign with your name on it. Your tour will be arranged by our team from this point onwards.

To ensure your comfort throughout the trip, it is recommended that you wear comfortable clothing, Swimming suit, towel, walking shoes or flip-flops, apply sunscreen, and bring sunglasses. Nusa Penida is a relatively new destination that is not fully developed yet, giving you a glimpse of Bali as it was 30 years ago. Approximately 20% of the roads in Nusa Penida are still bumpy, and public facilities are limited. Due to the narrow roads, we may encounter some traffic jams while moving from one spot to another.

Enjoy some leisure time at the beautiful Crystal Bay Beach! Relax on the sand and take a refreshing swim. Remember to bring your swimsuit. Paid changing rooms and showers are available for your convenience

Additionally, please bring some extra cash for restroom usage, shower facilities, and lunch. The local restaurants offer a variety of food options, including Indonesian, Western, and Chinese cuisine.

If you have any questions or need further assistance regarding this booking, please feel free to contact us.

Thank you,
Karma

CK Mart Matahari Terbit: https://maps.app.goo.gl/W4Y8V1NBk354mSi16 
      `.trim();
    }
  }

    else if (isCategoryAorNA) {
      
  
      message = `
${row.Group},
  
Dear Mr./Mrs ${namaTamu}
  
Greetings from Trip Gotik Get Your Guide Local Partner. We are excited to inform you that your booking for the Nusa Penida Trip with the following details is confirmed:
  
* Booking Code  : ${bookingCode}
* Activity Date : ${activityDate}
* Total Person  : ${pax} ${paxLabel}
${polaroidData ? `* Add on       : ${polaroidData}\n` : ''}
Please note that your pick-up time will be between ${row.PickupTime} - ${pickupTimeUpdated} AM at ${row.Location}. Upon arrival at your hotel, our driver will contact you. The driver will assist you with the check in process in Bali harbor. 

Sorry for this inconvenience that we have to pick you up slightly earlier than usual, cause of pickup distance from the harbor and for tomorrow we are scheduled depart at 07:00 AM from Sanur port, so that we can arrive in Nusa Penida earlier and cover all the destinations as planned specially for taking photo at tree house. Please be informed, this is a group tour and on rare occasions, some participants may not be punctual. However, rest assured that we will inform you in case of any delays when picking you up. Please don't worry, as you will still be picked up as scheduled. 

When you arrive in Nusa Penida, please be attentive and look for our team holding a white paper sign with your name on it. Your tour will be arranged by our team from this point onwards. 

If you would like to spend extra time swimming at a beach, you will have that time when you visit Diamond Beach. To ensure your comfort throughout the trip, it is recommended that you wear comfortable clothing, walking shoes, apply sunscreen, and bring sunglasses. 

Nusa Penida is a relatively new destination that is not fully developed yet, giving you a glimpse of Bali as it was 30 years ago. Approximately 20% of the roads in Nusa Penida are still bumpy, and public facilities are limited. Due to the narrow roads, we may encounter some traffic jams while moving from one spot to another.

Additionally, please bring some extra cash for restroom usage and lunch. The local restaurants offer a variety of food options, including Indonesian, Western, and Chinese cuisine.

Upon your return to Sanur Harbor around 5:30 PM, please make your way back to the ticket pick-up point. Your driver will be waiting there, ready to transport you back to your hotel.

If you have any questions or need further assistance regarding this booking, please feel free to contact us.

Thank you,
Karma 
      `.trim();
    }

    else if (isCategoryW) {
      
  
      message = `
${row.Group},
  
Dear Mr./Mrs ${namaTamu}
  
Greetings from Trip Gotik Get Your Guide Local Partner. We are excited to inform you that your booking for the Nusa Penida Trip with the following details is confirmed:
  
* Booking Code  : ${bookingCode}
* Activity Date : ${activityDate}
* Total Person  : ${pax} ${paxLabel}
${polaroidData ? `* Add on       : ${polaroidData}\n` : ''}
Please note that your pick-up time will be between ${row.PickupTime} - ${pickupTimeUpdated} AM from ${row.Location}, Jalan Hard Rock Hotel Kabupaten Badung. Upon arrival at your hotel, our driver will contact you. The driver will assist you with the check-in process at Bali harbor.

We apologize for the slight adjustment in the pick-up time due to  The scheduled departure time from Sanur Port is at 7:30 AM so that we can arrive in Nusa Penida earlier and cover all the destinations as planned. Please be informed that this is a group tour, and on rare occasions, some participants may not be punctual. However, rest assured that we will inform you in case of any delays when picking you up, and you will still be picked up as scheduled.

When you arrive in Nusa Penida, please be attentive and look for our team holding a white paper sign with your name on it. Your tour will be arranged by our team from this point onwards.

To ensure your comfort throughout the trip, it is recommended that you wear comfortable clothing, a swimsuit, towel, walking or trekking shoes, apply sunscreen, and bring sunglasses. Nusa Penida is a relatively new destination that is not fully developed yet, giving you a glimpse of Bali as it was 30 years ago. Approximately 20% of the roads in Nusa Penida are still bumpy, and public facilities are limited. Due to the narrow roads, we may encounter some traffic jams while moving from one spot to another.

Enjoy some leisure time at the beautiful Crystal Bay Beach! Relax on the sand and take a refreshing swim. Remember to bring your swimsuit. Paid changing rooms and showers are available for your convenience.

Additionally, please bring some extra cash for restroom usage, showers, and lunch. The local restaurants offer a variety of food options, including Indonesian, Western, and Chinese cuisine.

For your return journey, the boat is scheduled to depart from Nusa Penida at 4:00 PM/ 4:30 PM, and you will arrive back in Bali around 1 hour after. Upon your return to Sanur Harbor, please make your way back to the ticket pick-up point. Your driver will be waiting there, ready to transport you back to your hotel.

If you have any questions or need further assistance regarding this booking, please feel free to contact us.

Thank you,
Karma
      `.trim();
    }

    else if (isCategoryWMP) {
      
      if (row.AdditionalInfo.toLowerCase() === "nusa team") {
        message = `
${row.Group},
          
Dear Mr./Mrs ${namaTamu}
          
Greetings from Trip Gotik Get Your Guide Local Partner. We are excited to inform you that your booking for the Nusa Penida Trip with the following details is confirmed:
          
* Booking Code  : ${bookingCode}
* Activity Date : ${activityDate}
* Total Person  : ${pax} ${paxLabel}
${polaroidData ? `* Add on       : ${polaroidData}\n` : ''}
Please note that you must arrive at Sanur Matahari Terbit Harbor at 8:00 AM for 8:30 AM boat departure. Our leader, Mr. Galung, can be reached at +62 813-5312-3400 and will assist you with the check-in process. If you have any trouble communicating with our leader, please head to the WIJAYA BUYUK FAST BOAT COUNTER located next to COCO MART EXPRESS for further assistance.

During the tour, you will be accompanied by a tour leader. When you arrive in Nusa Penida, please be attentive and look for our team holding a white paper sign with your name on it. Your tour will be arranged by our team from this point onwards.

To ensure your comfort throughout the trip, it is recommended that you wear comfortable clothing, Swimming suit, towel, walking shoes or flip-flops, apply sunscreen, and bring sunglasses. Nusa Penida is a relatively new destination that is not fully developed yet, giving you a glimpse of Bali as it was 30 years ago. Approximately 20% of the roads in Nusa Penida are still bumpy, and public facilities are limited. Due to the narrow roads, we may encounter some traffic jams while moving from one spot to another.

Enjoy some leisure time at the beautiful Crystal Bay Beach! Relax on the sand and take a refreshing swim. Remember to bring your swimsuit. Paid changing rooms and showers are available for your convenience

Additionally, please bring some extra cash for restroom usage, shower facilities, and lunch. The local restaurants offer a variety of food options, including Indonesian, Western, and Chinese cuisine.

If you have any questions or need further assistance regarding this booking, please feel free to contact us.

Thank you,
Karma

Coco Express Pantai Matahari Terbit:https://maps.app.goo.gl/ifDWG2fmto3LEMCF8 
      `.trim();
      }
            
      
      else{
        message = `
${row.Group},
  
Dear Mr./Mrs ${namaTamu}
  
Greetings from Trip Gotik Get Your Guide Local Partner. We are excited to inform you that your booking for the Nusa Penida Trip with the following details is confirmed:
  
* Booking Code  : ${bookingCode}
* Activity Date : ${activityDate}
* Total Person  : ${pax} ${paxLabel}
${polaroidData ? `* Add on       : ${polaroidData}\n` : ''}
Please note that you must arrive in THE ANGKAL FAST BOAT office, which is located directly next to CK Mart Matahari Terbit (address provided in the link below), at Sanur Matahari Terbit Harbor at 7:00 AM for the 7:30 AM boat departure. Should you encounter any difficulties with the timing or have trouble locating the office, kindly contact this number for immediate assistance.

During the tour, you will be accompanied by a tour leader. When you arrive in Nusa Penida, please be attentive and look for our team holding a white paper sign with your name on it. Your tour will be arranged by our team from this point onwards.

To ensure your comfort throughout the trip, it is recommended that you wear comfortable clothing, Swimming suit, towel, walking shoes or flip-flops, apply sunscreen, and bring sunglasses. Nusa Penida is a relatively new destination that is not fully developed yet, giving you a glimpse of Bali as it was 30 years ago. Approximately 20% of the roads in Nusa Penida are still bumpy, and public facilities are limited. Due to the narrow roads, we may encounter some traffic jams while moving from one spot to another.

Enjoy some leisure time at the beautiful Crystal Bay Beach! Relax on the sand and take a refreshing swim. Remember to bring your swimsuit. Paid changing rooms and showers are available for your convenience

Additionally, please bring some extra cash for restroom usage, shower facilities, and lunch. The local restaurants offer a variety of food options, including Indonesian, Western, and Chinese cuisine.

If you have any questions or need further assistance regarding this booking, please feel free to contact us.

Thank you,
Karma

CK Mart Matahari Terbit: https://maps.app.goo.gl/W4Y8V1NBk354mSi16 
      `.trim();
    }
  }
      
      

    else if (isCategoryET) {
      
  
      message = `
${row.Group},
  
Dear Mr./Mrs ${namaTamu}
  
Greetings from Trip Gotik Get Your Guide Local Partner. We are excited to inform you that your booking for the Nusa Penida Trip with the following details is confirmed:
  
* Booking Code  : ${bookingCode}
* Activity Date : ${activityDate}
* Total Person  : ${pax} ${paxLabel}
${polaroidData ? `* Add on       : ${polaroidData}\n` : ''}
Please note that your pick-up time will be between ${row.PickupTime} - ${pickupTimeUpdated} AM from ${row.Location}. Upon arrival at your hotel, our driver will contact you. The driver will assist you with the check in process in Bali harbor. We are scheduled depart at 07:30 AM from Sanur port. 

When you arrive in Nusa Penida, please be attentive and look for our team holding a white paper sign with your name on it. Your tour will be arranged by our team from this point onwards.

To ensure your comfort throughout the trip, it is recommended that you wear comfortable clothing, swimsuit, towel, walking shoes or trekking shoes, apply sunscreen, Towels and bring sunglasses. Nusa Penida is a relatively new destination that is not fully developed yet, giving you a glimpse of Bali as it was 30 years ago. Approximately 20% of the roads in Nusa Penida are still bumpy, and public facilities are limited. Due to the narrow roads, we may encounter some traffic jams while moving from one spot to another.

Enjoy some leisure time at the beautiful Diamond and Atuh Beach! Relax on the sand and take a refreshing swim. Remember to bring your swimsuit. Paid changing rooms and showers are available for your convenience

Additionally, please bring some extra cash for restroom usage, shower and lunch. The local restaurants offer a variety of food options, including Indonesian, Western, and Chinese cuisine.Upon your return to Sanur Harbor around 5:45- 6:00 PM, please make your way back to the ticket pick-up point. Your driver will be waiting there, ready to transport you back to your hotel.

If you have any questions or need further assistance regarding this booking, please feel free to contact us.

Thank you
Karma
      `.trim();
    }
    else if (isCategorySW) {
      
  
      message = `
${row.Group},
  
Dear Mr./Mrs ${namaTamu}
  
Greetings from Trip Gotik Get Your Guide Local Partner. We are excited to inform you that your booking for the Nusa Penida Trip with the following details is confirmed:
  
* Booking Code  : ${bookingCode}
* Activity Date : ${activityDate}
* Total Person  : ${pax} ${paxLabel}
${polaroidData ? `* Add on       : ${polaroidData}\n` : ''}
Please note that your pick-up time will be between ${row.PickupTime} - ${pickupTimeUpdated} AM from ${row.Location}. The driver will assist you with the check in process in Bali harbor. Please be informed that this is a group tour, and on rare occasions, some participants may not be punctual. However, rest assured that we will inform you in case of any delays when picking you up. Please don't worry, as you will still be picked up as scheduled

For tomorrow we are scheduled depart at 07:30 AM from Sanur port. When you arrive in Nusa Penida, please be attentive and look for our team holding a white paper sign with your name on it. Your tour will be arranged by our team from this point onwards.

To ensure your comfort throughout the trip, it is recommended that you wear comfortable clothing, walking shoes, sneakers, apply sunscreen, and bring sunglasses. 

Start your journey with snorkeling in 3 spots: Manta Point, Gamat Point and Crystal Bay Point. Finish with snorkel, enjoy the island tour to visit Kelingking Beach, Broken Beach, and Angel Billabong Beach. 

Nusa Penida is a relatively new destination that is not fully developed yet, giving you a glimpse of Bali as it was 30 years ago. Approximately 20% of the roads in Nusa Penida are still bumpy, and public facilities are limited. Due to the narrow roads, we may encounter some traffic jams while moving from one spot to another.

Additionally, please bring some extra cash for restroom usage and lunch. The local restaurants offer a variety of food options, including Indonesian, Western, and Chinese cuisine.

Upon your return to Sanur Harbor around 5:45- 6:00 PM, please make your way back to the ticket pick-up point. Your driver will be waiting there, ready to transport you back to your hotel.  

If you have any questions or need further assistance regarding this booking, please feel free to contact us.

Thank you,
Karma
      `.trim();
    }
    else if (isCategoryIBH) {
      
      if (row.Location === "Sanur Ferry Port Meeting Point") {
        message = `
${row.Group},
  
Dear Mr./Mrs ${namaTamu}
  
Greetings from Trip Gotik Get Your Guide Local Partner. We are excited to inform you that your booking for the Nusa Penida Trip with the following details is confirmed:
  
* Booking Code  : ${bookingCode}
* Activity Date : ${activityDate}
* Total Person  : ${pax} ${paxLabel}
${polaroidData ? `* Add on       : ${polaroidData}\n` : ''}
Please note that you must arrive at Sanur Matahari Terbit Harbor at 7:00 AM for 7:30 AM boat departure. Our leader, Mr. Galung, can be reached at +62 813-5312-3400 and will assist you with the check-in process. If you have any trouble communicating with our leader, please head to the WIJAYA BUYUK FAST BOAT COUNTER located next to COCO MART EXPRESS for further assistance.

During the tour, you will be accompanied by a tour leader. When you arrive in Nusa Penida, please be attentive and look for our team holding a white paper sign with your name on it. Your tour will be arranged by our team from this point onwards.

To ensure your comfort throughout the trip, it is recommended that you wear comfortable clothing, walking shoes, apply sunscreen, and bring sunglasses. As part of your itinerary, you will be visiting Diamond Beach, Atuh Beach, and Kelingking Beach, which involve full trekking. Please be prepared for a physically demanding experience, as you will be descending and ascending over 100 meters with slopes ranging from 30 to 45 degrees. It is essential to wear good-quality trekking shoes and be in optimal physical condition.

Nusa Penida is a relatively new destination that is not fully developed yet, giving you a glimpse of Bali as it was 30 years ago. Approximately 20% of the roads in Nusa Penida are still bumpy, and public facilities are limited. Due to the narrow roads, we may encounter some traffic jams while moving from one spot to another.

If you would like to spend extra time swimming at a beach you will be visiting tomorrow, please discuss this with the driver to adjust the schedule before reaching Nusa Penida Port. If time allows, the driver will be happy to accommodate your request. Please bring your swimsuit and towel, as showers and changing rooms are available at an additional cost.

Additionally, please bring some extra cash for restroom usage, shower facilities, and lunch. The local restaurants offer a variety of food options, including Indonesian, Western, and Chinese cuisine.

If you have any questions or need further assistance regarding this booking, please feel free to contact us.

Thank you
Karma

Coco Express Pantai Matahari Terbit:https://maps.app.goo.gl/ifDWG2fmto3LEMCF8 
      `.trim();
      }
      else{
        message = `
${row.Group},
  
Dear Mr./Mrs ${namaTamu}
  
Greetings from Trip Gotik Get Your Guide Local Partner. We are excited to inform you that your booking for the Nusa Penida Trip with the following details is confirmed:
  
* Booking Code  : ${bookingCode}
* Activity Date : ${activityDate}
* Total Person  : ${pax} ${paxLabel}
${polaroidData ? `* Add on       : ${polaroidData}\n` : ''}
Please note that your pick-up time will be between ${row.PickupTime} - ${pickupTimeUpdated} AM from ${row.Location}. Upon arrival at your hotel, our driver will contact you. The driver will assist you with the check-in process at Bali harbor.

We apologize for the slight adjustment in the pick-up time due to the scheduled departure time from Sanur Port is at 7:30 AM so that we can arrive in Nusa Penida earlier and cover all the destinations as planned. Please be informed that this is a group tour, and on rare occasions, some participants may not be punctual. However, rest assured that we will inform you in case of any delays when picking you up, and you will still be picked up as scheduled.

To ensure your comfort throughout the trip, it is recommended that you wear comfortable clothing, walking shoes, apply sunscreen, and bring sunglasses. As part of your itinerary, you will be visiting Diamond Beach, Atuh Beach, and Kelingking Beach, which involve full trekking. Please be prepared for a physically demanding experience, as you will be descending and ascending over 100 meters with slopes ranging from 30 to 45 degrees. It is essential to wear good-quality trekking shoes and be in optimal physical condition.

Nusa Penida is a relatively new destination that is not fully developed yet, giving you a glimpse of Bali as it was 30 years ago. Approximately 20% of the roads in Nusa Penida are still bumpy, and public facilities are limited. Due to the narrow roads, we may encounter some traffic jams while moving from one spot to another.

If you would like to spend extra time swimming at a beach you will be visiting tomorrow, please discuss this with the driver to adjust the schedule before reaching Nusa Penida Port. If time allows, the driver will be happy to accommodate your request. Please bring your swimsuit and towel, as showers and changing rooms are available at an additional cost.

Additionally, please bring some extra cash for restroom usage, shower facilities, and lunch. The local restaurants offer a variety of food options, including Indonesian, Western, and Chinese cuisine.

If you have any questions or need further assistance regarding this booking, please feel free to contact us.

Thank you
Karma
      `.trim();
      }
    }
    else if (isCategorySMP) {
      
  
      message = `
${row.Group},
  
Dear Mr./Mrs ${namaTamu}
  
Greetings from Trip Gotik Get Your Guide Local Partner. We are excited to inform you that your booking for the Nusa Penida Trip with the following details is confirmed:
  
* Booking Code  : ${bookingCode}
* Activity Date : ${activityDate}
* Total Person  : ${pax} ${paxLabel}
${polaroidData ? `* Add on       : ${polaroidData}\n` : ''}
Please note that you must arrive at Sanur Matahari Terbit Harbor at 7:00 AM for the 7:30 AM boat departure. please proceed to the THE ANGKAL FAST BOAT office, which is located directly next to CK Mart (address provided in the link below), . Should you encounter any difficulties with the timing or have trouble locating the office, kindly contact this number for immediate assistance.

When you arrive in Nusa Penida, please be attentive and look for our team holding a white paper sign with your name on it. Your tour will be arranged by our team from this point onwards.

To ensure your comfort throughout the trip, it is recommended that you wear comfortable clothing, walking shoes, sneakers, apply sunscreen, and bring sunglasses.

Start your journey with snorkeling in 3 spots: Manta Point, Gamat Point and Crystal Bay Point. Finish with snorkel, enjoy the island tour to visit Kelingking Beach, Broken Beach, and Angel Billabong Beach.

Nusa Penida is a relatively new destination that is not fully developed yet, giving you a glimpse of Bali as it was 30 years ago. Approximately 20% of the roads in Nusa Penida are still bumpy, and public facilities are limited. Due to the narrow roads, we may encounter some traffic jams while moving from one spot to another.

Additionally, please bring some extra cash for restroom usage and lunch. The local restaurants offer a variety of food options, including Indonesian, Western, and Chinese cuisine.

Upon your return to Sanur Harbor around 5:45- 6:00 PM, please make your way back to the ticket pick-up point. Your driver will be waiting there, ready to transport you back to your hotel.

If you have any questions or need further assistance regarding this booking, please feel free to contact us.

Thank you,
Karma

CK Mart Matahari Terbit: https://maps.app.goo.gl/W4Y8V1NBk354mSi16 
      `.trim();
    }
    
    else {
      message = `Halo ${row.NamaTamu}, saya ingin menghubungi Anda melalui informasi dari file Excel.`.trim();
    }
    
    // Salin pesan ke clipboard
    navigator.clipboard.writeText(message)
    .then(() => {
        console.log('Pesan berhasil disalin ke clipboard');

        // Kirim pesan ke WhatsApp berdasarkan pilihan user
        if (preference === 'web') {
          window.open(`https://web.whatsapp.com/send?phone=${row.Phone}`, '_blank'); // Buka di tab baru
      } else {
          window.location.href = `whatsapp://send?phone=${row.Phone}`; // Buka di aplikasi WhatsApp
      }
    })
    .catch(err => {
      console.error('Gagal menyalin pesan ke clipboard: ', err);
    });

  }

  sendEmail(row: any): void {
    const group = row.Group.trim();
    const isCategoryE = group.startsWith('E.')|| group.startsWith('PE.');
    const isCategoryEMP = group.startsWith('EMP.');
    const isCategoryET = group.startsWith('ET.');
    const isCategoryAorNA = group.startsWith('A.') || group.startsWith('NA.');
    const isCategoryWMP = group.startsWith('WMP.');
    const isCategoryW = group.startsWith('W.') || group.startsWith('NW.'); // Periksa apakah kategori dimulai dengan "E"
    const isCategorySW = group.startsWith('SW.');
    const isCategoryIBH = group.startsWith('IBH.');
    const isCategorySMP = group.startsWith('SMP.');
    let message = '';
    const bookingCode = row.BookingCode || 'Unknown Booking Code'; // Ambil Booking Code dari data
      const activityDate = this.getNextDayDate(); // Dapatkan tanggal besok
      const pax = parseInt(row.Pax, 10) || 0; // Konversi Pax ke angka, default 0 jika tidak valid
      const paxLabel = pax === 1 ? 'person' : 'people';
      const namaTamu = row.NamaTamu.replace(/\s+/g, ' ').trim().replace(/\s*\(.*?\)\s*/g, '');
      const pickupTimeStr = String(row.PickupTime); // Contoh: "6.00"
      let [hours, minutes] = pickupTimeStr.split('.').map(Number); // Pisahkan jam dan menit

      minutes += 10; // Tambah 10 menit

      if (minutes >= 60) { 
          hours += Math.floor(minutes / 60); // Tambahkan ke jam jika lebih dari 60 menit
          minutes %= 60; // Sisakan menit yang tersisa
      }

      // Format hasil agar menit tetap dua digit
      const pickupTimeUpdated = `${hours}.${minutes.toString().padStart(2, '0')}`;
      const additionalInfo = row.AdditionalInfo || '';
      const polaroidMatch = additionalInfo.match(/Polaroid.*?\)/); // Ambil teks "Polaroid" sampai tanda tutup kurung
      const polaroidData = polaroidMatch ? polaroidMatch[0] : null;
  
      if (isCategoryE) {
      
  
        message = `
  ${row.Group},
    
  Dear Mr./Mrs ${namaTamu}
    
  Greetings from Trip Gotik Get Your Guide Local Partner. We are excited to inform you that your booking for the Nusa Penida Trip with the following details is confirmed:
    
  * Booking Code  : ${bookingCode}
  * Activity Date : ${activityDate}
  * Total Person  : ${pax} ${paxLabel}
  ${polaroidData ? `* Add on       : ${polaroidData}\n` : ''}
  Please note that your pick-up time will be between ${row.PickupTime} - ${pickupTimeUpdated} AM from ${row.Location}. The driver will assist you with the check in process in Bali harbor. Please be informed that this is a group tour, and on rare occasions, some participants may not be punctual. However, rest assured that we will inform you in case of any delays when picking you up. Please don't worry, as you will still be picked up as scheduled
  
  For tomorrow we are scheduled depart at 07:30 AM from Sanur port. When you arrive in Nusa Penida, please be attentive and look for our team holding a white paper sign with your name on it. Your tour will be arranged by our team from this point onwards.
  
  To ensure your comfort throughout the trip, it is recommended that you wear comfortable clothing, walking shoes, sneakers, apply sunscreen, and bring sunglasses.  Feel free to bring your swimsuit. You will have the chance to go for a swim, especially when visiting Diamond Beach. 
  
  Nusa Penida is a relatively new destination that is not fully developed yet, giving you a glimpse of Bali as it was 30 years ago. Approximately 20% of the roads in Nusa Penida are still bumpy, and public facilities are limited. Due to the narrow roads, we may encounter some traffic jams while moving from one spot to another.
  
  Additionally, please bring some extra cash for restroom usage and lunch. The local restaurants offer a variety of food options, including Indonesian, Western, and Chinese cuisine.
  
  Upon your return to Sanur Harbor around 5:45- 6:00 PM, please make your way back to the ticket pick-up point. Your driver will be waiting there, ready to transport you back to your hotel.  
  
  If you have any questions or need further assistance regarding this booking, please feel free to contact us.
  
  Kindly replay this email via Whatsapp for effective communication +6287722748143 

  Thank you,
  Karma
        `.trim();
      } 
  
      else if (isCategoryEMP) {
     
  
        if (row.AdditionalInfo.toLowerCase() === "nusa team") {
          message = `
  ${row.Group},
            
  Dear Mr./Mrs ${namaTamu}
            
  Greetings from Trip Gotik Get Your Guide Local Partner. We are excited to inform you that your booking for the Nusa Penida Trip with the following details is confirmed:
            
  * Booking Code  : ${bookingCode}
  * Activity Date : ${activityDate}
  * Total Person  : ${pax} ${paxLabel}
  ${polaroidData ? `* Add on       : ${polaroidData}\n` : ''}
  Please note that you must arrive at Sanur Matahari Terbit Harbor at 8:00 AM for 8:30 AM boat departure. Our leader, Mr. Galung, can be reached at +62 813-5312-3400 and will assist you with the check-in process. If you have any trouble communicating with our leader, please head to the WIJAYA BUYUK FAST BOAT COUNTER located next to COCO MART EXPRESS for further assistance.
  
  During the tour, you will be accompanied by a tour leader. When you arrive in Nusa Penida, please be attentive and look for our team holding a white paper sign with your name on it. Your tour will be arranged by our team from this point onwards.
  
  To ensure your comfort throughout the trip, it is recommended that you wear comfortable clothing, Swimming suit, towel, walking shoes or flip-flops, apply sunscreen, and bring sunglasses. Nusa Penida is a relatively new destination that is not fully developed yet, giving you a glimpse of Bali as it was 30 years ago. Approximately 20% of the roads in Nusa Penida are still bumpy, and public facilities are limited. Due to the narrow roads, we may encounter some traffic jams while moving from one spot to another.
  
  Enjoy some leisure time at the beautiful Crystal Bay Beach! Relax on the sand and take a refreshing swim. Remember to bring your swimsuit. Paid changing rooms and showers are available for your convenience
  
  Additionally, please bring some extra cash for restroom usage, shower facilities, and lunch. The local restaurants offer a variety of food options, including Indonesian, Western, and Chinese cuisine.
  
  If you have any questions or need further assistance regarding this booking, please feel free to contact us.

  Kindly replay this email via Whatsapp for effective communication +6287722748143 
  
  Thank you,
  Karma
  
  Coco Express Pantai Matahari Terbit:https://maps.app.goo.gl/ifDWG2fmto3LEMCF8 
        `.trim();
        }
              
        
        else{
          message = `
  ${row.Group},
    
  Dear Mr./Mrs ${namaTamu}
    
  Greetings from Trip Gotik Get Your Guide Local Partner. We are excited to inform you that your booking for the Nusa Penida Trip with the following details is confirmed:
    
  * Booking Code  : ${bookingCode}
  * Activity Date : ${activityDate}
  * Total Person  : ${pax} ${paxLabel}
  ${polaroidData ? `* Add on       : ${polaroidData}\n` : ''}
  Please note that you must arrive in THE ANGKAL FAST BOAT office, which is located directly next to CK Mart Matahari Terbit (address provided in the link below), at Sanur Matahari Terbit Harbor at 7:00 AM for the 7:30 AM boat departure. Should you encounter any difficulties with the timing or have trouble locating the office, kindly contact this number for immediate assistance.
  
  During the tour, you will be accompanied by a tour leader. When you arrive in Nusa Penida, please be attentive and look for our team holding a white paper sign with your name on it. Your tour will be arranged by our team from this point onwards.
  
  To ensure your comfort throughout the trip, it is recommended that you wear comfortable clothing, Swimming suit, towel, walking shoes or flip-flops, apply sunscreen, and bring sunglasses. Nusa Penida is a relatively new destination that is not fully developed yet, giving you a glimpse of Bali as it was 30 years ago. Approximately 20% of the roads in Nusa Penida are still bumpy, and public facilities are limited. Due to the narrow roads, we may encounter some traffic jams while moving from one spot to another.
  
  Enjoy some leisure time at the beautiful Crystal Bay Beach! Relax on the sand and take a refreshing swim. Remember to bring your swimsuit. Paid changing rooms and showers are available for your convenience
  
  Additionally, please bring some extra cash for restroom usage, shower facilities, and lunch. The local restaurants offer a variety of food options, including Indonesian, Western, and Chinese cuisine.
  
  If you have any questions or need further assistance regarding this booking, please feel free to contact us.

  Kindly replay this email via Whatsapp for effective communication +6287722748143 
  
  Thank you,
  Karma
  
  CK Mart Matahari Terbit: https://maps.app.goo.gl/W4Y8V1NBk354mSi16 
        `.trim();
      }
    }
  
      else if (isCategoryAorNA) {
        
    
        message = `
  ${row.Group},
    
  Dear Mr./Mrs ${namaTamu}
    
  Greetings from Trip Gotik Get Your Guide Local Partner. We are excited to inform you that your booking for the Nusa Penida Trip with the following details is confirmed:
    
  * Booking Code  : ${bookingCode}
  * Activity Date : ${activityDate}
  * Total Person  : ${pax} ${paxLabel}
  ${polaroidData ? `* Add on       : ${polaroidData}\n` : ''}
  Please note that your pick-up time will be between ${row.PickupTime} - ${pickupTimeUpdated} AM at ${row.Location}. Upon arrival at your hotel, our driver will contact you. The driver will assist you with the check in process in Bali harbor. 
  
  Sorry for this inconvenience that we have to pick you up slightly earlier than usual, cause of pickup distance from the harbor and for tomorrow we are scheduled depart at 07:00 AM from Sanur port, so that we can arrive in Nusa Penida earlier and cover all the destinations as planned specially for taking photo at tree house. Please be informed, this is a group tour and on rare occasions, some participants may not be punctual. However, rest assured that we will inform you in case of any delays when picking you up. Please don't worry, as you will still be picked up as scheduled. 
  
  When you arrive in Nusa Penida, please be attentive and look for our team holding a white paper sign with your name on it. Your tour will be arranged by our team from this point onwards. 
  
  If you would like to spend extra time swimming at a beach, you will have that time when you visit Diamond Beach. To ensure your comfort throughout the trip, it is recommended that you wear comfortable clothing, walking shoes, apply sunscreen, and bring sunglasses. 
  
  Nusa Penida is a relatively new destination that is not fully developed yet, giving you a glimpse of Bali as it was 30 years ago. Approximately 20% of the roads in Nusa Penida are still bumpy, and public facilities are limited. Due to the narrow roads, we may encounter some traffic jams while moving from one spot to another.
  
  Additionally, please bring some extra cash for restroom usage and lunch. The local restaurants offer a variety of food options, including Indonesian, Western, and Chinese cuisine.
  
  Upon your return to Sanur Harbor around 5:30 PM, please make your way back to the ticket pick-up point. Your driver will be waiting there, ready to transport you back to your hotel.
  
  If you have any questions or need further assistance regarding this booking, please feel free to contact us.

  Kindly replay this email via Whatsapp for effective communication +6287722748143 
  
  Thank you,
  Karma 
        `.trim();
      }
  
      else if (isCategoryW) {
      
  
        message = `
  ${row.Group},
    
  Dear Mr./Mrs ${namaTamu}
    
  Greetings from Trip Gotik Get Your Guide Local Partner. We are excited to inform you that your booking for the Nusa Penida Trip with the following details is confirmed:
    
  * Booking Code  : ${bookingCode}
  * Activity Date : ${activityDate}
  * Total Person  : ${pax} ${paxLabel}
  ${polaroidData ? `* Add on       : ${polaroidData}\n` : ''}
  Please note that your pick-up time will be between ${row.PickupTime} - ${pickupTimeUpdated} AM from ${row.Location}, Jalan Hard Rock Hotel Kabupaten Badung. Upon arrival at your hotel, our driver will contact you. The driver will assist you with the check-in process at Bali harbor.
  
  We apologize for the slight adjustment in the pick-up time due to  The scheduled departure time from Sanur Port is at 7:30 AM so that we can arrive in Nusa Penida earlier and cover all the destinations as planned. Please be informed that this is a group tour, and on rare occasions, some participants may not be punctual. However, rest assured that we will inform you in case of any delays when picking you up, and you will still be picked up as scheduled.
  
  When you arrive in Nusa Penida, please be attentive and look for our team holding a white paper sign with your name on it. Your tour will be arranged by our team from this point onwards.
  
  To ensure your comfort throughout the trip, it is recommended that you wear comfortable clothing, a swimsuit, towel, walking or trekking shoes, apply sunscreen, and bring sunglasses. Nusa Penida is a relatively new destination that is not fully developed yet, giving you a glimpse of Bali as it was 30 years ago. Approximately 20% of the roads in Nusa Penida are still bumpy, and public facilities are limited. Due to the narrow roads, we may encounter some traffic jams while moving from one spot to another.
  
  Enjoy some leisure time at the beautiful Crystal Bay Beach! Relax on the sand and take a refreshing swim. Remember to bring your swimsuit. Paid changing rooms and showers are available for your convenience.
  
  Additionally, please bring some extra cash for restroom usage, showers, and lunch. The local restaurants offer a variety of food options, including Indonesian, Western, and Chinese cuisine.
  
  For your return journey, the boat is scheduled to depart from Nusa Penida at 4:00 PM/ 4:30 PM, and you will arrive back in Bali around 1 hour after. Upon your return to Sanur Harbor, please make your way back to the ticket pick-up point. Your driver will be waiting there, ready to transport you back to your hotel.
  
  If you have any questions or need further assistance regarding this booking, please feel free to contact us.

  Kindly replay this email via Whatsapp for effective communication +6287722748143 
  
  Thank you,
  Karma
        `.trim();
      }
  
      else if (isCategoryWMP) {
      
        if (row.AdditionalInfo.toLowerCase() === "nusa team") {
          message = `
  ${row.Group},
            
  Dear Mr./Mrs ${namaTamu}
            
  Greetings from Trip Gotik Get Your Guide Local Partner. We are excited to inform you that your booking for the Nusa Penida Trip with the following details is confirmed:
            
  * Booking Code  : ${bookingCode}
  * Activity Date : ${activityDate}
  * Total Person  : ${pax} ${paxLabel}
  ${polaroidData ? `* Add on       : ${polaroidData}\n` : ''}
  Please note that you must arrive at Sanur Matahari Terbit Harbor at 8:00 AM for 8:30 AM boat departure. Our leader, Mr. Galung, can be reached at +62 813-5312-3400 and will assist you with the check-in process. If you have any trouble communicating with our leader, please head to the WIJAYA BUYUK FAST BOAT COUNTER located next to COCO MART EXPRESS for further assistance.
  
  During the tour, you will be accompanied by a tour leader. When you arrive in Nusa Penida, please be attentive and look for our team holding a white paper sign with your name on it. Your tour will be arranged by our team from this point onwards.
  
  To ensure your comfort throughout the trip, it is recommended that you wear comfortable clothing, Swimming suit, towel, walking shoes or flip-flops, apply sunscreen, and bring sunglasses. Nusa Penida is a relatively new destination that is not fully developed yet, giving you a glimpse of Bali as it was 30 years ago. Approximately 20% of the roads in Nusa Penida are still bumpy, and public facilities are limited. Due to the narrow roads, we may encounter some traffic jams while moving from one spot to another.
  
  Enjoy some leisure time at the beautiful Crystal Bay Beach! Relax on the sand and take a refreshing swim. Remember to bring your swimsuit. Paid changing rooms and showers are available for your convenience
  
  Additionally, please bring some extra cash for restroom usage, shower facilities, and lunch. The local restaurants offer a variety of food options, including Indonesian, Western, and Chinese cuisine.
  
  If you have any questions or need further assistance regarding this booking, please feel free to contact us.

  Kindly replay this email via Whatsapp for effective communication +6287722748143 
  
  Thank you,
  Karma
  
  Coco Express Pantai Matahari Terbit:https://maps.app.goo.gl/ifDWG2fmto3LEMCF8 
        `.trim();
        }
              
        
        else{
          message = `
  ${row.Group},
    
  Dear Mr./Mrs ${namaTamu}
    
  Greetings from Trip Gotik Get Your Guide Local Partner. We are excited to inform you that your booking for the Nusa Penida Trip with the following details is confirmed:
    
  * Booking Code  : ${bookingCode}
  * Activity Date : ${activityDate}
  * Total Person  : ${pax} ${paxLabel}
  ${polaroidData ? `* Add on       : ${polaroidData}\n` : ''}
  Please note that you must arrive in THE ANGKAL FAST BOAT office, which is located directly next to CK Mart Matahari Terbit (address provided in the link below), at Sanur Matahari Terbit Harbor at 7:00 AM for the 7:30 AM boat departure. Should you encounter any difficulties with the timing or have trouble locating the office, kindly contact this number for immediate assistance.
  
  During the tour, you will be accompanied by a tour leader. When you arrive in Nusa Penida, please be attentive and look for our team holding a white paper sign with your name on it. Your tour will be arranged by our team from this point onwards.
  
  To ensure your comfort throughout the trip, it is recommended that you wear comfortable clothing, Swimming suit, towel, walking shoes or flip-flops, apply sunscreen, and bring sunglasses. Nusa Penida is a relatively new destination that is not fully developed yet, giving you a glimpse of Bali as it was 30 years ago. Approximately 20% of the roads in Nusa Penida are still bumpy, and public facilities are limited. Due to the narrow roads, we may encounter some traffic jams while moving from one spot to another.
  
  Enjoy some leisure time at the beautiful Crystal Bay Beach! Relax on the sand and take a refreshing swim. Remember to bring your swimsuit. Paid changing rooms and showers are available for your convenience
  
  Additionally, please bring some extra cash for restroom usage, shower facilities, and lunch. The local restaurants offer a variety of food options, including Indonesian, Western, and Chinese cuisine.
  
  If you have any questions or need further assistance regarding this booking, please feel free to contact us.

  Kindly replay this email via Whatsapp for effective communication +6287722748143 
  
  Thank you,
  Karma
  
  CK Mart Matahari Terbit: https://maps.app.goo.gl/W4Y8V1NBk354mSi16 
        `.trim();
      }
    }
  
      else if (isCategoryET) {
        
    
        message = `
  ${row.Group},
    
  Dear Mr./Mrs ${namaTamu}
    
  Greetings from Trip Gotik Get Your Guide Local Partner. We are excited to inform you that your booking for the Nusa Penida Trip with the following details is confirmed:
    
  * Booking Code  : ${bookingCode}
  * Activity Date : ${activityDate}
  * Total Person  : ${pax} ${paxLabel}
  ${polaroidData ? `* Add on       : ${polaroidData}\n` : ''}
  Please note that your pick-up time will be between ${row.PickupTime} - ${pickupTimeUpdated} AM from ${row.Location}. Upon arrival at your hotel, our driver will contact you. The driver will assist you with the check in process in Bali harbor. We are scheduled depart at 07:30 AM from Sanur port. 
  
  When you arrive in Nusa Penida, please be attentive and look for our team holding a white paper sign with your name on it. Your tour will be arranged by our team from this point onwards.
  
  To ensure your comfort throughout the trip, it is recommended that you wear comfortable clothing, swimsuit, towel, walking shoes or trekking shoes, apply sunscreen, Towels and bring sunglasses. Nusa Penida is a relatively new destination that is not fully developed yet, giving you a glimpse of Bali as it was 30 years ago. Approximately 20% of the roads in Nusa Penida are still bumpy, and public facilities are limited. Due to the narrow roads, we may encounter some traffic jams while moving from one spot to another.
  
  Enjoy some leisure time at the beautiful Diamond and Atuh Beach! Relax on the sand and take a refreshing swim. Remember to bring your swimsuit. Paid changing rooms and showers are available for your convenience
  
  Additionally, please bring some extra cash for restroom usage, shower and lunch. The local restaurants offer a variety of food options, including Indonesian, Western, and Chinese cuisine.Upon your return to Sanur Harbor around 5:45- 6:00 PM, please make your way back to the ticket pick-up point. Your driver will be waiting there, ready to transport you back to your hotel.
  
  If you have any questions or need further assistance regarding this booking, please feel free to contact us.

  Kindly replay this email via Whatsapp for effective communication +6287722748143 
  
  Thank you
  Karma
        `.trim();
      }
      else if (isCategorySW) {
        
    
        message = `
  ${row.Group},
    
  Dear Mr./Mrs ${namaTamu}
    
  Greetings from Trip Gotik Get Your Guide Local Partner. We are excited to inform you that your booking for the Nusa Penida Trip with the following details is confirmed:
    
  * Booking Code  : ${bookingCode}
  * Activity Date : ${activityDate}
  * Total Person  : ${pax} ${paxLabel}
  ${polaroidData ? `* Add on       : ${polaroidData}\n` : ''}
  Please note that your pick-up time will be between ${row.PickupTime} - ${pickupTimeUpdated} AM from ${row.Location}. The driver will assist you with the check in process in Bali harbor. Please be informed that this is a group tour, and on rare occasions, some participants may not be punctual. However, rest assured that we will inform you in case of any delays when picking you up. Please don't worry, as you will still be picked up as scheduled
  
  For tomorrow we are scheduled depart at 07:30 AM from Sanur port. When you arrive in Nusa Penida, please be attentive and look for our team holding a white paper sign with your name on it. Your tour will be arranged by our team from this point onwards.
  
  To ensure your comfort throughout the trip, it is recommended that you wear comfortable clothing, walking shoes, sneakers, apply sunscreen, and bring sunglasses. 
  
  Start your journey with snorkeling in 3 spots: Manta Point, Gamat Point and Crystal Bay Point. Finish with snorkel, enjoy the island tour to visit Kelingking Beach, Broken Beach, and Angel Billabong Beach. 
  
  Nusa Penida is a relatively new destination that is not fully developed yet, giving you a glimpse of Bali as it was 30 years ago. Approximately 20% of the roads in Nusa Penida are still bumpy, and public facilities are limited. Due to the narrow roads, we may encounter some traffic jams while moving from one spot to another.
  
  Additionally, please bring some extra cash for restroom usage and lunch. The local restaurants offer a variety of food options, including Indonesian, Western, and Chinese cuisine.
  
  Upon your return to Sanur Harbor around 5:45- 6:00 PM, please make your way back to the ticket pick-up point. Your driver will be waiting there, ready to transport you back to your hotel.  
  
  If you have any questions or need further assistance regarding this booking, please feel free to contact us.

  Kindly replay this email via Whatsapp for effective communication +6287722748143 
  
  Thank you,
  Karma
        `.trim();
      }
      else if (isCategoryIBH) {
        
        if (row.Location === "Sanur Ferry Port Meeting Point") {
          message = `
  ${row.Group},
    
  Dear Mr./Mrs ${namaTamu}
    
  Greetings from Trip Gotik Get Your Guide Local Partner. We are excited to inform you that your booking for the Nusa Penida Trip with the following details is confirmed:
    
  * Booking Code  : ${bookingCode}
  * Activity Date : ${activityDate}
  * Total Person  : ${pax} ${paxLabel}
  ${polaroidData ? `* Add on       : ${polaroidData}\n` : ''}
  Please note that you must arrive at Sanur Matahari Terbit Harbor at 7:00 AM for 7:30 AM boat departure. Our leader, Mr. Galung, can be reached at +62 813-5312-3400 and will assist you with the check-in process. If you have any trouble communicating with our leader, please head to the WIJAYA BUYUK FAST BOAT COUNTER located next to COCO MART EXPRESS for further assistance.
  
  During the tour, you will be accompanied by a tour leader. When you arrive in Nusa Penida, please be attentive and look for our team holding a white paper sign with your name on it. Your tour will be arranged by our team from this point onwards.
  
  To ensure your comfort throughout the trip, it is recommended that you wear comfortable clothing, walking shoes, apply sunscreen, and bring sunglasses. As part of your itinerary, you will be visiting Diamond Beach, Atuh Beach, and Kelingking Beach, which involve full trekking. Please be prepared for a physically demanding experience, as you will be descending and ascending over 100 meters with slopes ranging from 30 to 45 degrees. It is essential to wear good-quality trekking shoes and be in optimal physical condition.
  
  Nusa Penida is a relatively new destination that is not fully developed yet, giving you a glimpse of Bali as it was 30 years ago. Approximately 20% of the roads in Nusa Penida are still bumpy, and public facilities are limited. Due to the narrow roads, we may encounter some traffic jams while moving from one spot to another.
  
  If you would like to spend extra time swimming at a beach you will be visiting tomorrow, please discuss this with the driver to adjust the schedule before reaching Nusa Penida Port. If time allows, the driver will be happy to accommodate your request. Please bring your swimsuit and towel, as showers and changing rooms are available at an additional cost.
  
  Additionally, please bring some extra cash for restroom usage, shower facilities, and lunch. The local restaurants offer a variety of food options, including Indonesian, Western, and Chinese cuisine.
  
  If you have any questions or need further assistance regarding this booking, please feel free to contact us.

  Kindly replay this email via Whatsapp for effective communication +6287722748143 
  
  Thank you
  Karma
  
  Coco Express Pantai Matahari Terbit:https://maps.app.goo.gl/ifDWG2fmto3LEMCF8 
        `.trim();
        }
        else{
          message = `
  ${row.Group},
    
  Dear Mr./Mrs ${namaTamu}
    
  Greetings from Trip Gotik Get Your Guide Local Partner. We are excited to inform you that your booking for the Nusa Penida Trip with the following details is confirmed:
    
  * Booking Code  : ${bookingCode}
  * Activity Date : ${activityDate}
  * Total Person  : ${pax} ${paxLabel}
  ${polaroidData ? `* Add on       : ${polaroidData}\n` : ''}
  Please note that your pick-up time will be between ${row.PickupTime} - ${pickupTimeUpdated} AM from ${row.Location}. Upon arrival at your hotel, our driver will contact you. The driver will assist you with the check-in process at Bali harbor.
  
  We apologize for the slight adjustment in the pick-up time due to the scheduled departure time from Sanur Port is at 7:30 AM so that we can arrive in Nusa Penida earlier and cover all the destinations as planned. Please be informed that this is a group tour, and on rare occasions, some participants may not be punctual. However, rest assured that we will inform you in case of any delays when picking you up, and you will still be picked up as scheduled.
  
  To ensure your comfort throughout the trip, it is recommended that you wear comfortable clothing, walking shoes, apply sunscreen, and bring sunglasses. As part of your itinerary, you will be visiting Diamond Beach, Atuh Beach, and Kelingking Beach, which involve full trekking. Please be prepared for a physically demanding experience, as you will be descending and ascending over 100 meters with slopes ranging from 30 to 45 degrees. It is essential to wear good-quality trekking shoes and be in optimal physical condition.
  
  Nusa Penida is a relatively new destination that is not fully developed yet, giving you a glimpse of Bali as it was 30 years ago. Approximately 20% of the roads in Nusa Penida are still bumpy, and public facilities are limited. Due to the narrow roads, we may encounter some traffic jams while moving from one spot to another.
  
  If you would like to spend extra time swimming at a beach you will be visiting tomorrow, please discuss this with the driver to adjust the schedule before reaching Nusa Penida Port. If time allows, the driver will be happy to accommodate your request. Please bring your swimsuit and towel, as showers and changing rooms are available at an additional cost.
  
  Additionally, please bring some extra cash for restroom usage, shower facilities, and lunch. The local restaurants offer a variety of food options, including Indonesian, Western, and Chinese cuisine.
  
  If you have any questions or need further assistance regarding this booking, please feel free to contact us.

  Kindly replay this email via Whatsapp for effective communication +6287722748143 
  
  Thank you
  Karma
        `.trim();
        }
      }
      else if (isCategorySMP) {
        
    
        message = `
  ${row.Group},
    
  Dear Mr./Mrs ${namaTamu}
    
  Greetings from Trip Gotik Get Your Guide Local Partner. We are excited to inform you that your booking for the Nusa Penida Trip with the following details is confirmed:
    
  * Booking Code  : ${bookingCode}
  * Activity Date : ${activityDate}
  * Total Person  : ${pax} ${paxLabel}
  ${polaroidData ? `* Add on       : ${polaroidData}\n` : ''}
  Please note that you must arrive at Sanur Matahari Terbit Harbor at 7:00 AM for the 7:30 AM boat departure. please proceed to the THE ANGKAL FAST BOAT office, which is located directly next to CK Mart (address provided in the link below), . Should you encounter any difficulties with the timing or have trouble locating the office, kindly contact this number for immediate assistance.
  
  When you arrive in Nusa Penida, please be attentive and look for our team holding a white paper sign with your name on it. Your tour will be arranged by our team from this point onwards.
  
  To ensure your comfort throughout the trip, it is recommended that you wear comfortable clothing, walking shoes, sneakers, apply sunscreen, and bring sunglasses.
  
  Start your journey with snorkeling in 3 spots: Manta Point, Gamat Point and Crystal Bay Point. Finish with snorkel, enjoy the island tour to visit Kelingking Beach, Broken Beach, and Angel Billabong Beach.
  
  Nusa Penida is a relatively new destination that is not fully developed yet, giving you a glimpse of Bali as it was 30 years ago. Approximately 20% of the roads in Nusa Penida are still bumpy, and public facilities are limited. Due to the narrow roads, we may encounter some traffic jams while moving from one spot to another.
  
  Additionally, please bring some extra cash for restroom usage and lunch. The local restaurants offer a variety of food options, including Indonesian, Western, and Chinese cuisine.
  
  Upon your return to Sanur Harbor around 5:45- 6:00 PM, please make your way back to the ticket pick-up point. Your driver will be waiting there, ready to transport you back to your hotel.
  
  If you have any questions or need further assistance regarding this booking, please feel free to contact us.

  Kindly replay this email via Whatsapp for effective communication +6287722748143 
  
  Thank you,
  Karma
  
  CK Mart Matahari Terbit: https://maps.app.goo.gl/W4Y8V1NBk354mSi16 
        `.trim();
      }
      
      else {
        message = `Halo ${row.NamaTamu}, saya ingin menghubungi Anda melalui informasi dari file Excel.`.trim();
      }

    navigator.clipboard.writeText(message)
    .then(() => {
      console.log('Pesan berhasil disalin ke clipboard');
      // Buka Gmail dengan pesan
      const emailLink = `https://mail.google.com/mail/?view=cm&fs=1&to=${row.Email}&su=Judul Email&body=${encodeURIComponent(message)}`;
      window.open(emailLink, '_blank');
    })
    .catch(err => {
      console.error('Gagal menyalin pesan ke clipboard: ', err);
    });
    
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
    return tomorrow.toLocaleDateString('en-US', options); // Format dalam Bahasa Indonesia
  }
  
  
}
