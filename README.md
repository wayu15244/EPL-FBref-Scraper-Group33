# EPL FBref Scraper 2024-2025

โปรเจคนี้เป็น Web Scraper ที่ใช้ดึงข้อมูลผลบอลพรีเมียร์ลีกจากเว็บ FBref.com แล้ว export ออกมาเป็น Excel ครับ

ทำโดยใช้ Java + Selenium + Apache POI

---

## สมาชิกกลุ่ม 33

| รหัส | ชื่อ |
|------|------|
| 67011211012 | ณภัทร ดีจันทึก |
| 67011211013 | ณรงค์ฤทธิ์ พิมพ์แพทย์ |
| 67011211014 | ณัฐชา จำเนียรสุข |
| 67011211017 | ณัฐวุฒิ พละศักดิ์ |

---

## วิธีดาวน์โหลดโปรเจคมาใช้งาน (สำหรับคนไม่เคยใช้ Git)

### ขั้นตอนที่ 1: ติดตั้ง Git
- ดาวน์โหลด Git จาก https://git-scm.com/downloads
- ติดตั้งตามปกติ กด Next ไปเรื่อยๆ

### ขั้นตอนที่ 2: Clone โปรเจค
- เปิด Command Prompt หรือ PowerShell
- พิมพ์คำสั่งนี้:
```
git clone https://github.com/wayu15244/EPL-FBref-Scraper-Group33.git
```
- จะได้โฟลเดอร์ชื่อ `EPL-FBref-Scraper-Group33` มา

### ขั้นตอนที่ 3: เข้าไปในโฟลเดอร์
```
cd EPL-FBref-Scraper-Group33
```

---

## ต้องใช้อะไรบ้าง

- Java 11+ (ดาวน์โหลดจาก https://adoptium.net/)
- Chrome Browser
- ChromeDriver (version ต้องตรงกับ Chrome ที่ใช้)
  - ดาวน์โหลดจาก https://chromedriver.chromium.org/downloads
  - แตกไฟล์แล้ววางไว้ในโฟลเดอร์โปรเจค

---

## โปรแกรมทำอะไรได้บ้าง

- ดึงข้อมูลได้ทั้ง 380 นัดในซีซั่น 2024-2025
- ได้ข้อมูลพวก Goals, Assists, xG, Big Chances, Through Balls, Blocks ฯลฯ
- export เป็น Excel มี Dashboard แยกแต่ละนัด
- เวลาที่แสดงเป็นเวลาอังกฤษ (BST/GMT)

---

## วิธีรันโปรแกรม

### ขั้นตอนที่ 1: คอมไพล์
เปิด Command Prompt ในโฟลเดอร์โปรเจค แล้วพิมพ์:
```
javac -encoding UTF-8 -cp ".;lib/*" src/FBrefScraper.java
```

### ขั้นตอนที่ 2: รันโปรแกรม

รันแค่ 5 นัดสำหรับทดสอบ:
```
java -cp ".;lib/*;src" FBrefScraper --test=5
```

รันเต็มๆ 380 นัด (ใช้เวลานานมาก):
```
java -cp ".;lib/*;src" FBrefScraper
```

---

## Output

รันเสร็จจะได้ไฟล์ `EPL_2024_2025_FBref.xlsx` ที่มีข้อมูลแต่ละนัดแยก sheet

---

## ถ้ามีปัญหา

- ถ้า Chrome เปิดไม่ได้ ลองเช็คว่า ChromeDriver version ตรงกับ Chrome หรือเปล่า
- ถ้า compile ไม่ผ่าน ลองเช็คว่ามี Java ติดตั้งแล้วหรือยัง พิมพ์ `java -version` ดู
