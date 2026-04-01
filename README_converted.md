# คู่มือการติดตั้งโปรเจกต์ (Django + SQL Server)

## ส่วนที่ 1: ติดตั้ง SQL Server, SSMS และ ODBC Driver 17

1. ติดตั้ง **SQL Server 2019** โดยเลือกติดตั้งแบบ **SQL Server Authentication**
2. ติดตั้ง **SSMS (SQL Server Management Studio)**
3. ติดตั้ง **ODBC Driver 17 for SQL Server**


## ส่วนที่ 2: การตั้งเเละเชื่อมต่อฐานข้อมูลของโปรเจกต์

### 1) ติดตั้ง Library พื้นฐาน

```bash
pip install django mssql-django django-crispy-forms crispy-bootstrap5 openpyxl
```

### 2) นำข้อมูลเข้า (Restore Database)

- ใน **SSMS** คลิกขวาที่ **Databases** > **Restore Database...**
- เลือก **Device** และเลือกไฟล์ `.bak` ที่ได้รับมา
- การ Restore วิธีนี้จะทำให้ **View, Function และ Trigger** มาครบทั้งหมด

### 3) ตั้งค่าการเชื่อมต่อใน `settings.py`

แก้ไขส่วน `DATABASES` ให้ตรงกับเครื่อง:

```python
DATABASES = {
    'default': {
        'ENGINE': 'mssql',
        'NAME': 'LaundryDB',
        'USER': 'sa',
        'PASSWORD': '123456789',
        'HOST': 'ใส่ชื่อ device ตัวเอง เช่น DESKTOP-6445F79\\SQLEXPRESS2019',
        'PORT': '',
        'OPTIONS': {
            'driver': 'ODBC Driver 17 for SQL Server',
            'extra_params': 'TrustServerCertificate=yes;',
        },
    }
}
```

### 4) เริ่มใช้งานระบบ

```bash
python manage.py runserver
```
จากนั้นกด Ctrl + คลิก

### 5) เข้าใช้งานระบบ

### - หน้าล็อคอิน
[![login](https://img2.pic.in.th/login.jpg)](https://pic.in.th/image/login.86ORpM)


```bash
- เข้าสู่ระบบในฐานะ Admin ด้วย 
username : admin 
password :1234

- เข้าสู่ระบบในฐานะ พนักงาน 1 ด้วย 
username : staff01
password : staff01

- เข้าสู่ระบบในฐานะ พนักงาน 2 ด้วย 
username : staff02
password : staff02

- เข้าสู่ระบบในฐานะ พนักงาน 3 ด้วย 
username : staff03
password : staff03
```

### 6) เข้าสู่ระบบเรียบร้อย

 ### - เเสดงหน้าตาระบบ

[![Screenshot 2026 04 01 210608](https://img2.pic.in.th/Screenshot-2026-04-01-210608.png)](https://pic.in.th/image/Screenshot-2026-04-01-210608.86QY7z)


 