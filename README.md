<div align="center">

# 🏢 Asset Management System

![.NET](https://img.shields.io/badge/.NET_10-512BD4?style=for-the-badge&logo=dotnet&logoColor=white)
![ASP.NET Core](https://img.shields.io/badge/ASP.NET_Core_MVC-0078D4?style=for-the-badge&logo=microsoft&logoColor=white)
![SQL Server](https://img.shields.io/badge/SQL_Server-CC2927?style=for-the-badge&logo=microsoftsqlserver&logoColor=white)
![Bootstrap](https://img.shields.io/badge/Bootstrap_5-7952B3?style=for-the-badge&logo=bootstrap&logoColor=white)
![AdminLTE](https://img.shields.io/badge/AdminLTE_3-3C8DBC?style=for-the-badge&logo=adminlte&logoColor=white)

**Sistem manajemen aset perusahaan berbasis web yang komprehensif mencakup transfer aset, alur persetujuan, hingga pelaporan.**

</div>

---

## 📋 Overview

**Asset Management System** adalah aplikasi web enterprise yang dibangun menggunakan **ASP.NET Core MVC (.NET 10)** untuk mengelola siklus hidup aset perusahaan secara end-to-end. Mulai dari pencatatan aset, permintaan transfer, proses persetujuan berlapis, yang semua terkelola dalam satu platform terintegrasi.

Aplikasi ini dirancang dengan mengikuti **best practices** arsitektur perangkat lunak, prinsip **Clean Code**, serta standar keamanan **OWASP Top 10**.

---

## ✨ Features

### 🔐 Authentication & Authorization

- Role-based access control dengan level: **Admin**, **Requestor**, **Approver**, **Finance (FBP)**, **Security**
- Cookie Authentication dengan konfigurasi HttpOnly, Secure, SameSite
- Policy-based authorization dengan 8+ kebijakan akses
- Auto-refresh claims setelah update data user

### 📦 Asset Management

- **Asset List** — Daftar lengkap aset dengan informasi: nomor aset, deskripsi, kelas, cost center, nilai buku, dan status

### 🔄 Transfer & Gate Pass

- **Gate Pass (Asset)** — Formulir dan daftar transfer aset dengan persetujuan
- **Gate Pass (Non-Asset)** — Transfer untuk item non-aset (consumable, equipment)

### ✅ Approval Workflow

- Alur persetujuan bertingkat: **Requestor → Approver (HOD/FBP/PH) → Finance**
- Delegation approval — approver dapat mendelegasikan tugas ke pengguna lain
- Status tracking real-time pada setiap tahap
- Email notifikasi otomatis ke pihak terkait (via MailKit)

### 💰 Finance Management

- **Proforma Invoice** — Dokumen proforma untuk transfer/disposal
- **PEB (Pemberitahuan Ekspor Barang)** — Dokumen ekspor aset
- **Invoice Management** — Pembuatan dan persetujuan invoice
- Export PDF & Excel untuk pelaporan

### 🚢 Shipping

- Manajemen pengiriman untuk aset maupun non-aset
- Box management dengan dimensi & berat
- DHL AWB tracking
- Pelacakan status pengiriman

### 📊 Dashboard & Reports

- Dashboard interaktif dengan grafik
- Dashboard terpisah untuk aset dan non-aset
- Export data ke Excel & PDF

### 👥 Admin Panel

- Manajemen pengguna dan hak akses
- Delegasi approver
- Konfigurasi master data (department, plant, vendor, dll)

---

## 🛠️ Tech Stack

| Layer                | Technology                                         |
| -------------------- | -------------------------------------------------- |
| **Framework**        | ASP.NET Core MVC (.NET 10)                         |
| **Database**         | Microsoft SQL Server (2019+)                       |
| **ORM**              | Entity Framework Core 6 + Stored Procedures (DAL)  |
| **Authentication**   | Cookie Authentication + ASP.NET Core Identity      |
| **UI Framework**     | AdminLTE 3 + Bootstrap 5                           |
| **JavaScript**       | jQuery, DataTables, Chart.js, Select2, SweetAlert2 |
| **Email**            | MailKit + MimeKit                                  |
| **Excel Export**     | ClosedXML, EPPlus                                  |
| **PDF Export**       | iText7, Select.HtmlToPdf                           |
| **Image Processing** | SixLabors.ImageSharp                               |
| **Password Hashing** | BCrypt.Net-Next (bcrypt, work factor 12)           |
| **JSON**             | Newtonsoft.Json, System.Text.Json                  |
| **Dynamic Queries**  | System.Linq.Dynamic.Core                           |
| **Containerization** | Docker                                             |

---

## 🔒 Security & OWASP Top 10 Compliance

Aplikasi ini menerapkan prinsip keamanan sesuai standar OWASP Top 10:

| OWASP Category                     | Implementasi                                                                                |
| ---------------------------------- | ------------------------------------------------------------------------------------------- |
| **A1 - Broken Access Control**     | Policy-based authorization dengan role claims (`RequireRequestor`, `RequireApprover`, dll.) |
| **A2 - Cryptographic Failures**    | Password di-hash dengan **bcrypt** (work factor 12), cookie dengan Secure + HttpOnly        |
| **A3 - Injection**                 | Semua query menggunakan **Stored Procedures** atau **parameterized queries**                |
| **A4 - Insecure Design**           | Arsitektur MVC dengan separation of concerns                                                |
| **A5 - Security Misconfiguration** | CORS terbatas ke origin yang terdaftar, HSTS di production                                  |
| **A6 - Vulnerable Components**     | Seluruh dependency di-referensi via NuGet, update berkala                                   |
| **A7 - Auth Failures**             | Session timeout, cookie sliding expiration, anti-forgery tokens                             |
| **A8 - Integrity Failures**        | CSRF protection via `[ValidateAntiForgeryToken]`                                            |
| **A9 - Logging Failures**          | Logging via `ILogger<T>` di seluruh controller                                              |
| **A10 - SSRF**                     | HttpClientFactory dengan restricted endpoints                                               |

### 🔐 Detail Keamanan

- **CSRF Protection**: `@Html.AntiForgeryToken()` di semua form + `[ValidateAntiForgeryToken]` di POST endpoints
- **Password Security**: bcrypt dengan work factor 12 (menggantikan MD5 legacy)
- **Cookie Security**: HttpOnly, Secure (Always), SameSite Strict
- **SQL Injection Prevention**: Semua query via Stored Procedures
- **CORS**: Dibatasi hanya untuk origin yang terdaftar di `appsettings.json`

---

## ⚙️ Installation

### Prerequisites

- [.NET 10 SDK](https://dotnet.microsoft.com/download/dotnet/10.0)
- [SQL Server](https://www.microsoft.com/sql-server) (2019 atau lebih baru)
- [Visual Studio 2022/2026](https://visualstudio.microsoft.com/) atau [VS Code](https://code.visualstudio.com/)
- [Git](https://git-scm.com/)

### Setup & Installation

**1. Clone repository**

```bash
git clone https://github.com/LeaAntony/Asset_Management.git
cd Asset_Management
```

**2. Setup Database**

Jalankan script `asset.sql` di SQL Server Management Studio (SSMS) untuk membuat database dan seluruh objek:

```bash
# Atau restore via sqlcmd
sqlcmd -S localhost -i asset.sql
```

**3. Konfigurasi Connection String**

Buka `appsettings.json` dan sesuaikan konfigurasi database:

```json
{
  "ConnectionStrings": {
    "DefaultConnection": "Data Source=localhost;Initial Catalog=Asset_Management;Integrated Security=true;Persist Security Info=True;MultipleActiveResultSets=true"
  }
}
```

> **Catatan**: Atau bisa juga sesuaikan di `DatabaseAccessLayer.cs` jika tidak menggunakan `appsettings.json`.

**4. Konfigurasi Email (Opsional)**

Di `appsettings.json`, tambahkan konfigurasi untuk email:

```json
public string ConnectionString = "Data Source=SERVER-NAME;Initial Catalog=DATABASE-NAME;Integrated Security=true;Persist Security Info=True;MultipleActiveResultSets=true";
```

**6. Restore & Run**

```bash
dotnet restore
dotnet build
dotnet run
```

**7. Buka browser**

```
https://localhost:XXXX
```

Atau sesuai dengan konfigurasi di `Properties/launchSettings.json`.

---

## 🐳 Docker Support

### Prasyarat

- [Docker Desktop](https://www.docker.com/products/docker-desktop/)
- [Docker Compose](https://docs.docker.com/compose/)

### Menjalankan dengan Docker Compose

```bash
# Build dan jalankan container
docker-compose up -d

# Aplikasi akan berjalan di:
# https://localhost:XXXX
```

### Struktur Docker

- **Dockerfile** — Multi-stage build untuk .NET 10 aplikasi
- **docker-compose.yml** — Menggabungkan aplikasi + SQL Server container
- SQL Server menggunakan image `mcr.microsoft.com/mssql/server:2022-latest`

> Lihat file `Dockerfile` dan `docker-compose.yml` di root project untuk detail konfigurasi.

---

## 🗂️ Project Structure

```
Asset_Management/
│
├── Controllers/                  # Request handlers (MVC Controllers)
│   ├── AuthController.cs         #   Authentication (token refresh)
│   ├── HomeController.cs         #   Dashboard, Login (Manual & Security)
│   ├── AdminController.cs        #   User management, delegation
│   ├── RequestorController.cs    #   Gate pass request, asset list
│   ├── ApproverController.cs     #   Approval workflow
│   ├── FinanceController.cs      #   Proforma, invoice, PEB
│   ├── EmailController.cs        #   Email notifications
│   ├── ExportController.cs       #   Excel & PDF export
│   ├── ShippingController.cs     #   Shipping management
│   ├── AccountController.cs      #   Access denied handler
│   ├── RedirectController.cs     #   Role-based redirect routing
│   └── UserController.cs         #   Password change
│
├── Models/                       # Data models / ViewModels
│   ├── AssetListModel.cs
│   ├── GatePassModel.cs
│   ├── ApprovalModel.cs
│   ├── UserDetailModel.cs
│   ├── ChartModel.cs
│   ├── InvoiceModel.cs
│   ├── ProformaModel.cs
│   └── ...
│
├── Views/                        # Razor Views
│   ├── Home/                     #   Login, Dashboard
│   ├── Admin/                    #   User management
│   ├── Requestor/                #   Gate pass, asset list
│   ├── Approver/                 #   Approval list
│   ├── Finance/                  #   Proforma, invoice
│   ├── Shipping/                 #   Shipping management
│   └── Shared/                   #   _Layout.cshtml, partials
│
├── Function/                     # Helper / utility classes
│   ├── ApplicationDbContext.cs   #   EF Core DbContext
│   ├── DatabaseAccessLayer.cs    #   Data access via Stored Procedures
│   ├── Authentication.cs         #   bcrypt hashing, PKCE code verifier
│   ├── SessionCheck.cs           #   Session validation helper
│   ├── TagHelper.cs              #   HTML helper extensions
│   └── Utility.cs                #   Formatting utilities
│
├── Service/                      # Business logic services
│   ├── TokenService.cs           #   OAuth token refresh
│   ├── FileManagementService.cs  #   File upload handling
│   ├── ExcelServiceProvider.cs   #   Excel import/export
│   └── ImportExportFactory.cs    #   Asset import factory
│
├── wwwroot/                      # Static files
│   ├── css/                      #   Custom stylesheets
│   ├── js/                       #   Custom JavaScript
│   ├── images/                   #   Images & logos
│   ├── lib/                      #   Client-side libraries (AdminLTE, Bootstrap, etc.)
│   ├── Upload/                   #   File upload directories
│   │   ├── Asset/
│   │   ├── GatePass/
│   │   ├── Disposal/
│   │   ├── PEB/
│   │   └── Proforma/
│   └── database/                 #   Database scripts
│
├── appsettings.json              # Main configuration
├── appsettings.Development.json  # Development configuration
├── appsettings.Production.json   # Production configuration
├── sp_password_operations.sql    # Stored procedures for password management
├── asset.sql                     # Database creation script
├── Dockerfile                    # Docker image definition
├── docker-compose.yml            # Docker Compose configuration
├── Program.cs                    # Application entry point & middleware
└── Asset_Management.csproj       # Project file
```

---

## 📋 Arsitektur & Best Practices

### Arsitektur

- **MVC Pattern** — Pemisahan Controller, Model, View
- **Repository-like Pattern** — `DatabaseAccessLayer` sebagai abstraction layer untuk data access
- **Dependency Injection** — Services diregistrasi di `Program.cs` via DI container
- **Middleware Pipeline** — Request/response pipeline terstruktur (CORS → Auth → Session → MVC)

### Clean Code Principles

- **Naming conventions** — Method names deskriptif (PascalCase untuk C#, camelCase untuk JS)
- **Separation of concerns** — Controller hanya handle request, logic di Service layer
- **Exception handling** — Try-catch dengan logging di seluruh controller
- **DRY Principle** — Fungsi reusable di `Function/` dan `Service/`

### Security Best Practices

- ✅ **bcrypt password hashing** (not MD5)
- ✅ **Parameterized queries / Stored Procedures** (no SQL injection)
- ✅ **CSRF tokens** on all POST forms
- ✅ **HttpOnly + Secure + SameSite** cookies
- ✅ **CORS restricted** to specific origins
- ✅ **HSTS** enabled in production
- ✅ **Input sanitization** via model binding
- ✅ **Role-based authorization** policies

---

## 👥 Contributors / Pengembang

| Nama           | Peran                | Kontak                                 |
| -------------- | -------------------- | -------------------------------------- |
| **Lea Antony** | Full-Stack Developer | [GitHub](https://github.com/LeaAntony) |

---

## 📄 License

This project is proprietary software. All rights reserved.

---

<div align="center">
  Built with ❤️ using <strong>ASP.NET Core .NET 10</strong>
</div>
