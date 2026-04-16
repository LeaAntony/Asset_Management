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

---

## ✨ Features

### 🔐 Authentication & Authorization
- Role-based access control dengan level: **Admin**, **Requestor**, **Approver**, **Finance (FBP)**
- Cookie Authentication dengan session management

### 📦 Asset Management
- **Asset List** — Daftar lengkap aset dengan informasi: nomor aset, deskripsi, kelas, cost center, nilai buku, dan status

### 🔄 Transfer & Gate Pass
- **Gate Pass (Asset)** — Formulir dan daftar transfer aset dengan persetujuan
- **Gate Pass (Non-Asset)** — Transfer untuk item non-aset

### ✅ Approval Workflow
- Alur persetujuan bertingkat: **Requestor → Approver → Finance**
- Status tracking real-time pada setiap tahap
- Email notifikasi otomatis ke pihak terkait

### 💰 Finance Management
- **Proforma** — Dokumen proforma untuk transfer/disposal
- **PEB (Pemberitahuan Ekspor Barang)** — Dokumen ekspor aset

### 🚢 Shipping
- Manajemen pengiriman untuk aset maupun non-aset
- Pelacakan status pengiriman

### 📊 Dashboard & Reports
- Dashboard interaktif dengan grafik dan statistik
- Dashboard terpisah untuk aset dan non-aset

### 👥 Admin Panel
- Manajemen pengguna dan hak akses
- Konfigurasi sistem

---

## 🛠️ Tech Stack

| Layer | Technology |
|-------|-----------|
| **Framework** | ASP.NET Core MVC (.NET 10) |
| **Database** | Microsoft SQL Server |
| **ORM** | Entity Framework Core 6 + Raw SQL (DAL) |
| **Authentication** | Cookie Auth |
| **UI Framework** | AdminLTE 3 + Bootstrap 5 |
| **Email** | MailKit |
| **Excel Export** | ClosedXML, EPPlus |
| **PDF Export** | iText7, Select.HtmlToPdf |
| **Image Processing** | SixLabors.ImageSharp |
| **JSON** | Newtonsoft.Json |
| **Dynamic Queries** | System.Linq.Dynamic.Core |

---

## ⚙️ Installation

### Prerequisites

- [.NET 10 SDK](https://dotnet.microsoft.com/download)
- [SQL Server](https://www.microsoft.com/sql-server) (2019 atau lebih baru)
- [Visual Studio 2022/2026](https://visualstudio.microsoft.com/) atau VS Code

### Steps

**1. Clone repository**
```bash
git clone https://github.com/LeaAntony/Asset_Management.git
cd Asset_Management
```

**2. Konfigurasi database & SSO**

Buka `DatabaseAccessLayer.cs` dan sesuaikan:
```json
public string ConnectionString = "Data Source=SERVER-NAME;Initial Catalog=DATABASE-NAME;Integrated Security=true;Persist Security Info=True;MultipleActiveResultSets=true";
```

**3. Restore & run**
```bash
dotnet restore
dotnet run
```

**5. Buka browser**
```
https://localhost:XXXX
```

---

## 🗂️ Project Structure

```
Asset_Management/
├── Controllers/
│   ├── AuthController.cs        # Authentication
│   ├── HomeController.cs        # Dashboard & Login
│   ├── AdminController.cs       # Admin management
│   ├── RequestorController.cs   # Requestor features
│   ├── ApproverController.cs    # Approval workflow
│   ├── FinanceController.cs     # Finance
│   ├── EmailController.cs       # Email notifications
│   ├── ExportController.cs      # Excel & PDF export
│   └── ShippingController.cs    # Shipping management
│
├── Models/                      # Data models
│   ├── AssetListModel.cs
│   ├── GatePassModel.cs
│   └── ...
│
├── Views/                       # Razor Views
│   ├── Home/
│   │   ├── Dashboard.cshtml
│   │   └── DashboardNonAsset.cshtml
│   ├── Admin/
│   ├── Requestor/
│   ├── Approver/
│   ├── Finance/
│   └── Shared/
│       └── _Layout.cshtml
│
├── Function/
│   ├── ApplicationDbContext.cs  # EF Core DbContext
│   └── DatabaseAccessLayer.cs  # DAL / Raw SQL
│
├── Service/
│   └── ITokenService.cs        # Token service interface
│
└── wwwroot/
    ├── Upload/                  # File uploads
    │   ├── Asset/
    │   ├── GatePass/
    │   ├── Disposal/
    │   ├── PEB/
    │   └── Proforma/
    └── lib/                     # Client-side libraries
```


## 📄 License

This project is proprietary software. All rights reserved.

---

<div align="center">
  Built with ❤️ using <strong>ASP.NET Core .NET 10</strong>
</div>