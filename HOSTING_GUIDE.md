# SafeTrack - Server Hosting & Access Guide

## 1. Instant Online Server (Live Now)
The server is currently running and publicly accessible via Cloudflare Tunnel:

- **Public Web URL (Internet / Mobile / Anywhere)**:
  `https://modeling-wendy-flush-east.trycloudflare.com/login`
- **Local Network URL (Same Wi-Fi / Office LAN)**:
  `http://192.168.8.177:8091/login`
- **Local Machine URL**:
  `http://127.0.0.1:8091/login`

### How to Start the Online Server Anytime
Simply double-click:
`SafeTrack_Complete\START_ONLINE_SERVER.bat`
This automatically launches the server, generates an HTTPS URL, prints the link, and opens it in your default browser.

---

## 2. Login Credentials

### Administrator Account
- **Employee Code**: `SAFETYADMIN`
- **Password**: `12345678`
*(Admin can manage all reports, staff members, compliance, and department downloads).*

### Employee Roster Accounts
Use 4-digit EPF numbers from `staff_import.csv`:
- **Format**: Digits only (e.g., for `K1434`, code is `1434`)
- **Initial Password**: The 4 digits repeated twice (e.g., `14341434`)
- **Examples**:
  - `1434` / `14341434` (Nalin Udayanga, Division A)
  - `0312` / `03120312` (P.M.I.R. Pemasiri, Safety)
  - `0150` / `01500150` (R.A.C. Ranaweera, Planning)

---

## 3. Permanent 24/7 Cloud Deployment (e.g., Render.com)

If you need the portal hosted in the cloud permanently so it stays online even when this computer is turned off:

### Option A: Render.com (Recommended - Free Tier Available)
1. Push this project to GitHub (a private repository is recommended).
2. Go to [https://dashboard.render.com](https://dashboard.render.com) and click **New + > Web Service**.
3. Connect your GitHub repository.
4. Set:
   - **Environment**: `Python 3`
   - **Build Command**: *(leave empty)*
   - **Start Command**: `python app.py serve`
5. Render will automatically detect the port and provide a permanent `https://<your-service>.onrender.com` URL.

### Option B: Docker Container Hosting (Railway, Koyeb, Fly.io, AWS, VPS)
A preconfigured [Dockerfile](file:///C:/Users/Praveen/Desktop/CEAT_Workplace_Safety_Reporting_System_v4_7_CEAT_Logo_System_Data_Export%20(2)/SafeTrack_Complete/Dockerfile) is included. You can build and run it anywhere:
```bash
docker build -t ceat-safetrack .
docker run -p 8091:8091 ceat-safetrack
```
