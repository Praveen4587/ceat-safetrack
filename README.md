# SafeTrack Complete — Workplace Safety Reporting Portal

## What changed in this update
- The **Operation** field has been removed from the New Report form and report details.
- Downloads now use the supplied department-wise FY 2026–27 Excel templates.
- Select one department and download its own `.xlsx` workbook.
- The Downloads page no longer asks for DWM or a month.
- The original workbook sheets, existing content, formatting and sheet names are retained as far as the Excel library supports; new SafeTrack UC/UA and Near Miss reports are appended to the matching sheets.
- Existing staff, reports, actions, notifications, login and dashboard features are retained.
- The separate Staff Compliance monthly Excel export is retained.

## Requirements
- Python 3.10 or later.
- Windows 10/11 (or Linux with Python).
- No Node.js and no internet connection required for the application.
- The Excel library is bundled in `vendor/`; do not delete that folder.

## Run on Windows
1. Extract the complete ZIP to a normal folder, for example `C:\SafeTrack`. Do not run it from inside the ZIP.
2. If this is an existing SafeTrack installation, stop the old server with `Ctrl + C`.
3. Back up the old `data` folder before replacing anything. The `data` folder contains the database and uploaded evidence.
4. Copy the backed-up `data` folder into the new `SafeTrack_Complete` folder if you need your existing users and reports. Do not run first-time setup over an existing database.
5. For a brand-new installation only, double-click `FIRST_TIME_SETUP.bat` and create the administrator account. The administrator password must contain exactly 8 digits (numbers only).
6. Double-click `START_NEW_VERSION.bat`. Keep the black Command Prompt window open while using the app.
7. Open `http://127.0.0.1:8091/login`.
8. Sign in with your administrator employee code and password.

## Department Excel downloads
1. Sign in as `admin` or `safety`.
2. Open **Downloads**.
3. Select the required department/template.
4. Click **Download Excel (.xlsx)**.
5. UC and UA reports are appended to the department's UC/UA sheet. Near Miss reports are appended to its Near Miss sheet.
6. The original workbook is used as the starting template. The export does not overwrite the source template files.

The template files are stored in `templates/department_exports/`. Do not rename or delete them. The export uses report division names to decide which reports belong to each department. Make sure employee division values in Staff Directory match the appropriate department. The imported roster includes combined/legacy division labels; if a department does not receive the expected records, ask the Safety Administrator to standardize the division mapping before company-wide use.

## Staff and data
- `staff_import.csv` is a starter roster, not a guarantee of the full current company roster. Verify it against the approved staff list.
- Imported employee accounts must be provisioned with unique passwords and correct roles before use.
- The database is `data/safetrack.db`; uploaded photos are in `data/uploads/`.
- Back up the whole `data` folder after stopping the server.

## Included functions
- Employee-code login, PBKDF2 password hashing, session cookies and CSRF checks.
- Roles: staff, responsible, manager, safety and admin.
- UC, UA and Near Miss reports, location, process, description, priority and optional photo.
- Report assignment, due dates, corrective action updates, evidence, notifications and verification/closure.
- Dashboard, staff directory, monthly tracking and staff compliance.
- Department-wise Excel workbook export using the supplied company templates.
- Separate monthly staff compliance Excel export.
- SQLite database and audit records.

## Important safety and deployment warning
This is a local prototype / controlled pilot application, not approved production hosting. The built-in Python HTTP server is not hardened for public internet use. Do not expose it directly to the internet or use real employee credentials over plain HTTP.

Before real company use, the IT team must review the code and deploy it behind HTTPS with suitable authentication controls, network restrictions, backups, monitoring, data-retention rules and security testing. Immediate hazards must still be reported directly through the company's emergency and EHS process; this app is not an emergency alert service.


## Version 3.3: monthly department Excel downloads

The Downloads page requires both a department and a month. The export filters reports by the report creation month (`YYYY-MM`) and the selected department mapping, then writes only matching SafeTrack-submitted records into the supplied original Excel workbook template. Existing template worksheets and formatting are retained as far as supported by the workbook library. The New Report form keeps **Process** and does not display or collect **Operation**. Staff-compliance export remains separate.


## CEAT Safety Portal additions

- CEAT-oriented red / black / white interface theme (text wordmark; no unlicensed logo asset).
- 27 rostered employees with EPF-based numeric usernames. The employee with a missing EPF in the supplied roster is intentionally not provisioned.
- Username is the EPF digits without the leading letter. Initial password is the 4 digits repeated twice (8 digits).
- Staff users cannot change their own passwords and are not forced into a password-change page. Only the administrator can create accounts and reset passwords.
- All passwords are exactly 8 numeric digits. For an EPF such as K1434, the username is 1434 and the initial password is 14341434. Admin can set a different 8-digit password for any account, and the administrator can change their own password from the Admin password page.
- New Report includes a Responsible Division field. Submissions notify active members of that division and Safety/Admin users.
- Unread notifications appear in red and are marked read when the Notifications page is opened.
- Responsible-division users can view the report; manager/admin/safety users can manage it.
- Monthly department Excel exports are retained.

The CEAT branding is a text-only theme adaptation. Company context is based on CEAT Sri Lanka's official website: https://www.ceatsrilanka.com/about-us/.
