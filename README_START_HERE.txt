CEAT WORKPLACE SAFETY REPORTING SYSTEM
====================================

VERSION: CEAT Safety Portal (based on SafeTrack v3.3.1)

FIRST RUN / UPGRADE
-------------------
1. Extract this ZIP into a new folder.
2. For an existing installation, stop the old server (Ctrl+C) and BACK UP the entire data folder.
3. Copy the old data folder into this new SafeTrack_Complete folder.
4. Do NOT run FIRST_TIME_SETUP.bat on an existing database.
5. Run START_NEW_VERSION.bat and keep its command window open.
6. Open http://127.0.0.1:8091/login

FIRST-TIME INSTALLATION
-----------------------
1. Install Python 3.10 or newer.
2. Run FIRST_TIME_SETUP.bat.
3. Create the administrator username and an exactly 8-digit numeric password.
4. Run START_NEW_VERSION.bat.

EMPLOYEE LOGIN
--------------
The supplied employee roster contains 27 staff members with EPF numbers.
The staff member whose EPF number was blank in the supplied roster is NOT created automatically.
An administrator can create that account later when the EPF number is available.

For EPF numbers in the format K1434:
- Username: 1434 (remove the first letter; keep all digits including leading zeroes)
- Initial password: 14341434 (the four digits repeated twice)
- Password: the four digits repeated twice (8 digits, numbers only).

Employees cannot change their own passwords and are not sent to a forced password-change page.
Only the administrator can create accounts and reset passwords. Admin can create accounts from
Staff Directory. The form collects full name, EPF, designation, division, role and email.
Username and initial password are generated from the EPF. Admin password resets must contain
exactly 8 digits and take effect immediately. The administrator can change their own password
from the Admin password page.

SAFETY REPORTING
----------------
New Report fields include report type, preliminary priority, location, process, responsible
division, observation description and optional photo evidence. The old Operation field is
not collected. After submission, notifications are created for active users in the selected
responsible division and for Safety/Admin users. Unread notifications are highlighted in red.
Users in the responsible division can open the report; manager/admin/safety roles can manage it.

DOWNLOADS
---------
Select a department and month to download the matching original Excel workbook.
Only SafeTrack reports created in the selected month and mapped to the selected responsible
division are added to the template. Staff Compliance export remains separate.

DATA AND SECURITY
-----------------
Database: data/safetrack.db
Uploaded evidence: data/uploads/
Back up the entire data folder while the server is stopped.

IMPORTANT SECURITY NOTE
-----------------------
EPF-derived initial passwords are predictable. The administrator should reset them to unique
8-digit numeric passwords before wider use. Eight-digit numeric passwords are weaker than
longer passwords.
For company production use, use an IT-approved HTTPS deployment, strong authentication policy,
backups, monitoring, access review and a security assessment. Do not expose this local
development server directly to the public internet.

CEAT v4.3 DEPARTMENT UI
The New Report page displays exactly 14 official departments as visual cards.
Hover over a card to see the interaction effect; click a card to select it.
The Admin account creation division list uses the same 14 official departments.


CEAT DARK THEME UPDATE
- This build includes a dark charcoal / CEAT red interface, hover transitions,
  a local tyre-manufacturing background image on the login screen, and a CEAT
  wordmark asset.
- Planning is available as a responsible division. Existing workbook download
  templates remain limited to departments with supplied Excel templates.
- Existing data can be preserved by backing up and copying the complete data
  folder as described above. Do not run FIRST_TIME_SETUP on an existing database.


V4.7 EXPORT SOURCE-OF-TRUTH UPDATE
- The CEAT logo image supplied by the user is used in the application navigation.
- Department Excel files are used as templates only.
- Existing records embedded in the supplied Excel workbooks are cleared from exported workbooks.
- Monthly department exports contain only reports stored in this application's database
  for the selected month and responsible division.
- No old Excel records are imported into the application's database.
- Planning has its own workbook based on the existing General template layout.
- Before public/company deployment, migrate SQLite to managed PostgreSQL and configure
  persistent photo storage. Do not host real employee data on an ephemeral free web host.
