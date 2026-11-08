"""Local-only administrator credential recovery for CEAT Safety."""
import getpass
import secrets
import sqlite3
from app import con, init, import_staff, hp, set_pw


def main():
    init()
    with con() as c:
        existing = c.execute("SELECT code,name FROM users WHERE role='admin' ORDER BY id").fetchall()
    print('=== CEAT Safety Administrator Login Repair ===')
    print('This tool does not delete any existing safety reports.')
    if existing:
        print('Existing administrator codes: ' + ', '.join(row['code'] for row in existing))
    else:
        print('No admin account found. You can create one.')
    print('Only run this tool locally if you are authorized to administer this system.')
    code = input('Admin employee code to create/reset [SAFETYADMIN]: ').strip().upper() or 'SAFETYADMIN'
    if not all(c.isalnum() or c in '-_' for c in code):
        print('Invalid employee code. Use only letters, numbers, - or _.'); return 1
    first = getpass.getpass('New admin password (exactly 8 digits): ')
    if not (len(first) == 8 and first.isdigit()):
        print('Password must be exactly 8 digits; no changes made.'); return 1
    second = getpass.getpass('Confirm new admin password: ')
    if first != second:
        print('Passwords did not match; no changes made.'); return 1
    with con() as c:
        u = c.execute('SELECT id FROM users WHERE code=?', (code,)).fetchone()
        if u:
            aid = u['id']
            c.execute("UPDATE users SET role='admin', active=1, must_change_password=0 WHERE id=?",(aid,))
        else:
            salt = secrets.token_hex(16)
            aid = c.execute("INSERT INTO users(code,name,designation,division,role,salt,pwhash,active,must_change_password) VALUES(?,?,?,?,?,?,?,1,0)", (code,'System Administrator','Administrator','SAFETY','admin',salt,hp(first,salt))).lastrowid
        set_pw(c, aid, first)
        c.execute('DELETE FROM sessions WHERE uid=?', (aid,))
        c.execute("INSERT INTO audit(uid,event,rid,created) VALUES(?, 'local_admin_password_reset', NULL, datetime('now'))", (aid,))
    print('\nSUCCESS: Admin account is ready.')
    print('LOGIN EMPLOYEE CODE: ' + code)
    print('LOGIN PASSWORD: The new password you just entered.')
    print('Start START_SERVER.bat and open http://127.0.0.1:8091/login')
    return 0

if __name__ == '__main__':
    raise SystemExit(main())
