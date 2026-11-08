"""
Launch CEAT Safety Workplace System with a live public URL.
Starts the web server and a secure Cloudflare Tunnel.
"""
import os
import re
import sys
import time
import socket
import webbrowser
import subprocess
from pathlib import Path

ROOT = Path(__file__).resolve().parent
CLOUDFLARED = ROOT / "cloudflared.exe"
PORT = int(os.environ.get("PORT", "8091"))

def get_lan_ip():
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("8.8.8.8", 80))
        ip = s.getsockname()[0]
        s.close()
        return ip
    except Exception:
        return "127.0.0.1"

def is_port_in_use(port):
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        return s.connect_ex(('127.0.0.1', port)) == 0

def main():
    print("=" * 70)
    print("   CEAT SAFETY WORKPLACE REPORTING SYSTEM - ONLINE SERVER")
    print("=" * 70)

    # 1. Start Python server if not already running
    server_proc = None
    if not is_port_in_use(PORT):
        print(f"[*] Starting SafeTrack web server on 0.0.0.0:{PORT}...")
        server_proc = subprocess.Popen(
            [sys.executable, str(ROOT / "app.py"), "serve", "--host", "0.0.0.0", "--port", str(PORT)],
            cwd=str(ROOT),
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True
        )
        # Give it a moment to bind
        time.sleep(1.5)
    else:
        print(f"[*] SafeTrack server is already running on port {PORT}.")

    lan_ip = get_lan_ip()

    # 2. Check cloudflared.exe
    if not CLOUDFLARED.exists():
        print("[!] cloudflared.exe not found in app directory.")
        print(f"[*] Local access: http://localhost:{PORT}")
        print(f"[*] LAN access:   http://{lan_ip}:{PORT}")
        if server_proc:
            server_proc.wait()
        return

    # 3. Start Cloudflare Tunnel
    print("[*] Generating secure public URL via Cloudflare...")
    tunnel_cmd = [str(CLOUDFLARED), "tunnel", "--url", f"http://127.0.0.1:{PORT}"]
    tunnel_proc = subprocess.Popen(
        tunnel_cmd,
        cwd=str(ROOT),
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        text=True,
        encoding="utf-8",
        errors="replace"
    )

    public_url = None
    deadline = time.time() + 30

    while time.time() < deadline:
        line = tunnel_proc.stdout.readline()
        if not line:
            if tunnel_proc.poll() is not None:
                break
            continue
        m = re.search(r"https://[a-zA-Z0-9-]+\.trycloudflare\.com", line)
        if m:
            public_url = m.group(0)
            break

    print("\n" + "=" * 70)
    print("   SERVER IS NOW ONLINE AND ACCESSIBLE!")
    print("=" * 70)
    if public_url:
        print(f"\n  >> PUBLIC WEB URL (Any device / internet):")
        print(f"     {public_url}/login\n")
    else:
        print("\n  [!] Could not parse tunnel URL immediately.")

    print(f"  >> LOCAL NETWORK URL (Wi-Fi / Office):")
    print(f"     http://{lan_ip}:{PORT}/login\n")

    print(f"  >> LOCAL COMPUTER URL:")
    print(f"     http://127.0.0.1:{PORT}/login\n")

    print("-" * 70)
    print("  LOGIN CREDENTIALS:")
    print("  * Administrator: Code: SAFETYADMIN  | Password: 12345678")
    print("  * Staff example: Code: 1434         | Password: 14341434")
    print("-" * 70)
    print("  Press Ctrl + C in this window anytime to stop the server.")
    print("=" * 70 + "\n")

    if public_url:
        try:
            webbrowser.open(f"{public_url}/login")
        except Exception:
            pass

    try:
        tunnel_proc.wait()
    except KeyboardInterrupt:
        print("\n[*] Stopping server and tunnel...")
    finally:
        tunnel_proc.terminate()
        if server_proc:
            server_proc.terminate()
        print("[*] Server stopped cleanly.")

if __name__ == "__main__":
    main()
