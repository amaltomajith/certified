"""
Quick SMTP diagnostic — tests the connection and sends a tiny test email.
Run from project root:  python test_smtp.py
"""
import smtplib
import ssl
import sys
import os

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "backend"))
from config_utils import read_config
from session_store import load_from_disk, get_app_password, get_sender_email

load_from_disk()

cfg = read_config()
ecfg = cfg["email"]
password = get_app_password()
sender = get_sender_email() or ecfg.get("sender_email", "")

print(f"SMTP host : {ecfg['smtp_host']}")
print(f"SMTP port : {ecfg['smtp_port']}")
print(f"Sender    : {sender}")
print(f"Password  : {'*' * len(password) if password else '(NOT SET)'}")
print()

# ── Test 1: Port 587 + STARTTLS ──
print("=" * 50)
print("TEST 1: SMTP port 587 + STARTTLS")
print("=" * 50)
try:
    ctx = ssl.create_default_context()
    srv = smtplib.SMTP(ecfg["smtp_host"], 587, timeout=30)
    srv.set_debuglevel(2)  # full SMTP conversation
    print("[+] Connected")
    srv.ehlo()
    print("[+] EHLO done")
    srv.starttls(context=ctx)
    print("[+] STARTTLS done")
    srv.ehlo()
    print("[+] EHLO (post-TLS) done")
    srv.login(sender, password)
    print("[+] LOGIN success")

    # Try sending a tiny test email to ourselves
    from email.message import EmailMessage
    msg = EmailMessage()
    msg["Subject"] = "SMTP Test — delete me"
    msg["From"] = sender
    msg["To"] = sender
    msg.set_content("If you receive this, SMTP works.")

    srv.send_message(msg)
    print("[+] send_message() SUCCESS")
    srv.quit()
    print("[+] QUIT — Test 1 PASSED\n")
except Exception as exc:
    import traceback
    print(f"\n[-] Test 1 FAILED: {exc}")
    traceback.print_exc()
    print()

# ── Test 2: Port 465 + SMTP_SSL ──
print("=" * 50)
print("TEST 2: SMTP_SSL port 465")
print("=" * 50)
try:
    ctx = ssl.create_default_context()
    srv = smtplib.SMTP_SSL(ecfg["smtp_host"], 465, timeout=30, context=ctx)
    srv.set_debuglevel(2)
    print("[+] Connected (SSL)")
    srv.ehlo()
    print("[+] EHLO done")
    srv.login(sender, password)
    print("[+] LOGIN success")

    from email.message import EmailMessage
    msg = EmailMessage()
    msg["Subject"] = "SMTP_SSL Test — delete me"
    msg["From"] = sender
    msg["To"] = sender
    msg.set_content("If you receive this, SMTP_SSL works.")

    srv.send_message(msg)
    print("[+] send_message() SUCCESS")
    srv.quit()
    print("[+] QUIT — Test 2 PASSED\n")
except Exception as exc:
    import traceback
    print(f"\n[-] Test 2 FAILED: {exc}")
    traceback.print_exc()
    print()

print("Done. Check results above to see which method works.")
