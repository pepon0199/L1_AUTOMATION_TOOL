import win32com.client
import pythoncom
import time
import threading
import tkinter as tk
from tkinter import messagebox
import pygame
import os
import sys
import ctypes
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options

# === CONFIGURATION ===
CHECK_INTERVAL = 10
SHARED_MAILBOX = "shared_mailbox@example.com"
ALARM_FILENAME = "alarm.wav"  # or .mp3
CHROMEDRIVER_PATH = "chromedriver.exe"

# === GLOBAL FLAGS ===
monitoring = False
monitor_thread = None
alarm_playing = False
wa_monitoring_enabled = True
email_monitoring_enabled = True

# === Prevent Sleep ===
ES_CONTINUOUS = 0x80000000
ES_SYSTEM_REQUIRED = 0x00000001
ES_DISPLAY_REQUIRED = 0x00000002

def prevent_sleep():
    ctypes.windll.kernel32.SetThreadExecutionState(
        ES_CONTINUOUS | ES_SYSTEM_REQUIRED | ES_DISPLAY_REQUIRED)

def allow_sleep():
    ctypes.windll.kernel32.SetThreadExecutionState(ES_CONTINUOUS)

# === Alarm Playback ===
base_path = getattr(sys, '_MEIPASS', os.path.abspath("."))
ALARM_FILE = os.path.join(base_path, ALARM_FILENAME)

def play_alarm_loop():
    try:
        pygame.mixer.init()
        while alarm_playing:
            pygame.mixer.music.load(ALARM_FILE)
            pygame.mixer.music.play()
            while pygame.mixer.music.get_busy() and alarm_playing:
                time.sleep(1)
    except Exception as e:
        append_log(f"❌ Alarm error: {e}")

# === Outlook Monitoring ===
def monitor_emails():
    global monitoring, alarm_playing
    pythoncom.CoInitialize()

    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")

        inbox = None
        for account in namespace.Folders:
            if account.Name.lower() == SHARED_MAILBOX.lower():
                inbox = account.Folders["Inbox"]
                break

        if inbox is None:
            update_status(f"❌ Could not find Inbox for '{SHARED_MAILBOX}'")
            messagebox.showerror("Mailbox Error", f"Shared mailbox '{SHARED_MAILBOX}' not found.")
            monitoring = False
            return

        update_status("Monitoring started...")

        while monitoring:
            try:
                if not email_monitoring_enabled:
                    time.sleep(CHECK_INTERVAL)
                    continue

                unread_items = inbox.Items.Restrict("[Unread]=true")
                unread_count = unread_items.Count

                log_msg = f"[{time.strftime('%H:%M:%S')}] Unread emails: {unread_count}"
                append_log(log_msg)
                update_status(log_msg)

                if unread_count > 0 and not alarm_playing:
                    alarm_playing = True
                    threading.Thread(target=play_alarm_loop, daemon=True).start()

            except Exception as check_err:
                append_log(f"⚠️ Email error: {check_err}")

            time.sleep(CHECK_INTERVAL)

    except Exception as e:
        update_status("Monitoring stopped due to error.")
        messagebox.showerror("Monitoring Error", f"❌ Email Monitor Error:\n{e}")
        monitoring = False

    finally:
        pythoncom.CoUninitialize()

# === WhatsApp Monitoring ===
def monitor_whatsapp():
    global alarm_playing

    try:
        options = Options()
        profile_path = os.path.join(os.environ["TEMP"], "wa_profile_temp")
        options.add_argument(f"user-data-dir={profile_path}")
        options.add_argument("--disable-dev-shm-usage")
        options.add_argument("--no-sandbox")
        options.add_argument("--remote-debugging-port=9222")
        options.add_experimental_option("excludeSwitches", ["enable-logging"])

        driver = webdriver.Chrome(service=Service(CHROMEDRIVER_PATH), options=options)
        driver.get("https://web.whatsapp.com")

        append_log("🔗 Waiting for WhatsApp Web login...")
        time.sleep(15)

        append_log("✅ WhatsApp monitoring started...")

        while monitoring:
            try:
                if not wa_monitoring_enabled:
                    time.sleep(5)
                    continue

                unread = driver.find_elements("xpath", "//span[contains(@aria-label, 'unread message')]")
                append_log(f"🧪 Found {len(unread)} unread badge(s) on WhatsApp.")

                if unread:
                    append_log("📲 WhatsApp: Unread message(s) detected!")

                    if not alarm_playing:
                        alarm_playing = True
                        threading.Thread(target=play_alarm_loop, daemon=True).start()

                time.sleep(5)

            except Exception as inner:
                append_log(f"⚠️ WhatsApp error: {inner}")
                time.sleep(5)

    except Exception as e:
        append_log(f"❌ WhatsApp Monitor Error: {e}")

# === GUI Functions ===
def update_status(msg):
    status_var.set(f"Status: {msg}")

def append_log(msg):
    log_box.insert(tk.END, msg + "\n")
    log_box.see(tk.END)

def start_monitoring():
    global monitoring
    if not monitoring:
        monitoring = True
        prevent_sleep()
        append_log(">> Starting monitor thread...")

        if email_toggle_var.get():
            threading.Thread(target=monitor_emails, daemon=True).start()

        if whatsapp_toggle_var.get():
            threading.Thread(target=monitor_whatsapp, daemon=True).start()

    else:
        messagebox.showinfo("Info", "Monitoring already running.")

def stop_monitoring():
    global monitoring, alarm_playing
    monitoring = False
    alarm_playing = False
    allow_sleep()
    pygame.mixer.music.stop()
    update_status("Monitoring stopped.")
    append_log(">> Monitoring stopped.")

# === GUI Layout ===
root = tk.Tk()
root.title("SupportCompanion")
root.geometry("520x400")

status_var = tk.StringVar()
status_var.set("Status: Not monitoring")

status_label = tk.Label(root, textvariable=status_var, wraplength=480)
status_label.pack(pady=10)

# Monitoring Toggles
email_toggle_var = tk.BooleanVar(value=True)
whatsapp_toggle_var = tk.BooleanVar(value=True)

toggle_frame = tk.Frame(root)
toggle_frame.pack(pady=5)

tk.Checkbutton(toggle_frame, text="Monitor Outlook", variable=email_toggle_var).grid(row=0, column=0, padx=10)
tk.Checkbutton(toggle_frame, text="Monitor WhatsApp", variable=whatsapp_toggle_var).grid(row=0, column=1, padx=10)

btn_frame = tk.Frame(root)
btn_frame.pack()

tk.Button(btn_frame, text="Start Monitoring", width=20, command=start_monitoring).grid(row=0, column=0, padx=5)
tk.Button(btn_frame, text="Stop Monitoring", width=20, command=stop_monitoring).grid(row=0, column=1, padx=5)

log_box = tk.Text(root, height=10, width=64)
log_box.pack(pady=10)

tk.Button(root, text="Exit", width=20, command=root.quit).pack(pady=5)

root.mainloop()
