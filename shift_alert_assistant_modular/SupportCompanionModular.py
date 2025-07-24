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
SHARED_MAILBOX = "LE-HELPDESK.PH@fpt.com"
ALARM_FILENAME = "alarm.wav"  # or .mp3
CHROMEDRIVER_PATH = "chromedriver.exe"

# === GLOBAL FLAGS ===
monitoring = False
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

class AlarmPlayer:
    def __init__(self):
        pygame.mixer.init()
        self.alarm_playing = False
        self.lock = threading.Lock()

    def play_alarm_loop(self):
        with self.lock:
            if self.alarm_playing:
                return  # Already playing
            self.alarm_playing = True

        try:
            pygame.mixer.music.load(ALARM_FILE)
            pygame.mixer.music.play(-1)  # Loop indefinitely
            while self.alarm_playing:
                time.sleep(1)
            pygame.mixer.music.stop()
        except Exception as e:
            print(f"❌ Alarm error: {e}")
        finally:
            with self.lock:
                self.alarm_playing = False

    def stop_alarm(self):
        with self.lock:
            self.alarm_playing = False
        pygame.mixer.music.stop()

# === Outlook Monitoring ===
class EmailMonitor(threading.Thread):
    def __init__(self, alarm_player, update_status_callback, append_log_callback):
        super().__init__(daemon=True)
        self.alarm_player = alarm_player
        self.update_status = update_status_callback
        self.append_log = append_log_callback
        self.monitoring = False

    def run(self):
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
                self.update_status(f"❌ Could not find Inbox for '{SHARED_MAILBOX}'")
                messagebox.showerror("Mailbox Error", f"Shared mailbox '{SHARED_MAILBOX}' not found.")
                return

            self.update_status("Email monitoring started...")
            self.monitoring = True

            while self.monitoring:
                if not email_monitoring_enabled:
                    time.sleep(CHECK_INTERVAL)
                    continue

                try:
                    unread_items = inbox.Items.Restrict("[Unread]=true")
                    unread_count = unread_items.Count

                    log_msg = f"[{time.strftime('%H:%M:%S')}] Unread emails: {unread_count}"
                    self.append_log(log_msg)
                    self.update_status(log_msg)

                    if unread_count > 0 and not self.alarm_player.alarm_playing:
                        threading.Thread(target=self.alarm_player.play_alarm_loop, daemon=True).start()

                except Exception as check_err:
                    self.append_log(f"⚠️ Email error: {check_err}")

                time.sleep(CHECK_INTERVAL)

        except Exception as e:
            self.update_status("Email monitoring stopped due to error.")
            messagebox.showerror("Monitoring Error", f"❌ Email Monitor Error:\n{e}")

        finally:
            pythoncom.CoUninitialize()

    def stop(self):
        self.monitoring = False

# === WhatsApp Monitoring ===
class WhatsAppMonitor(threading.Thread):
    def __init__(self, alarm_player, update_status_callback, append_log_callback):
        super().__init__(daemon=True)
        self.alarm_player = alarm_player
        self.update_status = update_status_callback
        self.append_log = append_log_callback
        self.monitoring = False
        self.driver = None

    def run(self):
        try:
            options = Options()
            profile_path = os.path.join(os.environ["TEMP"], "wa_profile_temp")
            options.add_argument(f"user-data-dir={profile_path}")
            options.add_argument("--disable-dev-shm-usage")
            options.add_argument("--no-sandbox")
            options.add_argument("--remote-debugging-port=9222")
            options.add_experimental_option("excludeSwitches", ["enable-logging"])

            self.driver = webdriver.Chrome(service=Service(CHROMEDRIVER_PATH), options=options)
            self.driver.get("https://web.whatsapp.com")

            self.append_log("🔗 Waiting for WhatsApp Web login...")
            time.sleep(15)

            self.update_status("WhatsApp monitoring started...")
            self.monitoring = True

            while self.monitoring:
                if not wa_monitoring_enabled:
                    time.sleep(5)
                    continue

                try:
                    unread = self.driver.find_elements("xpath", "//span[contains(@aria-label, 'unread message')]")
                    self.append_log(f"🧪 Found {len(unread)} unread badge(s) on WhatsApp.")

                    if unread:
                        self.append_log("📲 WhatsApp: Unread message(s) detected!")

                        if not self.alarm_player.alarm_playing:
                            threading.Thread(target=self.alarm_player.play_alarm_loop, daemon=True).start()

                    time.sleep(5)

                except Exception as inner:
                    self.append_log(f"⚠️ WhatsApp error: {inner}")
                    time.sleep(5)

        except Exception as e:
            self.append_log(f"❌ WhatsApp Monitor Error: {e}")

    def stop(self):
        self.monitoring = False
        if self.driver:
            self.driver.quit()

# === GUI Functions and Layout ===
class SupportCompanionApp:
    def __init__(self, root):
        self.root = root
        self.root.title("SupportCompanion Modular")
        self.root.geometry("520x400")

        self.status_var = tk.StringVar()
        self.status_var.set("Status: Not monitoring")

        self.status_label = tk.Label(root, textvariable=self.status_var, wraplength=480)
        self.status_label.pack(pady=10)

        # Monitoring Toggles
        self.email_toggle_var = tk.BooleanVar(value=True)
        self.whatsapp_toggle_var = tk.BooleanVar(value=True)

        toggle_frame = tk.Frame(root)
        toggle_frame.pack(pady=5)

        tk.Checkbutton(toggle_frame, text="Monitor Outlook", variable=self.email_toggle_var).grid(row=0, column=0, padx=10)
        tk.Checkbutton(toggle_frame, text="Monitor WhatsApp", variable=self.whatsapp_toggle_var).grid(row=0, column=1, padx=10)

        btn_frame = tk.Frame(root)
        btn_frame.pack()

        self.start_button = tk.Button(btn_frame, text="Start Monitoring", width=20, command=self.start_monitoring)
        self.start_button.grid(row=0, column=0, padx=5)

        self.stop_button = tk.Button(btn_frame, text="Stop Monitoring", width=20, command=self.stop_monitoring, state=tk.DISABLED)
        self.stop_button.grid(row=0, column=1, padx=5)

        self.log_box = tk.Text(root, height=10, width=64)
        self.log_box.pack(pady=10)

        tk.Button(root, text="Exit", width=20, command=root.quit).pack(pady=5)

        self.alarm_player = AlarmPlayer()
        self.email_monitor = None
        self.whatsapp_monitor = None

    def update_status(self, msg):
        self.status_var.set(f"Status: {msg}")

    def append_log(self, msg):
        self.log_box.insert(tk.END, msg + "\n")
        self.log_box.see(tk.END)

    def start_monitoring(self):
        if self.email_monitor or self.whatsapp_monitor:
            messagebox.showinfo("Info", "Monitoring already running.")
            return

        global email_monitoring_enabled, wa_monitoring_enabled
        email_monitoring_enabled = self.email_toggle_var.get()
        wa_monitoring_enabled = self.whatsapp_toggle_var.get()

        self.update_status("Starting monitoring...")
        self.append_log(">> Starting monitor threads...")

        if email_monitoring_enabled:
            self.email_monitor = EmailMonitor(self.alarm_player, self.update_status, self.append_log)
            self.email_monitor.start()

        if wa_monitoring_enabled:
            self.whatsapp_monitor = WhatsAppMonitor(self.alarm_player, self.update_status, self.append_log)
            self.whatsapp_monitor.start()

        prevent_sleep()
        self.start_button.config(state=tk.DISABLED)
        self.stop_button.config(state=tk.NORMAL)

    def stop_monitoring(self):
        global monitoring
        monitoring = False
        self.alarm_player.stop_alarm()

        if self.email_monitor:
            self.email_monitor.stop()
            self.email_monitor = None

        if self.whatsapp_monitor:
            self.whatsapp_monitor.stop()
            self.whatsapp_monitor = None

        allow_sleep()
        self.update_status("Monitoring stopped.")
        self.append_log(">> Monitoring stopped.")
        self.start_button.config(state=tk.NORMAL)
        self.stop_button.config(state=tk.DISABLED)

if __name__ == "__main__":
    root = tk.Tk()
    app = SupportCompanionApp(root)
    root.mainloop()
