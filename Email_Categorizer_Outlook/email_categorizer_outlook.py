import win32com.client
import pythoncom
import time
import re
import threading
import tkinter as tk
from tkinter import scrolledtext, messagebox
from rapidfuzz import fuzz
import pygame
import os
import sys
from win10toast import ToastNotifier
from pathlib import Path
import logging

class EmailCategorizerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Email Categorizer v3")
        self.root.geometry("520x450")

        # Setup logging to file and GUI
        self.setup_logging()

        # Notification toaster
        self.toaster = ToastNotifier()

        # Configurable parameters
        self.shared_mailbox = "LE-HELPDESK.PH@fpt.com"
        self.primary_keywords = [
            "change", "update", "migrate", "migration", "modify", "mobile", "number", "phone",
            "unlock", "reset", "enroll", "new", "locked"
        ]
        self.secondary_keywords = ["ddt"]
        self.excluded_keywords = ["UCUBE"]
        self.excluded_senders = ["LE-HELPDESK.PH"]
        self.running = False
        self.selected_category = None
        self.monitor_thread = None

        # Fetch categories dynamically from Outlook
        self.allowed_categories = self.fetch_outlook_categories()

        # Setup GUI components
        self.create_widgets()

        # Initialize pygame mixer for sound
        pygame.mixer.init()

    def setup_logging(self):
        self.logger = logging.getLogger("EmailCategorizer")
        self.logger.setLevel(logging.DEBUG)
        formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')

        # File handler
        fh = logging.FileHandler("email_categorizer.log")
        fh.setLevel(logging.DEBUG)
        fh.setFormatter(formatter)
        self.logger.addHandler(fh)

        # GUI log handler
        self.gui_log_handler = GuiLogHandler(self)
        self.logger.addHandler(self.gui_log_handler)

    def create_widgets(self):
        # Frame for category buttons
        self.button_frame = tk.Frame(self.root)
        self.button_frame.pack(pady=10)

        self.category_buttons = {}
        # Dynamically create a grid layout for buttons based on number of categories
        categories = list(self.allowed_categories.keys())
        num_categories = len(categories)
        cols = 2
        rows = (num_categories + 1) // cols

        for idx, category in enumerate(categories):
            row = idx // cols
            col = idx % cols
            color = self.allowed_categories.get(category, "lightgray")
            btn = tk.Button(
                self.button_frame, text=category,
                command=lambda c=category: self.start_monitoring(c),
                font=("Arial", 12), bg=color, fg="black", width=12
            )
            btn.grid(row=row, column=col, padx=10, pady=5)
            self.category_buttons[category] = btn

        # Stop button
        self.stop_button = tk.Button(
            self.root, text="STOP", command=self.stop_monitoring,
            font=("Arial", 12), bg="red", fg="white", width=26
        )
        self.stop_button.pack(pady=5)

        # Status label
        self.status_var = tk.StringVar(value="Status: Stopped")
        self.status_label = tk.Label(self.root, textvariable=self.status_var, font=("Arial", 10, "italic"))
        self.status_label.pack(pady=2)

        # Log box
        self.log_box = scrolledtext.ScrolledText(self.root, width=65, height=15, font=("Arial", 10))
        self.log_box.pack(padx=10, pady=10)

    def log_message(self, message):
        self.logger.info(message)

    def start_monitoring(self, category):
        if self.running:
            messagebox.showinfo("Info", "Monitoring is already running. Please stop it first.")
            return

        self.selected_category = category
        self.running = True
        self.status_var.set(f"Status: Monitoring for '{category}'")
        self.log_message(f"✅ Monitoring started for '{category}'")

        # Disable category buttons and enable stop button
        self.set_buttons_state(tk.DISABLED)
        self.stop_button.config(state=tk.NORMAL)

        self.monitor_thread = threading.Thread(target=self.monitor_emails, daemon=True)
        self.monitor_thread.start()

    def stop_monitoring(self):
        if not self.running:
            return
        self.running = False
        self.selected_category = None
        self.status_var.set("Status: Stopped")
        self.log_message("🛑 Monitoring stopped.")

        # Enable category buttons and disable stop button
        self.set_buttons_state(tk.NORMAL)
        self.stop_button.config(state=tk.DISABLED)

    def set_buttons_state(self, state):
        for btn in self.category_buttons.values():
            btn.config(state=state)

    def monitor_emails(self):
        try:
            self.log_message(f"🔄 Connecting to Outlook... Monitoring for '{self.selected_category}'")
            pythoncom.CoInitialize()
            outlook = win32com.client.Dispatch("Outlook.Application")
            namespace = outlook.GetNamespace("MAPI")

            inbox = None
            for account in namespace.Folders:
                if account.Name == self.shared_mailbox:
                    inbox = account.Folders["Inbox"]
                    break

            if inbox is None:
                self.log_message(f"❌ Unable to find inbox for {self.shared_mailbox}")
                self.stop_monitoring()
                return

            # Fetch categories from Outlook session (master categories)
            self.outlook_categories = self.get_outlook_master_categories(namespace)

            self.log_message(f"🚀 Monitoring emails... Assigning category: {self.selected_category}")

            while self.running:
                self.log_message("📩 Checking unread emails...")

                messages = inbox.Items.Restrict("[Unread] = true")

                for message in messages:
                    try:
                        self.process_message(message)
                    except Exception as e:
                        self.log_message(f"❌ Error processing email: {str(e)}")

                self.log_message("🔄 Checking again in 3 seconds...")
                time.sleep(3)

        except Exception as e:
            self.log_message(f"❌ Unexpected error: {str(e)}")

        finally:
            pythoncom.CoUninitialize()
            self.stop_monitoring()

    def process_message(self, message):
        subject = (message.Subject or "").lower()
        sender = (message.SenderName or "").lower()
        existing_categories = message.Categories.split("; ") if message.Categories else []

        # Exclusions
        if any(excluded.lower() in sender for excluded in self.excluded_senders):
            self.log_message(f"🚫 Ignoring email from {message.SenderName}: {message.Subject}")
            return

        if any(excluded.lower() in subject for excluded in self.excluded_keywords):
            self.log_message(f"🚫 Ignoring email (excluded keyword): {message.Subject}")
            return

        regex_patterns = [
            r"\\bchange.*number\\b",
            r"\\bupdate.*number\\b",
            r"\\bmodify.*mobile\\b",
            r"\\breset.*password\\b",
            r"\\bunlock.*user\\b",
            r"\\brequest.*change.*number\\b"
        ]
        regex_matched = any(re.search(pattern, subject) for pattern in regex_patterns)

        # Edge case for DDT/DTT + request
        matched_by_fallback = False
        if ("dtt" in subject or "ddt" in subject) and "request" in subject:
            self.log_message("✨ Edge-case match: Subject contains 'ddt/dtt' and 'request'")
            matched_by_fallback = True

        if "number" in subject and not any(kw in subject for kw in ["change", "update", "modify", "reset", "migrate"]):
            self.log_message(f"🚫 Skipping email (Contains 'number' but lacks relevant action): {message.Subject}")
            return

        body = (message.Body or "").lower()

        if not (
            any(keyword in subject for keyword in self.primary_keywords)
            or regex_matched
            or self.fuzzy_match_keywords(subject, self.primary_keywords)
        ):
            if self.fuzzy_match_keywords(body, self.primary_keywords):
                self.log_message(f"📬 Matched via body content: {message.Subject or '(no subject)'}")
                matched_by_fallback = True
            else:
                matched_by_fallback = False

        if (not subject or subject.strip() == "(no subject)") and not body.strip() and message.Attachments.Count > 0:
            self.log_message(f"📎 Tagging email based on attachment only: {message.Subject or '(no subject)'}")
            matched_by_fallback = True

        attachment_names = []
        if message.Attachments.Count > 0:
            attachment_names = [att.FileName for att in message.Attachments]
            self.log_message(f"📎 Attachments found: {', '.join(attachment_names)}")

        important_attachment_keywords = ["change", "update", "ddt", "dtt", "reset", "unlock", "modify", "migrate"]
        for att_name in attachment_names:
            lower_name = att_name.lower()
            if any(kw in lower_name for kw in important_attachment_keywords):
                self.log_message(f"📎 Categorizing based on attachment name: {att_name}")
                matched_by_fallback = True
                break

        match_sources = []

        if any(keyword in subject for keyword in self.primary_keywords):
            match_sources.append("🔑 primary keyword (subject)")

        if regex_matched:
            match_sources.append("📏 regex match")

        if self.fuzzy_match_keywords(subject, self.primary_keywords):
            match_sources.append("🤏 fuzzy match")

        if matched_by_fallback:
            match_sources.append("🔁 fallback match (body or attachments)")

        if (
            any(keyword in subject for keyword in self.primary_keywords)
            or regex_matched
            or self.fuzzy_match_keywords(subject, self.primary_keywords)
            or matched_by_fallback
            or self.advanced_fuzzy_match(subject, [
                "change mobile number",
                "update contact number",
                "reset user password",
                "unlock user",
                "migrate phone",
                "modify phone number"
            ])
        ):

            if any(sec in subject for sec in self.secondary_keywords) and not any(kw in subject for kw in self.primary_keywords):
                self.log_message(f"🚫 Ignoring email (DDT but no relevant keyword): {message.Subject}")
                return

            if existing_categories and self.selected_category not in existing_categories:
                self.log_message(f"🚫 Skipping email (Already assigned to {', '.join(existing_categories)}): {message.Subject}")
                return

            if self.selected_category in existing_categories:
                self.log_message(f"⚠️ Email already categorized as '{self.selected_category}': {message.Subject}")
                return

            self.log_message(f"🟢 Categorizing: {message.Subject}")
            self.log_message(f"📋 Reason for tagging: {', '.join(match_sources)}")
            existing_categories.append(self.selected_category)
            message.Categories = "; ".join(existing_categories)
            message.Save()


    def fuzzy_match_keywords(self, subject, keywords, threshold=90):
        subject = subject.lower()
        for keyword in keywords:
            score = fuzz.partial_ratio(subject, keyword.lower())
            if score >= threshold:
                return True
        return False

    def advanced_fuzzy_match(self, subject, patterns, threshold=75):
        subject = subject.lower()
        for pattern in patterns:
            if fuzz.partial_ratio(pattern.lower(), subject) >= threshold:
                return True
        return False

    def play_custom_sound(self):
        try:
            base_path = getattr(sys, '_MEIPASS', os.path.abspath("."))
            sound_path = Path(base_path) / "custom_sound.mp3"
            if sound_path.exists():
                pygame.mixer.music.load(str(sound_path))
                pygame.mixer.music.play()
            else:
                self.log_message(f"❌ Sound file not found: {sound_path}")
        except Exception as e:
            self.log_message(f"❌ Error playing sound: {e}")

    def notify_popup(self, subject, category):
        try:
            title = "📌 Email Categorized"
            message = f"Assigned to: {category}\nSubject: {subject or '(no subject)'}"
            self.toaster.show_toast(title, message, duration=5, threaded=True)
        except Exception as e:
            self.log_message(f"⚠️ Toast notification error: {e}")

    def fetch_outlook_categories(self):
        """
        Fetch categories from the Outlook master categories list.
        Returns a dictionary of category name to color (hex or color name).
        """
        categories = {}
        try:
            pythoncom.CoInitialize()
            outlook = win32com.client.Dispatch("Outlook.Application")
            namespace = outlook.GetNamespace("MAPI")
            session_categories = self.get_outlook_master_categories(namespace)
            for cat in session_categories:
                # Outlook category color is an integer, map to color name or hex if needed
                color_name = self.map_outlook_color(cat.Color)
                categories[cat.Name] = color_name
        except Exception as e:
            self.log_message(f"❌ Error fetching Outlook categories: {e}")
        finally:
            pythoncom.CoUninitialize()
        if not categories:
            # Fallback to default categories if none found
            categories = {
                "KARL": "green",
                "ADRIAN": "yellow",
                "Borgz": "gray",
                "JB 👽": "orange"
            }
        return categories

    def get_outlook_master_categories(self, namespace):
        """
        Helper to get the master categories collection from Outlook session.
        """
        try:
            return namespace.Categories
        except Exception as e:
            self.log_message(f"❌ Error accessing Outlook master categories: {e}")
            return []

    def map_outlook_color(self, color_index):
        """
        Map Outlook category color index to a hex color string.
        Outlook color index reference:
        0 = None, 1 = Red, 2 = Orange, 3 = Yellow, 4 = Green, 5 = Blue, 6 = Purple, 7 = Maroon, 8 = Steel Blue, 9 = Dark Green, 10 = Teal, 11 = Olive, 12 = Gray, 13 = Dark Gray, 14 = Black
        """
        color_map = {
            0: "#D3D3D3",  # lightgray
            1: "#FF0000",  # red
            2: "#FFA500",  # orange
            3: "#FFFF00",  # yellow
            4: "#008000",  # green
            5: "#0000FF",  # blue
            6: "#800080",  # purple
            7: "#800000",  # maroon
            8: "#4682B4",  # steelblue
            9: "#006400",  # darkgreen
            10: "#008080", # teal
            11: "#808000", # olive
            12: "#808080", # gray
            13: "#A9A9A9", # darkgray
            14: "#000000"  # black
        }
        return color_map.get(color_index, "#D3D3D3")

class GuiLogHandler(logging.Handler):
    def __init__(self, app):
        super().__init__()
        self.app = app

    def emit(self, record):
        msg = self.format(record)
        def append():
            self.app.log_box.insert(tk.END, msg + "\n")
            self.app.log_box.yview(tk.END)
        self.app.root.after(0, append)


if __name__ == "__main__":
    root = tk.Tk()
    app = EmailCategorizerApp(root)
    root.mainloop()
