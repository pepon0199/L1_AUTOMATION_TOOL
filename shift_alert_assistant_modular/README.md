# SupportCompanionModular

This folder contains the `SupportCompanionModular.py` script and the required `chromedriver.exe` for WhatsApp monitoring.

## Description

`SupportCompanionModular.py` is a Python application that monitors Outlook emails and WhatsApp messages, playing an alarm when unread messages are detected. It uses Selenium with ChromeDriver for WhatsApp web automation.

**Difference from non-modular SupportCompanion:**
The modular version separates concerns better by organizing code into classes and threads, improving maintainability and scalability compared to the original monolithic script.

## Setup

1. Ensure you have Python 3.x installed.
2. Install the required Python packages listed in `requirements.txt`.
3. Run the script from this folder:

```bash
python SupportCompanionModular.py
```

## Building Executable (.exe)

To create a standalone executable using PyInstaller, use the following command:

```bash
python -m PyInstaller --clean --onefile --windowed --hidden-import=win32timezone --add-data "alarm.wav;." --name "SupportCompanionModular" SupportCompanionModular.py
```

This will generate a single executable named `SupportCompanionModular.exe`.

## Files

- `SupportCompanionModular.py`: Main monitoring script.
- `chromedriver.exe`: ChromeDriver executable for Selenium.

## Notes

- The script uses a temporary Chrome user profile stored in the system TEMP directory.
- Make sure to log in to WhatsApp Web when prompted on the first run.
