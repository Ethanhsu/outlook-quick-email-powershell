# Quick Email

Outlook email composer with pre-filled subject line for SPL ticket submission.

## Setup

1. Download or clone this repo
2. Edit `quick-email.ps1` line 8, change the recipient:
   ```powershell
   $fixedRecipient = "your.actual@email.com"
   ```
3. Double-click `Quick Email.vbs` to launch

## First Run

The first time you double-click `Quick Email.vbs`, it creates a Desktop shortcut automatically. Future runs just use the shortcut.

## What It Does

- Enter the SPL ticket ID (e.g. `14-41-13.00-UG-U00-STD-HEL-04/84`)
- Pick a date from the dropdown (1st and 5th of each month)
- Click **Create Email** — Outlook opens with pre-filled To and Subject:

```
To:      your.actual@email.com
Subject: [Power Automate Admin] Add SPL entry 14-41-13.00-UG-U00-STD-HEL-04/84<::>2026-05-05
```

## File Structure

```
outlook-quick-email-powershell/
├── quick-email.ps1     # Main script (edit recipient here)
├── Quick Email.vbs     # Launcher (double-click to run)
├── create-shortcut.bat # Creates Desktop shortcut
└── icon.ico            # App icon
```

## Requirements

- Windows with Outlook installed
- PowerShell (built into Windows)
- No admin rights needed