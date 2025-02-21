# Word Placeholder Replacer

## Overview
This PowerShell script replaces placeholders in Word documents. It:
- Detects and replaces `**text**` with just `text` in **bold**.
- Creates a copy of the document before modifying it (`_cleanedUp` version).
- Works on Windows 10+ with Microsoft Word installed.

## Prerequisites
- Windows 10 or later
- Microsoft Word installed
- PowerShell (pre-installed on Windows)

## Installation
1. **Clone the repository**:
   ```sh
   git clone https://github.com/your-username/WordPlaceholderReplacer.git
   
## Steps to Set It Up

    Install AutoHotkey (if you haven't already).
    Create a new AHK script:
        Open Notepad.
        Copy-paste the script below.
        Save it as "QuickSnip.ahk" (with .ahk extension).
    Run the script (double-click the .ahk file).
	
## How It Works

    Press Alt + Shift + S → Opens the built-in Snipping Tool.
    You select a region → Screenshot is copied to the clipboard.
    A confirmation message pops up to let you know it worked.
    Now, you can simply paste (Ctrl + V) into any application!
	
## Make It Run on Startup

To always have this script active:

    Press Win + R, type shell:startup, and hit Enter.
    Copy the QuickSnip.ahk file into the folder.
