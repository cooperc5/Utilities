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

## Usage

Navigate to the directory:

cd WordPlaceholderReplacer

(Optional) Allow script execution if needed:

    Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass

    Open PowerShell and run:

.\ReplaceAndBackup.ps1
