# Windows Folder Print Automation

## Overview

This script automates printing of files placed in a designated folder on Windows systems.
It is designed to run periodically (e.g., every minute) via Windows Task Scheduler.

The script processes incoming files, sends them to the default printer based on file type, and archives them after processing.

---

## Features

* Folder-based print automation
* Supports multiple file types:

  * PDF
  * Images (JPG, PNG, BMP, GIF)
  * Text files
  * Microsoft Word (DOCX)
  * Microsoft Excel (XLSX)
* Automatic file stability check before processing
* Temporary staging to avoid partial reads
* Logging with daily rotation
* Archive of processed files (timestamp-based)
* Basic locking mechanism to prevent duplicate processing

---

## How It Works

1. The script scans the input folder for new files
2. Each file is checked to ensure it is fully written (stable size)
3. Files are moved to a temporary processing directory
4. The appropriate application is used to send the file to the default printer
5. After processing, files are archived with a timestamp

The script is intended to be executed repeatedly via Task Scheduler (e.g., every minute).

---

## Requirements

* Windows (Client or Server)
* Python 3.x
* Installed applications depending on file types:

  * PDF viewer (e.g., PDF-XChange Editor)
  * Image viewer (e.g., IrfanView)
  * Microsoft Word (for DOCX)
  * Microsoft Excel (for XLSX)

Administrative privileges may be required depending on the printer configuration.

---

## Configuration

Key parameters can be adjusted in the script:

* Input folder
* Log directory
* Temporary directory
* Archive directory
* External application paths (PDF viewer, image viewer, Office)

Example (simplified):

```python
ROOT_PATH = r"C:\print\input"
LOG_PATH = r"C:\print\logs"
```

Ensure all required applications are installed and paths are valid.

---

## Scheduling (Task Scheduler)

Recommended setup:

* Trigger: every 1 minute
* Action: run Python script
* Run whether user is logged in or not
* Use highest privileges (if required for printing)

---

## Limitations

* Uses the system's default printer (no per-job printer selection)
* Depends on locally installed applications for printing
* Office automation (Word/Excel) may be unreliable in some server environments
* No support for unsupported file types (these are skipped)
* No real-time monitoring (scheduled execution only)

---

## Logging

Logs are written to the configured log directory and rotated daily.
Processed files are archived with timestamps for traceability.

---

## Use Case

This script is intended for internal automation scenarios such as:

* Shared folder → automatic printing
* Back-office document workflows
* Simple print queue automation without dedicated print server software

---

## Disclaimer

This script is provided as-is and is intended for controlled environments.
Thorough testing is recommended before production use.
