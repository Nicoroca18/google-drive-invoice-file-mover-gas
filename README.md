# Google Drive Invoice File Mover

A Google Apps Script that automatically moves invoice PDF files from a temporary Google Drive folder to destination folders defined in a Google Sheet, and writes the file link back to the sheet.

Designed to work with Shared Drives using the Google Drive Advanced API.

---

## Overview

This script helps automate invoice file management by:

- Matching invoice numbers with PDF files in a temporary folder
- Moving files to their corresponding destination folders
- Writing the Google Drive file link back to the spreadsheet
- Skipping already processed rows safely

---

## Features

- Supports Shared Drives
- Efficient file indexing (single Drive scan)
- Idempotent execution (safe to re-run)
- Clear processing summary logs
- Robust error handling

---

## Spreadsheet Structure

### Invoice Sheet

| Column | Description |
|------|------------|
| A | Invoice number |
| M | Destination folder ID |
| P | Google Drive file link (output) |

Rows with an existing file link are automatically skipped.

---

## File Naming Convention

Invoice files in the temporary folder must follow this format:

