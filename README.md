# Export Exchange Mailbox Delegation Permissions

## Overview
This PowerShell script connects to Exchange Online and retrieves **Full Access, Send As, and Send on Behalf** permissions for all **user and shared mailboxes**. The results are exported into a CSV file.

## Features
- Connects to Exchange Online securely.
- Retrieves all user and shared mailboxes.
- Extracts delegation permissions (Full Access, SendAs, and SendOnBehalf).
- Saves the data into a CSV file for further analysis.
- Ensures a clean disconnection from Exchange Online.

## Prerequisites
- **Exchange Online PowerShell Module** must be installed.
- You must have appropriate **Exchange Administrator** or **Global Administrator** privileges.
- **PowerShell 7 or later** is recommended.

## How to Use

### **1. Install Exchange Online PowerShell Module**
Before running the script, ensure that the **Exchange Online PowerShell Module** is installed:
```powershell
Install-Module ExchangeOnlineManagement -Force
