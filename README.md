# Export Exchange Mailbox Delegation Permissions

## Overview
This PowerShell script helps **Exchange administrators** retrieve and export **Full Access, Send As, and Send on Behalf** permissions for all **user and shared mailboxes** in Exchange Online. The data is saved in a CSV file, making it easy to review and audit mailbox delegation settings.

This script is useful for IT teams managing Exchange Online as it provides a clear overview of who has access to which mailboxes, ensuring security and compliance.

## Key Features
- ✅ **Automates the process** of retrieving mailbox permissions.
- ✅ **Exports data** to a CSV file (`C:\DelegationAccessSummary.csv`).
- ✅ **Supports both user and shared mailboxes**.
- ✅ **Uses a secure connection** to Exchange Online.
- ✅ **Disconnects automatically** after execution to avoid unnecessary sessions.

## Why Use This Script?
🔹 **Simplifies administration** by providing a quick summary of mailbox delegation.  
🔹 **Helps IT teams** efficiently manage mailbox access in Exchange Online.  
🔹 **Supports security audits** and ensures compliance with organizational policies.  
🔹 **Reduces manual effort** and human error in permission tracking.  

## How It Works
1. The script **connects securely** to Exchange Online.
2. It retrieves all **user and shared mailboxes**.
3. The script checks and records:
   - **Full Access permissions** (who can fully control the mailbox)
   - **Send As permissions** (who can send emails as the mailbox owner)
   - **Send on Behalf permissions** (who can send emails on behalf of the mailbox owner)
4. The collected data is **saved into a structured CSV file**.
5. The script **disconnects** from Exchange Online to close the session securely.

## Prerequisites
Before running this script, make sure you have the following:

### **System Requirements**
| Requirement | Description |
|------------|-------------|
| **PowerShell 7 or later** | Recommended for best performance |
| **Exchange Online Management Module** | Must be installed to connect to Exchange Online |
| **Exchange Admin or Global Admin Access** | Required to retrieve mailbox permissions |

## Installation Guide

### **Step 1: Install the Exchange Online PowerShell Module**
Before running the script, install the **Exchange Online Management Module** by running the following command in PowerShell:

```powershell
Install-Module ExchangeOnlineManagement -Force
```

### **Step 2: Run the Script**
Save the script as `Export-ExchangeDelegation.ps1` and **run it in PowerShell as an administrator**:

```powershell
.\Export-ExchangeDelegation.ps1
```

### **Step 3: Check the Output**
Once the script finishes running, you will find a CSV file in the following location:

📂 `C:\DelegationAccessSummary.csv`

Open this file in **Excel** or any spreadsheet application to review the delegation settings.

## Example Output (CSV Format)
| Mailbox                  | User                | Access Type   |
|--------------------------|---------------------|--------------|
| user1@company.com        | admin@company.com   | FullAccess  |
| user2@company.com        | user1@company.com   | SendAs      |
| shared@company.com       | user3@company.com   | SendOnBehalf |

## Troubleshooting

🔹 **Connection Issues** – Ensure your network allows **PowerShell remoting** and that your credentials are correct.  
🔹 **No Output in CSV** – Verify that your account has permission to query Exchange Online mailboxes.  
🔹 **Access Denied Errors** – Run the script with an account that has **Exchange Admin** or **Global Admin** rights.  

## License
This project is licensed under the **MIT License** – You are free to use and modify it as needed.

---
📌 **Perfect for Exchange administrators who need quick and easy visibility into mailbox delegation settings!** 🚀

