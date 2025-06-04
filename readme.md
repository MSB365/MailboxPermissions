# Exchange Mailbox Permission Script

---

This PowerShell script automates the management of "Send As" and "Full Access" permissions for Exchange mailboxes. It reads desired permissions from a CSV file, applies them, and removes any existing permissions not specified in the CSV. The script includes a graphical user interface (GUI) file picker for easy selection of your CSV input.

*Use PowerShell 7 to run this script.*

## Documentation
The documentation of this script can be found at the Blogpost <https://www.msb365.blog/?p=5907>

## Features

* **Automated Permission Management**: Efficiently set "Send As" and "Full Access" permissions.
* **Permission Synchronization**: Removes outdated or unauthorized permissions not listed in your CSV.
* **User-Friendly Interface**: A GUI file picker simplifies CSV file selection.

---

## CSV File Format

The CSV file should have the following headers and contain one row per permission entry:

| Header            | Description                                                     |
| :---------------- | :-------------------------------------------------------------- |
| `MailboxIdentity` | The email address or alias of the mailbox to modify.            |
| `UserIdentity`    | The email address or alias of the user to grant permissions to. |
| `SendAs`          | Set to `TRUE` to grant "Send As" permission, `FALSE` otherwise. |
| `FullAccess`      | Set to `TRUE` to grant "Full Access" permission, `FALSE` otherwise. |

### Example CSV

```csv
MailboxIdentity,UserIdentity,SendAs,FullAccess
shared.mailbox@company.com,john.doe@company.com,TRUE,TRUE
shared.mailbox@company.com,jane.smith@company.com,TRUE,FALSE
shared.mailbox@company.com,admin@company.com,FALSE,TRUE
finance@company.com,john.doe@company.com,TRUE,TRUE
finance@company.com,finance.manager@company.com,TRUE,TRUE
