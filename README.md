# Exchange Certificate Request Generator

The **Exchange Certificate Request Generator** is a Windows PowerShell GUI tool designed to simplify the management of SSL/TLS certificates for Microsoft Exchange environments. The script provides a user-friendly interface for all major certificate lifecycle operations, including generating requests, completing requests, importing/exporting, and assigning certificates to Exchange services.

## Features

- **Connect to Exchange**: Securely connect to an on-premises Exchange Management Shell using your credentials.
- **Generate Certificate Requests (CSRs)**: Create new certificate signing requests, supporting both SAN and wildcard certificates.
- **Complete Certificate Requests**: Import completed certificates and finish pending requests.
- **Assign Certificates**: Easily assign certificates to Exchange services such as IIS, SMTP, IMAP, and POP.
- **Export Certificates**: Export installed certificates in PFX format for backup or migration, with password protection.
- **Import Certificates**: Import PFX certificates to one or more Exchange servers and assign them to services in a single workflow.
- **Modern, Intuitive UI**: Step-by-step tabbed interface, field validation, and real-time feedback for each operation.

## Usage

1. **Connect** to your Exchange server with valid credentials.
2. **Generate a CSR** or **complete a pending request** as needed.
3. **Import** or **export** certificates in PFX format.
4. **Assign** certificates to Exchange services with a single click.

> **Note:** Requires appropriate Exchange administrative permissions and network access to the Exchange Management Shell.

## Requirements

- Windows PowerShell 5.1 or later
- Exchange Management Shell (Remote PowerShell)
- Permissions to manage certificates on Exchange servers

## Disclaimer

This tool is provided as-is and is intended for use by experienced Exchange administrators. Always test in a non-production environment before deploying in production.
