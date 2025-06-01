
# Exchange Certificate Request Generator

A Windows Forms-based PowerShell GUI tool to simplify the process of managing Exchange Server certificates. This tool allows administrators to connect to an Exchange environment and perform certificate-related operations such as generating CSRs, completing requests, assigning certificates to services, and importing/exporting certificates.

## Features

- 🔐 **Connect to Exchange**: Authenticate and establish a remote PowerShell session with an Exchange server.
- 📄 **Generate Certificate Signing Requests (CSR)**:
  - Supports SAN and Wildcard certificates.
  - Input friendly name, subject name, and SANs.
  - Save `.req` files for submission to a Certificate Authority.
- ✅ **Complete Certificate Requests**:
  - Import `.p7b` certificate chains.
  - Automatically extract and store the thumbprint.
- 🔧 **Assign Certificates to Services**:
  - Assign imported certificates to IIS, SMTP, IMAP, and POP.
  - Optionally force overwrite existing assignments.
- 📤 **Export Certificates**:
  - Export installed certificates as `.pfx` files.
  - Secure with password protection.
- 📥 **Import Certificates**:
  - Import `.pfx` files to one or more Exchange servers.
  - Assign to selected services during import.

## Requirements

- PowerShell 5.1+
- Exchange Management Shell access
- Windows OS with .NET Framework (for WinForms support)

## Usage

1. Launch the script in PowerShell.
2. Enter Exchange credentials and server address.
3. Choose an operation mode:
   - **New Certificate Request**
   - **Complete Certificate Request**
4. Follow the form inputs and click the appropriate action buttons.
5. Use the Export/Import buttons for certificate management.

## Notes

- Ensure you have the necessary permissions to manage certificates on the Exchange server.
- The tool uses Kerberos authentication and assumes domain connectivity.

## License

MIT License
