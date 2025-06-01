# Exchange Certificate Request GUI

A Windows Forms PowerShell GUI tool for generating and completing certificate requests in Microsoft Exchange environments.

## Features

- **Exchange Connection:**  
  Authenticate to on-premises Exchange Management Shell using user credentials.

- **Operation Modes:**  
  - **New Certificate Request:**  
    - SAN or Wildcard certificate options  
    - Supports subject name, friendly name, and SAN entries  
    - Generates a PEM-formatted CSR file (.req)
  - **Complete Certificate Request:**  
    - Import a certificate response/chain (.p7b) directly to Exchange

- **Safe Session Handling:**  
  - Cleans up existing Exchange sessions before connecting
  - Disconnect and reconnect buttons

- **User-Friendly UI:**  
  - Browse dialogs for file selection  
  - Dynamic enabling/disabling of sections based on the selected task  
  - Real-time status and error messages

## Requirements

- Windows with PowerShell (tested with PowerShell 5.1)
- Exchange Management Tools installed on the client
- User must have permission to run remote Exchange PowerShell

## Usage

1. **Clone or Download** this repository.

2. **Run the Script**  
   Start PowerShell as an administrator, then run:
   ```powershell
   .\Exchange-Cert-GUI.ps1
   ```

3. **Connect to Exchange**  
   - Enter your Exchange username, password, and server (FQDN).
   - Click **Connect**.

4. **New Certificate Request**  
   - Select "New Certificate Request".
   - Choose "SAN Certificate" or "Wildcard Certificate".
   - Fill in the required fields.
   - Click **Generate CSR** and save the `.req` file.

5. **Complete Certificate Request**  
   - Select "Complete Certificate Request".
   - Browse for your `.p7b` certificate chain file.
   - Click **Complete Request**.

6. **Disconnect**  
   - Use the **Disconnect** button to safely close the Exchange session.

## Notes

- The script automatically removes any pending certificate requests on the target server before generating a new one.
- The script converts DER to PEM for SAN certificates and handles PEM output for wildcards.
- If not connected to Exchange, the script will prompt for credentials and attempt connection when required.

## Troubleshooting

- **WinRM/Remote Connection Issues:**  
  Ensure that WinRM is enabled and the client has network access to the Exchange server.
- **Permissions:**  
  You must have rights to manage certificates on the target Exchange server.
- **PowerShell Execution Policy:**  
  You may need to run `Set-ExecutionPolicy RemoteSigned` (or less restrictive) to allow the script to execute.

## License

MIT

## Credits

Developed by [andrew-kemp](https://github.com/andrew-kemp) with assistance from GitHub Copilot.
