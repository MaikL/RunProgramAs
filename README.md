# RunProgramAs

RunProgramAs is a PowerShell-based application that allows users to run programs with different credentials. It provides a graphical user interface (GUI) for selecting programs, managing credentials, and configuring program arguments.

## Features

- **Run Programs with Alternate Credentials**: Start `.exe`, `.bat`, or `.cmd` files using saved credentials.
- **Credential Management**: Save, delete, and reuse multiple sets of credentials.
- **Multilingual Support**: Supports English and German languages, with translations stored in `translations.json`.
- **GUI-Based Interaction**: User-friendly interface built with Windows Forms.
- **Logging**: Logs application events and errors to `Log/RunProgramAs.log` with automatic log rotation.

## Prerequisites

- Windows operating system
- PowerShell 5.1 or later
- .NET Framework (required for Windows Forms)

## Installation

. Clone or download this repository.
2. Ensure the following files are in the same directory:
   - `RunProgramAs.ps1`
   - `RunProgramAs.cmd`
   - `translations.json`
3. Run the program using the provided batch file:
   ```cmd
   RunProgramAs.cmd
   ```
## Usage

1. Launch the program by running RunProgramAs.cmd.

2. Use the GUI to:
- [?] Select a program to run.
- [?] Enter or select credentials.
- [?] Provide optional arguments for the program.

3. Click Start Program to execute the selected program with the specified credentials.

## Language Selection
Use the "Deutsch" or "English" buttons to switch between German and English.

## Managing Credentials
- Add New Credentials: Click the New Credentials button to save a new set of credentials.
- Delete Credentials: Right-click on the credentials dropdown and select Delete to remove a saved credential.
## Logs
Logs are stored in the Log directory. The log file rotates automatically when it exceeds 1 MB.

## Known Issues
Only .exe, .bat, and .cmd files are supported.
Ensure the selected program's directory has appropriate read permissions.
## Contributing
Contributions are welcome! Feel free to open issues or submit pull requests.

## License
This project is licensed under the MIT License. See the LICENSE file for details.

## Acknowledgments
Windows Forms for GUI components
PowerShell for scripting capabilities

## Contact
MaikL - maikl0124@outlook.com