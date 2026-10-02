# Powershell-Bulk-Font-Installer

A PowerShell script (`Fonts.ps1`) that installs all fonts from a network share on a Windows computer. It copies the font files to `C:\Windows\Fonts`, registers them in the registry for all users, and writes a log file per computer.

## What it does

1. Creates the working folders `C:\Temp\Fonts\Logs\` and `C:\Temp\Fonts\Files\` if needed.
2. Starts a log entry in `C:\Temp\Fonts\Logs\<COMPUTERNAME>.log`.
3. Checks that the script runs as administrator. If not, it writes `E - Run as Administrator` to the log and installs nothing. The log folder is created before this check, so the message only reaches the log if the current user can create or write `C:\Temp\Fonts\Logs\` (see [Notes and limitations](#notes-and-limitations)).
4. Checks that the file server answers (`Test-NetConnection -ComputerName $FileServer`). If not, it logs an error and installs nothing.
5. Checks that the font source folder exists (`Test-Path`). If not, it logs an error and installs nothing.
6. Copies the complete content of the source folder to the temporary folder `C:\Temp\Fonts\Files\`.
7. Searches the source folder recursively for `*.ttf`, `*.ttc` and `*.otf` files and handles each file as follows:
   - If a file with the same name already exists in `C:\Windows\Fonts`, the file is not copied again.
   - Otherwise the file is copied from the temporary folder to `C:\Windows\Fonts`. If the copy fails, registration is skipped for that font.
   - If the registry value for the font does not exist yet under `HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts`, it is created as a string value. If it already exists, it is left unchanged.
8. Deletes the temporary folder `C:\Temp\Fonts\Files\`.
9. Writes the totals (total, successful and failed fonts) and an end timestamp to the log.

### Registry entries

The registry value name is built from the file name without its extension plus a type suffix. The value data is the file name.

| Extension | Suffix       | Example value name      | Example value data |
|-----------|--------------|-------------------------|--------------------|
| `.ttf`    | `(TrueType)` | `MyFont (TrueType)`     | `MyFont.ttf`       |
| `.ttc`    | `(TrueType)` | `MyFont (TrueType)`     | `MyFont.ttc`       |
| `.otf`    | `(OpenType)` | `MyFont (OpenType)`     | `MyFont.otf`       |

### Logging

Status messages are written to `C:\Temp\Fonts\Logs\<COMPUTERNAME>.log`. The script appends to this file, so it contains the history of all runs. Each line is prefixed with a level:

- `E`: Error
- `S`: Success
- `i`: Information

Example:

```text
10/02/2026 14:30 - i - Start

E - Error
S - Success
i - Information

i - MyFont
S - Copying of MyFont succeeded
S - Registration MyFont Succeeded

Total Fonts = 1
Successful Fonts = 1
Failed Fonts = 0

10/02/2026 14:30 - i - End
```

A font that is already present and already registered is logged as information and counted as successful.

## Requirements

- Windows
- PowerShell with the `Test-NetConnection` cmdlet available
- Administrator rights (the script checks this and stops without installing anything if they are missing)
- A file server with a shared folder that contains the font files, reachable from the computer
- Write access to `C:\` (the script creates `C:\Temp\Fonts\`)

## Usage

The script has no parameters. Edit the variables at the top of `Fonts.ps1` before running it:

```powershell
$FileServer        = "Fileserver" # FileServer Hostname #
$FontSourceFolder  = "\\Filserver\Font" # Font Folder SMB Address #
```

Set both to your own values, for example:

```powershell
$FileServer        = "fs01"
$FontSourceFolder  = "\\fs01\Fonts"
```

Then run the script from an elevated PowerShell session (Run as administrator):

```powershell
.\Fonts.ps1
```

If script execution is blocked by the execution policy on your system, you can bypass it for a single run:

```powershell
powershell.exe -ExecutionPolicy Bypass -File .\Fonts.ps1
```

Afterwards, check `C:\Temp\Fonts\Logs\<COMPUTERNAME>.log` for the result.

## Variables

| Variable             | Default                                    | Description                                                                 |
|----------------------|--------------------------------------------|-----------------------------------------------------------------------------|
| `$FileServer`        | `"Fileserver"`                             | Host name of the file server, used only for the connectivity check. Must be edited. |
| `$FontSourceFolder`  | `"\\Filserver\Font"`                       | UNC path of the folder that contains the fonts. Must be edited.             |
| `$WindowsFontFolder` | `"C:\Windows\Fonts"`                       | Folder the fonts are copied to.                                             |
| `$TempFileFolder`    | `"C:\Temp\Fonts\Files\"`                   | Temporary local copy of the source folder. Deleted at the end of the run.   |
| `$LogFile`           | `"C:\Temp\Fonts\Logs\<COMPUTERNAME>.log"`  | Log file, named after the computer (`$env:computername`).                   |
| `$RegPath`           | `"HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts"` | Registry key in which the fonts are registered.           |

## Notes and limitations

- `$FileServer` and `$FontSourceFolder` are independent variables. In the script as shipped, the placeholder values do not match (`Fileserver` and `\\Filserver\Font`), so both must be set.
- The working folders and the first log lines are written before the administrator check. When the script runs without elevation and the user cannot create `C:\Temp\Fonts\Logs\` (and it does not exist yet), folder creation and logging fail, so the `E - Run as Administrator` message is not written to a log file.
- The folder creation at the start of the script uses the hardcoded path `C:\Temp\Fonts\`. Changing `$TempFileFolder` or `$LogFile` alone does not change which folders are created.
- The connectivity check uses `Test-NetConnection` without a port, which is a ping test. If the file server does not answer ping, the script logs an error and installs nothing, even if the share itself is reachable.
- Only `.ttf`, `.ttc` and `.otf` files are installed. The whole source folder is copied to the temporary folder first, including any other files in it.
- Fonts in subfolders of the source folder are found by the recursive search, but the script expects every font file directly in the root of the temporary folder. Fonts in subfolders therefore fail at the copy step unless a file with the same name already exists in `C:\Windows\Fonts`. Keep all font files in the root of the source folder.
- The registry value name is derived from the file name, not from the font name stored inside the font file.
- "Already installed" is decided by file name only. An existing file with the same name in `C:\Windows\Fonts` is never overwritten, so the script does not update fonts.
- The script only copies files and writes registry values. It does not notify running applications of new fonts.
- The failure counter is not reliable in one case: when registration fails for a font whose file was already present in `C:\Windows\Fonts`, the counter is set to 1 instead of being increased (`$FailerCount =+ 1`).
- Timestamps in the log use the format `MM/dd/yyyy HH:mm`.

## License

GNU General Public License v3.0. See [LICENSE](LICENSE).
