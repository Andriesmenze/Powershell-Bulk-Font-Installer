# Powershell-Bulk-Font-Installer

A PowerShell script (`Fonts.ps1`) that installs all fonts from a network share on a Windows computer. It copies the font files to `C:\Windows\Fonts`, registers them in the registry for all users, and writes a log file per computer.

## What it does

1. Checks that it runs elevated (`#Requires -RunAsAdministrator`). Without elevation PowerShell stops before the script does anything: it prints an error, returns exit code `1` and writes no log.
2. Creates the log folder (default `C:\Windows\Logs\FontInstaller`) and starts a log entry in `<COMPUTERNAME>.log`. If the log folder can't be created, for example a share that isn't reachable yet, the log goes to the default folder instead and the problem is logged as an error. A log larger than 1 MB is first renamed to `<COMPUTERNAME>.log.old`.
3. Waits until the font source folder can be reached, for about `-WaitForSourceSeconds` (default 30 seconds). This covers computer startup scripts that run before the network is ready. If the folder still can't be reached, it logs an error and installs nothing.
4. Searches the source folder and its subfolders for `.ttf`, `.ttc` and `.otf` files, skipping macOS metadata files (`._Name.ttf`). Fonts in the source folder itself come first, then the fonts in subfolders. Each file is handled as follows:
   - If an earlier font had the same file name, the file is skipped and counted as failed. `C:\Windows\Fonts` has no subfolders, so only one font per file name can be installed.
   - If the font file isn't registered yet but its registry value name (see below) is already used for a different file, the font is skipped and counted as failed. The existing value is left alone.
   - If a file with the same name already exists in `C:\Windows\Fonts`, it is not copied again. If its size differs from the file on the share, the log says so.
   - Otherwise the file is copied straight from the share to `C:\Windows\Fonts`, under a temporary name (`~FontInstaller-<id>.tmp`) that is renamed once the copy is complete. If the copy fails, the temporary file is removed, the error is logged and registration is skipped.
   - If the font file is already registered under `HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts`, under any value name, nothing changes. Otherwise a string value is created.
5. Writes the totals and an end timestamp to the log, prints a one-line summary and exits with code `0` when everything succeeded, or `1` when anything failed.

Temporary files left behind by a run that was stopped in the middle of a copy are removed by a later run once they are more than an hour old.

### Registry entries

The registry value name is built from the file name without its extension plus a type suffix. The value data is the file name.

| Extension | Suffix       | Example value name      | Example value data |
|-----------|--------------|-------------------------|--------------------|
| `.ttf`    | `(TrueType)` | `MyFont (TrueType)`     | `MyFont.ttf`       |
| `.ttc`    | `(TrueType)` | `MyFont (TrueType)`     | `MyFont.ttc`       |
| `.otf`    | `(OpenType)` | `MyFont (OpenType)`     | `MyFont.otf`       |

A font file that is already registered under another name, for example because it was installed through Windows Settings, keeps that registration and doesn't get a second value.

### Logging

Status messages are written to `C:\Windows\Logs\FontInstaller\<COMPUTERNAME>.log` (UTF-8). The script appends to this file, so it contains the history of all runs, until it grows past 1 MB and is renamed to `<COMPUTERNAME>.log.old`. Each line is prefixed with a level:

- `E`: Error
- `S`: Success
- `i`: Information

Example:

```text
2026-10-02T14:30:05 - i - Start

E - Error
S - Success
i - Information

i - Running As CONTOSO\PC01$
i - Font Source \\fs01.contoso.com\Fonts

i - MyFont
S - Copying of MyFont succeeded
S - Registration MyFont Succeeded

Total Fonts = 1
Successful Fonts = 1
Failed Fonts = 0

2026-10-02T14:30:06 - i - End
```

When the script runs as SYSTEM, `Running As` shows the computer account (`PC01$`), which is the account that needs read access to the share. Error lines include the reason, for example `E - Copying of MyFont failed: Access to the path ... is denied.` A font that is already present and already registered is logged as information and counted as successful.

## Requirements

- Windows with Windows PowerShell 5.1 or PowerShell 7
- Administrator rights, or running as SYSTEM, for example as a Group Policy computer startup script
- A shared folder with the font files that the account running the script can read. A computer startup script runs as SYSTEM and reaches the share with the computer account, so give `Domain Computers` read access on both the share and the folder.

The script only uses what PowerShell allows in Constrained Language Mode, so it also works where App Control for Business (WDAC) or AppLocker runs unsigned scripts in that mode.

## Security

The script installs every font on the share, as SYSTEM, on every computer that runs it. Treat the share accordingly:

- Only administrators should be able to write to the share.
- Use a fully qualified server name or a DFS path (`\\fs01.contoso.com\Fonts`) instead of a short name like `\\Fileserver\Font`. Short names that don't resolve in DNS fall back to broadcast name resolution, which anyone on the network can answer.
- Keep the log folder writable for administrators only. The default under `C:\Windows\Logs` is. Don't point `-LogFolder` at a folder like `C:\Temp`, where standard users can create and change files.

## Usage

Run the script from an elevated PowerShell session (Run as administrator) and pass your share:

```powershell
.\Fonts.ps1 -FontSourceFolder '\\fs01.contoso.com\Fonts'
```

If script execution is blocked by the execution policy on your system, you can bypass it for a single run:

```powershell
powershell.exe -ExecutionPolicy Bypass -File .\Fonts.ps1 -FontSourceFolder "\\fs01.contoso.com\Fonts"
```

As a Group Policy computer startup script, add `Fonts.ps1` on the **PowerShell Scripts** tab of the startup script settings and enter `-FontSourceFolder \\fs01.contoso.com\Fonts` as the script parameters. You can also change the default value of `$FontSourceFolder` in the `param` block at the top of the script.

Afterwards, check `C:\Windows\Logs\FontInstaller\<COMPUTERNAME>.log` or the exit code for the result. New fonts appear for users at their next sign-in.

`Get-Help .\Fonts.ps1 -Full` shows the built-in help.

## Parameters

| Parameter               | Default                          | Description                                                                 |
|-------------------------|----------------------------------|-----------------------------------------------------------------------------|
| `-FontSourceFolder`     | `\\Fileserver\Font`              | UNC path of the folder that contains the fonts. Set it to your own share.   |
| `-LogFolder`            | `C:\Windows\Logs\FontInstaller`  | Folder for the log file `<COMPUTERNAME>.log`. Must not be writable by standard users. If it can't be created, the default folder is used. |
| `-WaitForSourceSeconds` | `30`                             | About how long to keep trying to reach the source folder. `0` tries once. One attempt against a server that doesn't answer can take about 20 seconds by itself. |

## Exit codes

| Code | Meaning                                                                 |
|------|-------------------------------------------------------------------------|
| `0`  | Every font was installed or was already installed.                      |
| `1`  | The source folder couldn't be reached or read, the log folder couldn't be used, or at least one font failed; details are in the log. Also returned when the script isn't elevated, in which case PowerShell refuses to run it and there is no log. |

## Notes and limitations

- The registry value name is derived from the file name, not from the font name stored inside the font file.
- Fonts that are already in `C:\Windows\Fonts` are never replaced, because Windows locks fonts that are in use. If the version on the share has a different size, the log says so. To roll out a new version of a font, give the file a new name.
- Removing a font from the share doesn't remove it from the computers.
- Fonts with the same file name in different subfolders of the share can't all be installed. Only the first one found is; the others are logged as errors.
- The script copies files and writes registry values. It doesn't notify running applications, so a new font appears in sessions that start after the installation (the next sign-in).
- Only `.ttf`, `.ttc` and `.otf` files are installed. Other files on the share are ignored.
- Run the script through one deployment method at a time (for example Group Policy or Intune, not both). Overlapping runs don't damage the installation, but they can report errors and mix up each other's log lines.

### Upgrading from the old version

- Earlier versions had the share hardcoded in two variables at the top of the script (`$FileServer` and `$FontSourceFolder`). Pass your share with `-FontSourceFolder` or set it as the default in the `param` block; otherwise the script looks for `\\Fileserver\Font`.
- Earlier versions used `C:\Temp\Fonts\` for a temporary copy of the share and for the log. That folder isn't used anymore and can be deleted.

## License

GNU General Public License v3.0. See [LICENSE](LICENSE).
