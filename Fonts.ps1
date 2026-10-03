#Requires -RunAsAdministrator

<#
.SYNOPSIS
    Installs all fonts from a network share for all users of this computer.

.DESCRIPTION
    Searches FontSourceFolder, including its subfolders, for .ttf, .ttc and .otf files. Every
    font that is not in the Windows Fonts folder yet is copied there straight from the share,
    and every font is registered under HKLM so it is available to all users after their next
    sign-in.

    The script must run elevated, for example as a Group Policy computer startup script
    (SYSTEM) or from an elevated PowerShell session. As SYSTEM the share is accessed with the
    computer account, so the share and NTFS permissions must allow Domain Computers to read it.

    Every font on the share is installed as SYSTEM on every computer that runs this script, so
    only administrators should be able to write to the share.

    The log is written to LogFolder\<COMPUTERNAME>.log. The exit code is 0 when every font is
    installed and 1 when the share can't be reached or anything failed.

.PARAMETER FontSourceFolder
    UNC path of the folder that holds the fonts. Prefer a fully qualified name or a DFS path
    (\\fileserver.contoso.com\Fonts) over a short host name.

.PARAMETER LogFolder
    Folder for the log file. The default can only be written by administrators. Don't use a
    folder that standard users can write to, such as C:\Temp. If the folder can't be created,
    for example a share that can't be reached yet, the log goes to the default folder instead.

.PARAMETER WaitForSourceSeconds
    About how long to keep trying when the share can't be reached yet, for example because the
    network is still starting during a computer startup script. 0 tries once. One attempt
    against a server that doesn't answer can take about 20 seconds by itself.

.EXAMPLE
    .\Fonts.ps1 -FontSourceFolder '\\fileserver.contoso.com\Fonts'
#>
[CmdletBinding()]
param (
    [ValidateNotNullOrEmpty()]
    [string]$FontSourceFolder = '\\Fileserver\Font',

    [ValidateNotNullOrEmpty()]
    [string]$LogFolder = (Join-Path $env:SystemRoot 'Logs\FontInstaller'),

    [ValidateRange(0, 3600)]
    [int]$WaitForSourceSeconds = 30
)
# Variables #
$WindowsFontFolder = Join-Path $env:SystemRoot 'Fonts'
$RegPath           = 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts'
$DefaultLogFolder  = Join-Path $env:SystemRoot 'Logs\FontInstaller'
$LogFolderError    = $null
$MaxLogSize        = 1MB
$TempFontPrefix    = '~FontInstaller-'
$FontTypes         = @{ # File extension and the matching registry value name suffix #
    '.ttf' = '(TrueType)'
    '.ttc' = '(TrueType)'
    '.otf' = '(OpenType)'
}
$TotalFonts        = 0
$SuccessCount      = 0
$FailureCount      = 0
$ExitCode          = 0
# Variables #
# Functions #
# Only cmdlets and core types are used, so the script also works in Constrained Language Mode #
function Write-Log {
    param (
        [string]$Message = ''
    )
    Add-Content -LiteralPath $LogFile -Value $Message -Encoding UTF8
}
function Get-TimeStamp {
    # ISO 8601, the same on every computer whatever its regional settings #
    Get-Date -Format 's'
}
function Wait-FontSource {
    # The network may still be starting when this runs as a computer startup script #
    $Deadline = (Get-Date).AddSeconds($WaitForSourceSeconds)
    while (-not (Test-Path -LiteralPath $FontSourceFolder -PathType Container)) {
        if ((Get-Date) -ge $Deadline) {
            return $false
        }
        Start-Sleep -Seconds 5
    }
    return $true
}
function Get-RegisteredFont {
    # Registry value name and file (the value data) of every registered font #
    $Fonts = @{}
    foreach ($Value in (Get-ItemProperty -LiteralPath $RegPath -ErrorAction Stop).PSObject.Properties) {
        if ($Value.Name -notin 'PSPath', 'PSParentPath', 'PSChildName', 'PSDrive', 'PSProvider') {
            $Fonts[$Value.Name] = [string]$Value.Value
        }
    }
    $Fonts
}
function Install-Font {
    param (
        [Parameter(Mandatory = $true)]
        [System.IO.FileInfo]$Font
    )
    $FontName     = $Font.BaseName
    $NewFontPath  = Join-Path $WindowsFontFolder $Font.Name
    $RegKeyName   = '{0} {1}' -f $FontName, $FontTypes[$Font.Extension]
    $IsRegistered = ($Registered.Values -contains $Font.Name) -or ($Registered.Values -contains $NewFontPath)

    if (-not $IsRegistered -and $Registered.ContainsKey($RegKeyName)) {
        # Never take over a registry value that belongs to another font file #
        Write-Log "E - Registration $FontName failed: $RegKeyName is already registered for $($Registered[$RegKeyName])"
        return $false
    }
    if (Test-Path -LiteralPath $NewFontPath -PathType Leaf) {
        Write-Log "i - $FontName Is Already In Windows Font Directory"
        if ((Get-Item -LiteralPath $NewFontPath).Length -ne $Font.Length) {
            # A font that is in use can't be replaced, so a changed font is reported instead #
            Write-Log "i - $FontName Differs From $($Font.FullName) And Is Not Updated"
        }
    }
    else {
        # A temporary name until the copy is complete, so an interrupted copy never looks installed #
        $TempFontPath = Join-Path $WindowsFontFolder "$TempFontPrefix$(New-Guid).tmp"
        try {
            Copy-Item -LiteralPath $Font.FullName -Destination $TempFontPath -ErrorAction Stop
            # Windows PowerShell 5.1 doesn't always raise an error when nothing was copied #
            if ((Get-Item -LiteralPath $TempFontPath -ErrorAction Stop).Length -ne $Font.Length) {
                throw "The copy of $($Font.FullName) is incomplete"
            }
            Move-Item -LiteralPath $TempFontPath -Destination $NewFontPath -ErrorAction Stop
            Write-Log "S - Copying of $FontName succeeded"
        }
        catch {
            Remove-Item -LiteralPath $TempFontPath -Force -ErrorAction SilentlyContinue
            Write-Log "E - Copying of $FontName failed: $($_.Exception.Message)"
            Write-Log "E - Skipping Registration For $FontName"
            return $false
        }
    }
    if ($IsRegistered) {
        Write-Log "i - $FontName Is Already Registered"
        return $true
    }
    try {
        $null = New-ItemProperty -LiteralPath $RegPath -Name $RegKeyName -Value $Font.Name -PropertyType String -ErrorAction Stop
        $Registered[$RegKeyName] = $Font.Name
        Write-Log "S - Registration $FontName Succeeded"
        return $true
    }
    catch {
        Write-Log "E - Registration $FontName failed: $($_.Exception.Message)"
        return $false
    }
}
# Functions #
# Prepare Log #
try {
    $null = New-Item -Path $LogFolder -ItemType Directory -Force -ErrorAction Stop
    if (-not (Test-Path -LiteralPath $LogFolder -PathType Container)) {
        throw "$LogFolder is not a folder"
    }
}
catch {
    # For example a log share that can't be reached yet during a computer startup script #
    $LogFolderError = "E - Can't use log folder $LogFolder, using $DefaultLogFolder instead: $($_.Exception.Message)"
    $LogFolder      = $DefaultLogFolder
    $null = New-Item -Path $LogFolder -ItemType Directory -Force -ErrorAction Stop
}
$LogFile = Join-Path $LogFolder "$env:COMPUTERNAME.log"
if ((Test-Path -LiteralPath $LogFile -PathType Leaf) -and (Get-Item -LiteralPath $LogFile).Length -gt $MaxLogSize) {
    Move-Item -LiteralPath $LogFile -Destination "$LogFile.old" -Force
}
# Prepare Log #
# Script #
Write-Log "$(Get-TimeStamp) - i - Start"
Write-Log ''
Write-Log 'E - Error'
Write-Log 'S - Success'
Write-Log 'i - Information'
Write-Log ''
Write-Log "i - Running As $env:USERDOMAIN\$env:USERNAME"
Write-Log "i - Font Source $FontSourceFolder"
Write-Log ''
if ($LogFolderError) {
    Write-Log $LogFolderError
    Write-Log ''
    $ExitCode = 1
}
# Leftovers of copies that were interrupted, for example when a script timeout ended the run #
Get-ChildItem -LiteralPath $WindowsFontFolder -Filter "$TempFontPrefix*.tmp" -File |
    Where-Object { $_.CreationTime -lt (Get-Date).AddHours(-1) } |
    Remove-Item -Force -ErrorAction SilentlyContinue
if (Wait-FontSource) {
    # macOS metadata files (._Name.ttf) have font extensions but aren't fonts #
    $FontFiles = Get-ChildItem -LiteralPath $FontSourceFolder -File -Recurse -ErrorAction SilentlyContinue -ErrorVariable ListErrors |
        Where-Object { $FontTypes.ContainsKey($_.Extension) -and $_.Name -notlike '._*' } |
        Sort-Object DirectoryName, Name
    foreach ($ListError in $ListErrors) {
        Write-Log "E - Can't read $($ListError.TargetObject): $($ListError.Exception.Message)"
        $ExitCode = 1
    }
    if (-not $FontFiles) {
        Write-Log "i - No fonts found in $FontSourceFolder"
        Write-Log ''
    }
    $Registered = Get-RegisteredFont
    $Handled    = @{} # File name and the full path of the font that was handled under that name #
    foreach ($Font in $FontFiles) {
        $TotalFonts += 1
        Write-Log "i - $($Font.BaseName)"
        if ($Handled.ContainsKey($Font.Name)) {
            # The Windows Fonts folder is flat, so only one font per file name can be installed #
            Write-Log "E - Skipping $($Font.FullName), it has the same file name as $($Handled[$Font.Name])"
            $FailureCount += 1
        }
        else {
            $Handled[$Font.Name] = $Font.FullName
            if (Install-Font -Font $Font) {
                $SuccessCount += 1
            }
            else {
                $FailureCount += 1
            }
        }
        Write-Log ''
    }
}
else {
    Write-Log "E - Can't access $FontSourceFolder (server unreachable, folder missing or access denied)"
    Write-Log ''
    $ExitCode = 1
}
if ($FailureCount -gt 0) {
    $ExitCode = 1
}
Write-Log "Total Fonts = $TotalFonts"
Write-Log "Successful Fonts = $SuccessCount"
Write-Log "Failed Fonts = $FailureCount"
Write-Log ''
Write-Log "$(Get-TimeStamp) - i - End"
Write-Output "Fonts: $TotalFonts total, $SuccessCount successful, $FailureCount failed. Log: $LogFile"
exit $ExitCode
# Script #
