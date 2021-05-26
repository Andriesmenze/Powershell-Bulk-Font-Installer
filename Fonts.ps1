# varibles #
$FileServer        = "Fileserver" # FileServer Hostname #
$FontSourceFolder  = "\\Filserver\Font" # Font Folder SMB Address #
$Fonts             = $FontSourceFolder,"\*" -join ""
$WindowsFontFolder = "C:\Windows\Fonts"
$TempFileFolder    = "C:\Temp\Fonts\Files\"
$LogFile           = "C:\Temp\Fonts\Logs\",$env:computername,".log" -join ""
$RegPath           = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts"
$ttf               = "(TrueType)"
$ttc               = "(TrueType)"
$otf               = "(OpenType)"
$TotalFonts        = 0
$SuccessCount      = 0
$FailerCount       = 0
# varibles #
# Functions #
Function LogWrite{
    Param (
        [string]$logstring
    )
    Add-content $Logfile -value $logstring
}
function Test-Administrator{  
    $User = [Security.Principal.WindowsIdentity]::GetCurrent();
    (New-Object Security.Principal.WindowsPrincipal $User).IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)  
}
function Test-RegistryValue{
    param ( 
        [parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]$Path,
        [parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]$Value
    )
    try {
        Get-ItemProperty -Path $Path | Select-Object -ExpandProperty $Value -ErrorAction Stop | Out-Null
        return $true
    }
    catch {
        return $false
    } 
}
# Functions #
# Prepare Folders #
If (-not(Test-Path "C:\Temp\")){
	New-Item -Path "C:\" -Name "Temp" -ItemType "directory" -Force
}
If (-not(Test-Path "C:\Temp\Fonts\")){
	New-Item -Path "C:\Temp\" -Name "Fonts" -ItemType "directory"
}
If (-not(Test-Path "C:\Temp\Fonts\Logs\")){
	New-Item -Path "C:\Temp\Fonts\" -Name "Logs" -ItemType "directory"
}
New-Item -Path "C:\Temp\Fonts\" -Name "Files" -ItemType "directory"
# Prepare Folders #
# script #
$Time = Get-Date -Format "MM/dd/yyyy HH:mm"
LogWrite "$Time - i - Start"
LogWrite ""
LogWrite "E - Error"
LogWrite "S - Success"
LogWrite "i - Information"
LogWrite ""
if (Test-Administrator){
    if (Test-NetConnection -ComputerName $FileServer -InformationLevel Quiet){
        if (Test-Path -Path $FontSourceFolder){
            Copy-Item -Path $Fonts -Destination $TempFileFolder -Recurse
            Get-ChildItem -Path $Fonts -Include '*.ttf','*.ttc','*.otf' -Recurse | ForEach-Object {
                $TotalFonts += 1
                $FontFileName  = $_.Name.ToString()
                $FontName      = $FontFileName.ToString().Substring(0, $FontFileName.Length-4)
                $LocalFontPath = $TempFileFolder,$FontFileName -join "\"
                $NewFontPath = $WindowsFontFolder,$FontFileName -join "\"
                LogWrite "i - $FontName"
                if ($FontFileName -like "*.ttf"){
                    $FileType = $ttf
                }
                elseif ($FontFileName -like "*.ttc") {
                    $FileType = $ttc    
                }
                elseif ($FontFileName -like "*.otf") {
                    $FileType = $otf        
                }
                $RegKeyName = $FontName,$FileType -join " "
                $RegKeyValue = $FontFileName
                if (Test-Path -Path $NewFontPath -PathType Leaf){
                    LogWrite "i - $FontName Is Already In Windows Font Directory"
                    if(Test-RegistryValue -Path $RegPath -Value $RegKeyName){
                        LogWrite "i - $FontName Is Already Registered"
                        $SuccessCount += 1
                    }
                    else{
                        Try{
                            $null = New-ItemProperty -Path $RegPath -Name $RegKeyname -Value $RegKeyValue -PropertyType String -Force -ErrorAction Stop
                            LogWrite "S - Registration $FontName Succeeded"
                            $SuccessCount += 1
                        }
                        Catch{
                            LogWrite "E - Registration $FontName failed!"
                            $FailerCount =+ 1
                        }
                    }
                }
                else{
                    try{
                            Copy-Item -Path $LocalFontPath -Destination $WindowsFontFolder -Force -ErrorAction Stop 
                            LogWrite "S - Copying of $FontName succeeded"
                        if(Test-RegistryValue -Path $RegPath -Value $RegKeyName){
                            LogWrite "i - $FontName Is Already Registered"
                            $SuccessCount += 1
                        }
                        else{
                            Try{
                                $null = New-ItemProperty -Path $RegPath -Name $RegKeyname -Value $RegKeyValue -PropertyType String -Force -ErrorAction Stop
                                LogWrite "S - Registration $FontName Succeeded"
                                $SuccessCount += 1
                            }
                            Catch{
                                LogWrite "E - Registration $FontName failed!"
                                $FailerCount += 1
                            }
                        }
                    }
                    Catch{
                        LogWrite "E - Copying of $FontName failed!"
                        LogWrite "E - Skipping Registration For $FontName"
                        $FailerCount += 1
                    }
                }
                LogWrite ""
            }
        }
        else{
        LogWrite "E - Can't find the folder $FontSourceFolder"
        LogWrite ""
        }
    }
    else{
    LogWrite "E - Can't contact the FileServer $FileServer"
    LogWrite ""
    }
}
else{
    LogWrite "E - Run as Administrator"
    LogWrite ""
}
Remove-Item $TempFileFolder -Recurse -Force
LogWrite "Total Fonts = $TotalFonts"
LogWrite "Successful Fonts = $SuccessCount"
LogWrite "Failed Fonts = $FailerCount"
LogWrite ""
$Time = Get-Date -Format "MM/dd/yyyy HH:mm"
LogWrite "$Time - i - End"
# script #