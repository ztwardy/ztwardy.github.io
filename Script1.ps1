<#

	Using PowerShell to Add “Open PowerShell Here” and “Open Command Prompt Here” to Explorer Context Menus
	
	see https://www.ucunleashed.com/3872
#>
# folders
if (-not (Test-Path -Path "Registry::HKEY_CLASSES_ROOT\Directory\shell\PowerShellHere")) {
    New-Item -ItemType String -Path "Registry::HKEY_CLASSES_ROOT\Directory\shell\PowerShellHere" -Value "Open PowerShell Here" -Force | Out-Null
    New-Item -ItemType String -Path "Registry::HKEY_CLASSES_ROOT\Directory\shell\PowerShellHere\command" -Value "$env:SystemRoot\system32\WindowsPowerShell\v1.0\powershell.exe -NoExit -Command Set-Location '%V'" -Force | Out-Null
    New-ItemProperty -PropertyType String -Path "Registry::HKEY_CLASSES_ROOT\Directory\shell\PowerShellHere" -Name "Icon" -Value "$env:SystemRoot\system32\WindowsPowerShell\v1.0\powershell.exe" -Force | Out-Null
}
if (-not (Test-Path -Path "Registry::HKEY_CLASSES_ROOT\Directory\shell\PowerShellHereAdmin")) {
    New-Item -ItemType String -Path "Registry::HKEY_CLASSES_ROOT\Directory\shell\PowerShellHereAdmin" -Value "Open PowerShell Here (Admin)" -Force | Out-Null
    New-Item -ItemType String -Path "Registry::HKEY_CLASSES_ROOT\Directory\shell\PowerShellHereAdmin\command" -Value "PowerShell -Command `"Start-Process $env:SystemRoot\system32\WindowsPowerShell\v1.0\powershell.exe`" -verb runAs -ArgumentList '-NoExit','cd','%V'" -Force | Out-Null
    New-ItemProperty -PropertyType String -Path "Registry::HKEY_CLASSES_ROOT\Directory\shell\PowerShellHereAdmin" -Name "Icon" -Value "$env:SystemRoot\system32\WindowsPowerShell\v1.0\powershell.exe,1" -Force | Out-Null
}
if (-not (Test-Path -Path "Registry::HKEY_CLASSES_ROOT\Directory\shell\CmdHere")) {
    New-Item -Path "Registry::HKEY_CLASSES_ROOT\Directory\shell\CmdHere" -Value "Open Command Prompt Here" -ItemType string -Force | Out-Null
    New-Item -Path "Registry::HKEY_CLASSES_ROOT\Directory\shell\CmdHere\command" -Value 'cmd.exe /k pushd %L' -ItemType string -Force | Out-Null
    New-ItemProperty -Path "Registry::HKEY_CLASSES_ROOT\Directory\shell\CmdHere" -Name "Icon" -Value "C:\Windows\System32\cmd.exe,0" -Type string -Force | Out-Null
}

# Explorer background
if (-not (Test-Path -Path "Registry::HKEY_CLASSES_ROOT\Directory\Background\shell\PowerShellHere")) {
    New-Item -ItemType String -Path "Registry::HKEY_CLASSES_ROOT\Directory\Background\shell\PowerShellHere" -Value "Open PowerShell Here" -Force | Out-Null
    New-Item -ItemType String -Path "Registry::HKEY_CLASSES_ROOT\Directory\Background\shell\PowerShellHere\command" -Value "$env:SystemRoot\system32\WindowsPowerShell\v1.0\powershell.exe -NoExit -Command Set-Location '%V'" -Force | Out-Null
    New-ItemProperty -PropertyType String -Path "Registry::HKEY_CLASSES_ROOT\Directory\Background\shell\PowerShellHere" -Name "Icon" -Value "$env:SystemRoot\system32\WindowsPowerShell\v1.0\powershell.exe" -Force | Out-Null
}
if (-not (Test-Path -Path "Registry::HKEY_CLASSES_ROOT\Directory\Background\shell\PowerShellHereAdmin")) {
    New-Item -ItemType String -Path "Registry::HKEY_CLASSES_ROOT\Directory\Background\shell\PowerShellHereAdmin" -Value "Open PowerShell Here (Admin)" -Force | Out-Null
    New-Item -ItemType String -Path "Registry::HKEY_CLASSES_ROOT\Directory\Background\shell\PowerShellHereAdmin\command" -Value "PowerShell -Command `"Start-Process $env:SystemRoot\system32\WindowsPowerShell\v1.0\powershell.exe`" -verb runAs -ArgumentList '-NoExit','cd','%V'" -Force | Out-Null
    New-ItemProperty -PropertyType String -Path "Registry::HKEY_CLASSES_ROOT\Directory\Background\shell\PowerShellHereAdmin" -Name "Icon" -Value "$env:SystemRoot\system32\WindowsPowerShell\v1.0\powershell.exe,1" -Force | Out-Null
}
if (-not (Test-Path -Path "Registry::HKEY_CLASSES_ROOT\Directory\Background\shell\CmdHere")) {
    New-Item -Path "Registry::HKEY_CLASSES_ROOT\Directory\Background\shell\CmdHere" -Value "Open Command Prompt Here" -ItemType string -Force | Out-Null
    New-Item -Path "Registry::HKEY_CLASSES_ROOT\Directory\Background\shell\CmdHere\command" -Value 'cmd.exe /k pushd %L' -ItemType string -Force | Out-Null
    New-ItemProperty -Path "Registry::HKEY_CLASSES_ROOT\Directory\Background\shell\CmdHere" -Name "Icon" -Value "C:\Windows\System32\cmd.exe,0" -Type string -Force | Out-Null
}

# root drives
if (-not (Test-Path -Path "Registry::HKEY_CLASSES_ROOT\Drive\shell\PowerShellHere")) {
    New-Item -ItemType String -Path "Registry::HKEY_CLASSES_ROOT\Drive\shell\PowerShellHere" -Value "Open PowerShell Here" -Force | Out-Null
    New-Item -ItemType String -Path "Registry::HKEY_CLASSES_ROOT\Drive\shell\PowerShellHere\command" -Value "$env:SystemRoot\system32\WindowsPowerShell\v1.0\powershell.exe -NoExit -Command Set-Location '%V'" -Force | Out-Null
    New-ItemProperty -PropertyType String -Path "Registry::HKEY_CLASSES_ROOT\Drive\shell\PowerShellHere" -Name "Icon" -Value "$env:SystemRoot\system32\WindowsPowerShell\v1.0\powershell.exe" -Force | Out-Null
}
if (-not (Test-Path -Path "Registry::HKEY_CLASSES_ROOT\Drive\shell\PowerShellHereAdmin")) {
    New-Item -ItemType String -Path "Registry::HKEY_CLASSES_ROOT\Drive\shell\PowerShellHereAdmin" -Value "Open PowerShell Here (Admin)" -Force | Out-Null
    New-Item -ItemType String -Path "Registry::HKEY_CLASSES_ROOT\Drive\shell\PowerShellHereAdmin\command" -Value "PowerShell -Command `"Start-Process $env:SystemRoot\system32\WindowsPowerShell\v1.0\powershell.exe`" -verb runAs -ArgumentList '-NoExit','cd','%V'" -Force | Out-Null
    New-ItemProperty -PropertyType String -Path "Registry::HKEY_CLASSES_ROOT\Drive\shell\PowerShellHereAdmin" -Name "Icon" -Value "$env:SystemRoot\system32\WindowsPowerShell\v1.0\powershell.exe,1" -Force | Out-Null
}
if (-not (Test-Path -Path "Registry::HKEY_CLASSES_ROOT\Drive\shell\CmdHere")) {
    New-Item -Path "Registry::HKEY_CLASSES_ROOT\Drive\shell\CmdHere" -Value "Open Command Prompt Here" -ItemType string -Force | Out-Null
    New-Item -Path "Registry::HKEY_CLASSES_ROOT\Drive\shell\CmdHere\command" -Value 'cmd.exe /k pushd %L' -ItemType string -Force | Out-Null
    New-ItemProperty -Path "Registry::HKEY_CLASSES_ROOT\Drive\shell\CmdHere" -Name "Icon" -Value "C:\Windows\System32\cmd.exe,0" -Type string -Force | Out-Null
}

#
# If you’d rather have them on a menu when you hit SHIFT+Right Click
#folders
if (-Not (Get-ItemProperty -Path "Registry::HKEY_CLASSES_ROOT\Directory\shell\PowerShellHere" -Name "extended")) {
    New-ItemProperty -PropertyType String -Path "Registry::HKEY_CLASSES_ROOT\Directory\shell\PowerShellHere" -Name "extended" -Value $null | Out-Null
}
if (-Not (Get-ItemProperty -Path "Registry::HKEY_CLASSES_ROOT\Directory\shell\PowerShellHereAdmin" -Name "extended")) {
    New-ItemProperty -PropertyType String -Path "Registry::HKEY_CLASSES_ROOT\Directory\shell\PowerShellHereAdmin" -Name "extended" -Value $null | Out-Null
}
if (-Not (Get-ItemProperty -Path "Registry::HKEY_CLASSES_ROOT\Directory\shell\CmdPromptHere" -Name "extended")) {
    New-ItemProperty -PropertyType String -Path "Registry::HKEY_CLASSES_ROOT\Directory\shell\CmdPromptHere" -Name "extended" -Value $null | Out-Null
}

#background
if (-Not (Get-ItemProperty -Path "Registry::HKEY_CLASSES_ROOT\Directory\Background\shell\PowerShellHere" -Name "extended")) {
    New-ItemProperty -PropertyType String -Path "Registry::HKEY_CLASSES_ROOT\Directory\Background\shell\PowerShellHere" -Name "extended" -Value $null | Out-Null
}
if (-Not (Get-ItemProperty -Path "Registry::HKEY_CLASSES_ROOT\Directory\Background\shell\PowerShellHereAdmin" -Name "extended")) {
    New-ItemProperty -PropertyType String -Path "Registry::HKEY_CLASSES_ROOT\Directory\Background\shell\PowerShellHereAdmin" -Name "extended" -Value $null | Out-Null
}
if (-Not (Get-ItemProperty -Path "Registry::HKEY_CLASSES_ROOT\Directory\Background\shell\CmdPromptHere" -Name "extended")) {
    New-ItemProperty -PropertyType String -Path "Registry::HKEY_CLASSES_ROOT\Directory\Background\shell\CmdPromptHere" -Name "extended" -Value $null | Out-Null
}

#root drives
if (-Not (Get-ItemProperty -Path "Registry::HKEY_CLASSES_ROOT\Drive\shell\PowerShellHere" -Name "extended")) {
    New-ItemProperty -PropertyType String -Path "Registry::HKEY_CLASSES_ROOT\Drive\shell\PowerShellHere" -Name "extended" -Value $null | Out-Null
}
if (-Not (Get-ItemProperty -Path "Registry::HKEY_CLASSES_ROOT\Drive\shell\PowerShellHereAdmin" -Name "extended")) {
    New-ItemProperty -PropertyType String -Path "Registry::HKEY_CLASSES_ROOT\Drive\shell\PowerShellHereAdmin" -Name "extended" -Value $null | Out-Null
}
if (-Not (Get-ItemProperty -Path "Registry::HKEY_CLASSES_ROOT\Drive\shell\CmdPromptHere" -Name "extended")) {
    New-ItemProperty -PropertyType String -Path "Registry::HKEY_CLASSES_ROOT\Drive\shell\CmdPromptHere" -Name "extended" -Value $null | Out-Null
}