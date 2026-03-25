# install:
#   powershell -ExecutionPolicy Bypass -File vscode-remote.ps1 -- --install
# uninstall:
#   powershell -ExecutionPolicy Bypass -File vscode-remote.ps1 -- --uninstall

param([string]$uri)

if ($uri -eq '--install') {
    $script = $MyInvocation.MyCommand.Path

    New-Item -Path "HKCU:\Software\Classes\vscode-remote" -Force | Out-Null
    Set-ItemProperty -Path "HKCU:\Software\Classes\vscode-remote" `
        -Name "(Default)" -Value "URL:vscode-remote Protocol"
    Set-ItemProperty -Path "HKCU:\Software\Classes\vscode-remote" `
        -Name "URL Protocol" -Value ""
    New-Item -Path "HKCU:\Software\Classes\vscode-remote\shell\open\command" -Force | Out-Null
    Set-ItemProperty -Path "HKCU:\Software\Classes\vscode-remote\shell\open\command" `
        -Name "(Default)" -Value "powershell.exe -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$script`" `"%1`""

    Write-Host "Done. $script"
    exit 0
}

if ($uri -eq '--uninstall') {
    Remove-Item -Path "HKCU:\Software\Classes\vscode-remote" -Recurse -Force
    Write-Host "Uninstalled."
    exit 0
}

if ($uri -match '^vscode-remote://wsl\+([^/]+)(/.*)')  {
    $distro = $matches[1]
    $path = $matches[2] -replace '/', '\'
    $uncPath = "\\wsl.localhost\$distro$path"
    Start-Process explorer.exe -ArgumentList $uncPath
} else {
    # vscode-remote以外はVSCodeで開く
    $codePath = "$env:LOCALAPPDATA\Programs\Microsoft VS Code\Code.exe"
    Start-Process $codePath -ArgumentList "$uri"
}
