function Test-Python312Installed {
    $pythonInstalled = $false
    $path = [System.Environment]::GetEnvironmentVariable("Path", [System.EnvironmentVariableTarget]::Machine)
    $path += ";" + [System.Environment]::GetEnvironmentVariable("Path", [System.EnvironmentVariableTarget]::User)
    $paths = $path.Split(';')
    foreach ($p in $paths) {
        if (Test-Path "$p\python.exe") {
            $version = & "$p\python.exe" --version
            if ($version -like "*Python 3.12*") {
                $pythonInstalled = $true
                break
            }
        }
    }

    # Check the registry if not found in the PATH
    if (-not $pythonInstalled) {
        try {
            $regPaths = @(
                "HKLM:\SOFTWARE\Python\PythonCore\3.12\",
                "HKCU:\SOFTWARE\Python\PythonCore\3.12\"
            )
            foreach ($regPath in $regPaths) {
                if (Test-Path $regPath) {
                    $pythonInstalled = $true
                    break
                }
            }
        }
        catch {
            Write-Error "Error checking registry for Python 3.12 installation."
        }
    }

    return $pythonInstalled
}

if (Test-Python312Installed) {
    Write-Output "Python 3.12 is already installed."
}
else {
    Write-Output "Python 3.12 is not installed. Beginning installation..."

    # Specify the Python installer URL
    $installerUrl = "www.python.org/ftp/python/3.12.2/python-3.12.2-embed-amd64.zip"
    $installerPath = "$env:TEMP\python-3.12-installer.exe"

    # Download the installer
    Invoke-WebRequest -Uri $installerUrl -OutFile $installerPath

    # Install Python 3.12 silently, adjust the installer flags as needed
    Start-Process -FilePath $installerPath -Args "/quiet InstallAllUsers=1 PrependPath=1" -Wait -NoNewWindow

    # Verify installation
    if (Test-Python312Installed) {
        Write-Output "Python 3.12 has been successfully installed."
    }
    else {
        Write-Error "Python 3.12 installation failed."
    }
}

Write-Output "Cloning Repo."
Invoke-WebRequest -Uri "github.com/MainAlexStark/EI_Protocol/archive/refs/heads/main.zip" -OutFile "repository.zip"

Write-Output "Extracting Repo."
Expand-Archive -LiteralPath "repository.zip" -DestinationPath "." 

Write-Output "Removing Zip."
Remove-Item -Path "repository.zip" -Force

Set-Location "EI_Protocol-main"
python3.exe -m pip install -r requirements.txt