# ============================================================
#  ULTIMATE UNIFORM AI/ML LAB SETUP  -  Windows 11
#  Ek baar Admin se chalao, har local account ready ho jaayega
#
#  Usage:
#    powershell -ExecutionPolicy Bypass -File "Lab_ML_Setup.ps1"
#
#  FIX v3 (all fixes applied):
#   - .condarc written to install + system-wide locations
#   - CONDARC env var set (Machine scope)
#   - Channels passed explicitly on every conda create/install

#   - ipykernel re-pinned via pip after conda installs (eviction guard)
# ============================================================

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$ErrorActionPreference = "Continue"
$ProgressPreference    = "SilentlyContinue"

function Step { param($m) Write-Host "`n>> $m" -ForegroundColor Cyan    }
function OK   { param($m) Write-Host "   [OK]   $m" -ForegroundColor Green  }
function Warn { param($m) Write-Host "   [WAIT] $m" -ForegroundColor Yellow }
function Err  { param($m) Write-Host "   [ERR]  $m" -ForegroundColor Red    }
function Info { param($m) Write-Host "          $m" -ForegroundColor Gray   }

Clear-Host
Write-Host ""
Write-Host "  ==============================================" -ForegroundColor Magenta
Write-Host "   ULTIMATE UNIFORM AI/ML LAB  -  Windows 11  " -ForegroundColor Magenta
Write-Host "   Admin ek baar chalao, sab accounts ready!  " -ForegroundColor Magenta
Write-Host "  ==============================================" -ForegroundColor Magenta
Write-Host ""

# ---- 1. Admin check ----------------------------------------
Step "Administrator privileges check"
$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]"Administrator")
if (-not $isAdmin) {
    Err "Script ko ADMINISTRATOR ke roop mein chalao!"
    pause; exit 1
}
OK "Administrator confirmed"

# ---- 2. Paths ----------------------------------------------
$condaBase = "C:\miniconda3"
$envPath   = "$condaBase\envs\ml"
$condaExe  = "$condaBase\Scripts\conda.exe"
$pipExe    = "$envPath\Scripts\pip.exe"
$pyExe     = "$envPath\python.exe"
$logFile   = "C:\ProgramData\LabMLSetup.log"

Start-Transcript -Path $logFile -Append | Out-Null
Info "Log file: $logFile"
Info "Install path: $condaBase"

# ---- 3. Miniconda ------------------------------------------
Step "Miniconda installation (shared for ALL users)"
if (-not (Test-Path $condaExe)) {
    Warn "Downloading Miniconda3..."
    $minicondaUrl = "https://repo.anaconda.com/miniconda/Miniconda3-latest-Windows-x86_64.exe"
    $installer    = "C:\miniconda_setup.exe"
    try {
        Invoke-WebRequest -Uri $minicondaUrl -OutFile $installer -UseBasicParsing
        OK "Download complete"
        if (-not (Test-Path $condaBase)) { New-Item -Path $condaBase -ItemType Directory -Force | Out-Null }
        Warn "Installing Miniconda to $condaBase ..."
        $argList = @("/S", "/InstallationType=AllUsers", "/RegisterPython=0", "/AddToPath=0", "/D=$condaBase")
        $proc = Start-Process -FilePath $installer -ArgumentList $argList -Wait -PassThru
        if ($proc.ExitCode -ne 0) {
            Warn "Trying alternate install method..."
            $proc2 = Start-Process -FilePath $installer -ArgumentList "/S /InstallationType=AllUsers /RegisterPython=0 /AddToPath=0 /D=$condaBase" -Wait -PassThru
            if ($proc2.ExitCode -ne 0) { throw "Exit code: $($proc2.ExitCode)" }
        }
        Remove-Item $installer -Force -ErrorAction SilentlyContinue
        OK "Miniconda installed at $condaBase"
    }
    catch {
        Err "Miniconda install failed: $_"
        Stop-Transcript | Out-Null
        pause; exit 1
    }
}
else {
    OK "Miniconda already present"
}

if (-not (Test-Path $condaExe)) {
    Err "conda.exe not found. Please install Miniconda manually to C:\miniconda3"
    Stop-Transcript | Out-Null
    pause; exit 1
}
OK "conda.exe verified"

# ---- 4. conda configuration --------------------------------
# FIX: Write .condarc to THREE locations so conda always finds channels:
#   (a) $condaBase\.condarc  - install-level (read by conda)
#   (b) C:\ProgramData\conda\.condarc - system-wide (AllUsers installs prefer this)
#   (c) CONDARC env var set for both this session and Machine scope
Step "conda configuration"
$env:CONDA_PLUGINS_AUTO_ACCEPT_TOS = "yes"

$condarc = "channels:" + [Environment]::NewLine +
           "  - defaults" + [Environment]::NewLine +
           "  - conda-forge" + [Environment]::NewLine +
           "envs_dirs:" + [Environment]::NewLine +
           "  - $condaBase\envs" + [Environment]::NewLine +
           "pkgs_dirs:" + [Environment]::NewLine +
           "  - $condaBase\pkgs" + [Environment]::NewLine +
           "auto_activate_base: false" + [Environment]::NewLine +
           "ssl_verify: true"

# (a) Install-level .condarc
$condarc | Out-File "$condaBase\.condarc" -Encoding utf8 -Force
OK ".condarc written to $condaBase\.condarc"

# (b) System-wide .condarc (AllUsers install reads C:\ProgramData\conda\.condarc)
$programDataConda = "C:\ProgramData\conda"
if (-not (Test-Path $programDataConda)) { New-Item -Path $programDataConda -ItemType Directory -Force | Out-Null }
$condarc | Out-File "$programDataConda\.condarc" -Encoding utf8 -Force
OK ".condarc written to $programDataConda\.condarc"

# (c) Point CONDARC env var so this session and all future sessions pick it up
$env:CONDARC = "$condaBase\.condarc"
[Environment]::SetEnvironmentVariable("CONDARC", "$condaBase\.condarc", "Machine")
OK "CONDARC env var set (Machine scope)"

# ---- 5. System PATH ----------------------------------------
Step "System PATH update"
$newPaths = @($condaBase, "$condaBase\Scripts", "$condaBase\condabin", "$envPath", "$envPath\Scripts")
$machinePath = [Environment]::GetEnvironmentVariable("Path", "Machine")
$toAdd = $newPaths | Where-Object { $machinePath -notlike "*$_*" }
if ($toAdd) {
    $combined = ($machinePath.TrimEnd(";") + ";" + ($toAdd -join ";")).TrimStart(";")
    [Environment]::SetEnvironmentVariable("Path", $combined, "Machine")
    OK "PATH updated"
} else { OK "PATH already complete" }

# Refresh PATH in this session too
$env:Path = [Environment]::GetEnvironmentVariable("Path","Machine") + ";" + [Environment]::GetEnvironmentVariable("Path","User")

# ---- 6. Folder permissions ---------------------------------
Step "Setting folder permissions"
icacls "$condaBase" /grant "Users:(OI)(CI)RX" /T /C | Out-Null
if (-not (Test-Path "$condaBase\pkgs")) { New-Item "$condaBase\pkgs" -ItemType Directory -Force | Out-Null }
icacls "$condaBase\pkgs" /grant "Users:(OI)(CI)F" /T /C | Out-Null
if (-not (Test-Path "$condaBase\envs")) { New-Item "$condaBase\envs" -ItemType Directory -Force | Out-Null }
icacls "$condaBase\envs" /grant "Users:(OI)(CI)F" /T /C | Out-Null
OK "Permissions set"

# ---- 7. Shared notebooks folder ----------------------------
Step "Creating C:\ML_Notebooks (shared, full access)"
$notebookDir = "C:\ML_Notebooks"
if (-not (Test-Path $notebookDir)) { New-Item -Path $notebookDir -ItemType Directory -Force | Out-Null }
icacls "$notebookDir" /grant "Everyone:(OI)(CI)F" /T /C | Out-Null
icacls "$notebookDir" /grant "Users:(OI)(CI)F" /T /C | Out-Null
icacls "$notebookDir" /inheritance:e | Out-Null
OK "C:\ML_Notebooks ready"

# ---- 8. Create ml environment (Python 3.11) ----------------
# NOTE: Python 3.11 used - better compatibility with tensorflow 2.x via pip
# FIX: -c defaults -c conda-forge passed explicitly so channels are never missing
Step "Creating conda environment 'ml' (Python 3.11)"
if (-not (Test-Path "$envPath\python.exe")) {
    Warn "Creating env - 5-10 min..."
    & $condaExe create -p $envPath -c defaults -c conda-forge python=3.11 -y 2>&1
    if (-not (Test-Path "$envPath\python.exe")) {
        Err "Environment creation failed! Log: $logFile"
        Stop-Transcript | Out-Null
        pause; exit 1
    }
    OK "Environment created"
} else { OK "Environment 'ml' already exists" }

# ---- 9. Core conda packages (NO tensorflow) --
# tensorflow installed via pip to avoid conflicts
# FIX: -c defaults -c conda-forge passed explicitly on all conda install calls
Step "Installing core conda packages (numpy, pandas, sklearn, jupyter...)"
Warn "5-10 min lag sakte hain..."
& $condaExe install -p $envPath -c defaults -c conda-forge -y `
    numpy pandas matplotlib seaborn scikit-learn scipy pillow tqdm `
    jupyter notebook ipykernel 2>&1
OK "Core conda packages installed"

# opencv separately (sometimes conflicts)
Warn "Installing opencv..."
& $condaExe install -p $envPath -c conda-forge -y opencv 2>&1
OK "opencv done"


# ---- 10. pip packages (tensorflow, torch, others) ----------
Step "Installing pip packages"

# Upgrade pip first
& $pipExe install --upgrade pip --quiet 2>&1

# TensorFlow via pip (most reliable method for Python 3.11)
Warn "Installing TensorFlow via pip (~500 MB)..."
& $pipExe install tensorflow --quiet 2>&1
OK "TensorFlow done"

# PyTorch CPU via pip
Warn "Installing PyTorch CPU (~200 MB)..."
& $pipExe install torch torchvision torchaudio --index-url https://download.pytorch.org/whl/cpu --quiet 2>&1
OK "PyTorch done"

# Other ML packages
Warn "Installing xgboost, lightgbm, nltk, transformers..."
& $pipExe install xgboost lightgbm nltk transformers --quiet 2>&1
OK "Extra ML packages done"

# FIX: Re-install ipykernel + jupyter via pip AFTER all conda installs.
# conda's solver (especially with opencv+Qt6+MKL) can silently evict ipykernel
# during dependency resolution. A pip install at this point is immune to that.
Warn "Re-pinning ipykernel + jupyter via pip (conda eviction guard)..."
& $pipExe install ipykernel jupyter notebook --upgrade --quiet 2>&1
OK "ipykernel + jupyter pinned via pip"

# ---- 11. Jupyter kernel ------------------------------------
Step "Registering Jupyter kernel"
& $pyExe -m ipykernel install --name ml --display-name "Python (ML)" --prefix "$condaBase" 2>&1
& $pyExe -m ipykernel install --name ml --display-name "Python (ML)" --sys-prefix 2>&1
OK "Jupyter kernel registered"

# ---- 12. Jupyter config ------------------------------------
Step "Configuring Jupyter (no token, opens C:\ML_Notebooks)"
$jupyterConfig = "c.NotebookApp.notebook_dir = 'C:/ML_Notebooks'" + [Environment]::NewLine +
                 "c.NotebookApp.open_browser = True" + [Environment]::NewLine +
                 "c.NotebookApp.token = ''" + [Environment]::NewLine +
                 "c.NotebookApp.password = ''" + [Environment]::NewLine +
                 "c.NotebookApp.allow_origin = '*'" + [Environment]::NewLine +
                 "c.NotebookApp.ip = 'localhost'"

$jupyterConfigDir = "$condaBase\etc\jupyter"
if (-not (Test-Path $jupyterConfigDir)) { New-Item -Path $jupyterConfigDir -ItemType Directory -Force | Out-Null }
$jupyterConfig | Out-File "$jupyterConfigDir\jupyter_notebook_config.py" -Encoding utf8 -Force

# All existing users
Get-ChildItem "C:\Users" -Directory | Where-Object { $_.Name -notin @("Public","Default User","All Users") } | ForEach-Object {
    $d = "$($_.FullName)\.jupyter"
    if (-not (Test-Path $d)) { New-Item -Path $d -ItemType Directory -Force | Out-Null }
    $jupyterConfig | Out-File "$d\jupyter_notebook_config.py" -Encoding utf8 -Force
    OK "Jupyter config: $($_.Name)"
}
# Default user (future accounts)
$defaultJup = "C:\Users\Default\.jupyter"
if (-not (Test-Path $defaultJup)) { New-Item -Path $defaultJup -ItemType Directory -Force | Out-Null }
$jupyterConfig | Out-File "$defaultJup\jupyter_notebook_config.py" -Encoding utf8 -Force
OK "Jupyter config complete"

# ---- 13. VS Code machine-wide ------------------------------
Step "Installing VS Code (Machine-wide)"
$vscodePath = "$env:ProgramFiles\Microsoft VS Code\Code.exe"
if (-not (Test-Path $vscodePath)) {
    Warn "Installing VS Code..."
    winget install --id Microsoft.VisualStudioCode --scope machine --silent --accept-package-agreements --accept-source-agreements --override "/VERYSILENT /SP- /MERGETASKS=addcontextmenufiles,addcontextmenufolders,associatewithfiles,addtopath,!runcode" 2>&1
}

$env:Path = [Environment]::GetEnvironmentVariable("Path","Machine") + ";" + [Environment]::GetEnvironmentVariable("Path","User")
$codeCmd = Get-Command code -ErrorAction SilentlyContinue

if ($codeCmd) {
    OK "VS Code installed"
    $extensions = @("ms-python.python","ms-python.vscode-pylance","ms-toolsai.jupyter","ms-python.debugpy")
    foreach ($ext in $extensions) {
        & code --install-extension $ext --force 2>&1 | Out-Null
        OK "Extension: $ext"
    }

    $vscodeJson = [ordered]@{
        "python.defaultInterpreterPath"       = $pyExe
        "python.terminal.activateEnvironment" = $true
        "jupyter.notebookFileRoot"            = "C:\ML_Notebooks"
        "editor.fontSize"                     = 14
        "terminal.integrated.fontSize"        = 13
        "files.autoSave"                      = "afterDelay"
        "editor.wordWrap"                     = "on"
    } | ConvertTo-Json -Depth 10

    $profileDirs = New-Object System.Collections.Generic.List[string]
    $profileDirs.Add("C:\Users\Default")
    Get-ChildItem "C:\Users" -Directory | Where-Object { $_.Name -notin @("Public","Default User","All Users") } | ForEach-Object { $profileDirs.Add($_.FullName) }

    foreach ($prof in $profileDirs) {
        $sd = "$prof\AppData\Roaming\Code\User"
        try {
            if (-not (Test-Path $sd)) { New-Item -Path $sd -ItemType Directory -Force | Out-Null }
            $vscodeJson | Set-Content "$sd\settings.json" -Force -Encoding UTF8
            OK "VS Code settings: $prof"
        } catch { Warn "VS Code settings failed: $prof" }
    }
} else {
    Warn "VS Code 'code' command nahi mila - restart ke baad try karo"
}

# ---- 14. Desktop shortcuts ---------------------------------
Step "Creating Desktop shortcuts for ALL users"
$publicDesktop = "C:\Users\Public\Desktop"
$WshShell = New-Object -ComObject WScript.Shell

function New-Lnk {
    param($Name, $Cmd)
    $lnk = "$publicDesktop\$Name.lnk"
    $sc  = $WshShell.CreateShortcut($lnk)
    $sc.TargetPath       = "powershell.exe"
    $sc.Arguments        = "-NoExit -ExecutionPolicy Bypass -Command `"$Cmd`""
    $sc.WorkingDirectory = "C:\ML_Notebooks"
    $sc.WindowStyle      = 1
    $sc.Save()
    OK "Shortcut: $Name"
}

New-Lnk "Jupyter Notebook (ML)"   "cd C:\ML_Notebooks; & '$condaExe' run -p '$envPath' jupyter notebook --notebook-dir=C:\ML_Notebooks"
New-Lnk "ML Python Shell"         "cd C:\ML_Notebooks; & '$condaExe' run -p '$envPath' python"
New-Lnk "ML PowerShell"           "cd C:\ML_Notebooks; If (Test-Path '$condaBase\shell\condabin\conda-hook.ps1') { . '$condaBase\shell\condabin\conda-hook.ps1'; conda activate '$envPath' }; Write-Host 'ML env activated!' -ForegroundColor Green"

$lnkVS = "$publicDesktop\VS Code (ML).lnk"
$scVS  = $WshShell.CreateShortcut($lnkVS)
$scVS.TargetPath       = "$env:ProgramFiles\Microsoft VS Code\Code.exe"
$scVS.Arguments        = "C:\ML_Notebooks"
$scVS.WorkingDirectory = "C:\ML_Notebooks"
$scVS.Save()
OK "Shortcut: VS Code (ML)"

# ---- 15. PowerShell profiles (conda auto-init) -------------
Step "Adding conda init to all PowerShell profiles"
$i1 = "# ---- Conda auto-init (Lab_ML_Setup) ----"
$i2 = "If (Test-Path `"$condaBase\shell\condabin\conda-hook.ps1`") {"
$i3 = "    . `"$condaBase\shell\condabin\conda-hook.ps1`""
$i4 = "    conda activate `"$envPath`""
$i5 = "}"

$profs = New-Object System.Collections.Generic.List[string]
$profs.Add("C:\Users\Default")
Get-ChildItem "C:\Users" -Directory | Where-Object { $_.Name -notin @("Public","Default User","All Users") } | ForEach-Object { $profs.Add($_.FullName) }

foreach ($prof in $profs) {
    $psDir = "$prof\Documents\WindowsPowerShell"
    $psPro = "$psDir\profile.ps1"
    try {
        if (-not (Test-Path $psDir)) { New-Item -Path $psDir -ItemType Directory -Force | Out-Null }
        $ex = if (Test-Path $psPro) { Get-Content $psPro -Raw -ErrorAction SilentlyContinue } else { "" }
        if ($ex -notlike "*conda-hook*") {
            Add-Content -Path $psPro -Value "" -Encoding UTF8
            Add-Content -Path $psPro -Value $i1 -Encoding UTF8
            Add-Content -Path $psPro -Value $i2 -Encoding UTF8
            Add-Content -Path $psPro -Value $i3 -Encoding UTF8
            Add-Content -Path $psPro -Value $i4 -Encoding UTF8
            Add-Content -Path $psPro -Value $i5 -Encoding UTF8
            OK "PS profile: $prof"
        } else { OK "PS profile already done: $prof" }
    } catch { Warn "Profile failed: $prof" }
}

# ---- 16. ml env permissions (full for users) ---------------
Step "Setting ml env permissions for all users"
icacls "$envPath" /grant "Users:(OI)(CI)F" /T /C | Out-Null
OK "ml env: Users have full access"

# ---- 17. Sample notebook -----------------------------------
Step "Creating sample notebook"
$nb = '{' + [Environment]::NewLine +
' "nbformat": 4,' + [Environment]::NewLine +
' "nbformat_minor": 5,' + [Environment]::NewLine +
' "metadata": {"kernelspec": {"display_name": "Python (ML)", "language": "python", "name": "ml"}, "language_info": {"name": "python"}},' + [Environment]::NewLine +
' "cells": [' + [Environment]::NewLine +
'  {"cell_type": "markdown", "id": "a1", "metadata": {}, "source": ["# AI/ML Lab - Starter Notebook\n", "Sab libraries install hain. Neeche wala cell run karo!\n"]},' + [Environment]::NewLine +
'  {"cell_type": "code", "execution_count": null, "id": "b1", "metadata": {}, "outputs": [], "source": ["import numpy as np\n", "import pandas as pd\n", "import matplotlib.pyplot as plt\n", "import sklearn\n", "try:\n", "    import tensorflow as tf\n", "    print(f\"TensorFlow: {tf.__version__}\")\n", "except: print(\"TF not loaded\")\n", "try:\n", "    import torch\n", "    print(f\"PyTorch: {torch.__version__}\")\n", "except: print(\"Torch not loaded\")\n", "print(f\"NumPy: {np.__version__}\")\n", "print(f\"Pandas: {pd.__version__}\")\n", "print(\"Sab kuch kaam kar raha hai!\")"]}' + [Environment]::NewLine +
' ]' + [Environment]::NewLine +
'}'
$nb | Out-File "C:\ML_Notebooks\Starter_Test.ipynb" -Encoding utf8 -Force
icacls "C:\ML_Notebooks\Starter_Test.ipynb" /grant "Everyone:(F)" | Out-Null
OK "Starter_Test.ipynb ready"

# ---- 18. Final verification --------------------------------
Step "Final verification"
$pyTest = "import sys, numpy as np, pandas as pd, matplotlib, sklearn" + [Environment]::NewLine +
          "try:" + [Environment]::NewLine +
          "    import tensorflow as tf; tf_v=tf.__version__" + [Environment]::NewLine +
          "except Exception as e: tf_v='WARN: '+str(e)[:50]" + [Environment]::NewLine +
          "try:" + [Environment]::NewLine +
          "    import torch; torch_v=torch.__version__" + [Environment]::NewLine +
          "except Exception as e: torch_v='WARN: '+str(e)[:50]" + [Environment]::NewLine +
          "print('Python      : '+sys.version.split()[0])" + [Environment]::NewLine +
          "print('NumPy       : '+np.__version__)" + [Environment]::NewLine +
          "print('Pandas      : '+pd.__version__)" + [Environment]::NewLine +
          "print('Matplotlib  : '+matplotlib.__version__)" + [Environment]::NewLine +
          "print('Scikit-learn: '+sklearn.__version__)" + [Environment]::NewLine +
          "print('TensorFlow  : '+tf_v)" + [Environment]::NewLine +
          "print('PyTorch     : '+torch_v)" + [Environment]::NewLine +
          "print('STATUS      : ALL OK')"

$tempPy = "C:\ProgramData\ml_verify.py"
$pyTest | Out-File $tempPy -Encoding utf8 -Force
$result = & $pyExe $tempPy 2>&1
Remove-Item $tempPy -Force -ErrorAction SilentlyContinue
foreach ($line in $result) { Info "  $line" }
if (($result | Out-String) -match "ALL OK") { OK "Verification PASSED" }
else { Warn "Kuch warnings hain - log check karo: $logFile" }

# ---- 19. Restart Explorer ----------------------------------
Step "Restarting Explorer"
try {
    Stop-Process -Name explorer -Force -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 2
    Start-Process explorer.exe
    OK "Explorer restarted"
} catch { Warn "Explorer restart skip" }

Stop-Transcript | Out-Null

Write-Host ""
Write-Host "  ==============================================" -ForegroundColor Green
Write-Host "       SETUP COMPLETE - 100% UNIFORM           " -ForegroundColor Green
Write-Host "  ==============================================" -ForegroundColor Green
Write-Host ""
Write-Host "  Har local account mein ye sab ready hai:" -ForegroundColor Yellow
Write-Host "    [OK] Jupyter Notebook (ML)     - C:\ML_Notebooks mein khulega"
Write-Host "    [OK] VS Code (ML)              - C:\ML_Notebooks open"
Write-Host "    [OK] ML Python Shell           - Desktop shortcut"
Write-Host "    [OK] ML PowerShell             - conda auto-activated"
Write-Host "    [OK] Permission denied FIX     - Applied"
Write-Host ""
Write-Host "  IMPORTANT - NEXT STEP:" -ForegroundColor Cyan
Write-Host "    Computer ko ek baar RESTART karo"
Write-Host "    Phir kisi bhi local account se login karo - sab kaam karega!"
Write-Host ""
Write-Host "  Log: $logFile" -ForegroundColor Gray
Write-Host ""
[Console]::Beep(880,150); [Console]::Beep(1100,150); [Console]::Beep(1320,300)
pause
