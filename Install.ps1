<#
  VDOT Project Mapper for MicroStation - offline installer
  Created by Robert Ratliff

  1. Installs the bundled Python libraries (wheels\) into Bentley's PowerPlatform Python (no internet needed).
  2. Copies vdot_project_mapper.py and VDOT_Project_Mapper.mvba into <Organization>\Macros\ for every MicroStation-based product found
     (or the one given with -OrganizationPath) and adds a small .cfg so MicroStation finds it from any WorkSet.

  Usage (Install.bat runs this elevated):
     powershell -ExecutionPolicy Bypass -File Install.ps1 [-PythonExe <path>] [-OrganizationPath <path>]
#>
param(
    [string]$PythonExe = "",
    [string[]]$OrganizationPath = @()
)
$ErrorActionPreference = "Stop"
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$script = Join-Path $here "vdot_project_mapper.py"
$mvbaNative = Join-Path $here "VDOT_Project_Mapper.mvba"
$wheels = Join-Path $here "wheels"
$req = Join-Path $here "requirements.txt"
$log = Join-Path $env:TEMP "vdot_georef_install.log"
Start-Transcript -Path $log -Force | Out-Null

function Step($m) { Write-Host ""; Write-Host "== $m" -ForegroundColor Cyan }
function Ok($m)   { Write-Host "   OK  $m" -ForegroundColor Green }
function Warn($m) { Write-Host "   !!  $m" -ForegroundColor Yellow }
function Fail($m) { Write-Host "   XX  $m" -ForegroundColor Red; Stop-Transcript | Out-Null; exit 1 }

try {
    # ---------------------------------------------------------------- Python
    Step "Locating Bentley Python"
    if (-not $PythonExe) {
        $cands = @("C:\ProgramData\Bentley\PowerPlatformPython\python\python.exe") +
                 @(Get-ChildItem "C:\ProgramData\Bentley\*\python\python.exe" -ErrorAction SilentlyContinue | ForEach-Object FullName)
        $PythonExe = $cands | Where-Object { Test-Path $_ } | Select-Object -First 1
    }
    if (-not $PythonExe -or -not (Test-Path $PythonExe)) {
        Fail "Bentley PowerPlatform Python not found. Install MicroStation 2024+ (with Python) or pass -PythonExe."
    }
    $ver = (& $PythonExe -c "import sys;print('%d.%d' % sys.version_info[:2])").Trim()
    Ok "$PythonExe (Python $ver)"
    if ($ver -ne "3.12") { Fail "The bundled libraries are for Python 3.12; this machine has Python $ver." }

    # ---------------------------------------------------------------- Libraries
    Step "Installing Python libraries from the package (offline)"
    & $PythonExe -m pip install --no-index --find-links "$wheels" -r "$req" --disable-pip-version-check
    if ($LASTEXITCODE -ne 0) { Fail "pip install failed (see $log). Try running Install.bat as Administrator." }

    Step "Verifying libraries"
    $check = @'
import importlib, warnings
warnings.simplefilter("ignore")
bad = []
for m in ("pymupdf", "numpy", "pyproj", "shapely", "PyQt5.QtWidgets"):
    try:
        importlib.import_module(m); print("   OK  " + m)
    except Exception as ex:
        bad.append(m); print("   XX  %s: %s" % (m, ex))
raise SystemExit(1 if bad else 0)
'@
    $check | & $PythonExe -
    if ($LASTEXITCODE -ne 0) { Fail "Some libraries could not be imported (PyQt5 ships with MicroStation's Python)." }

    # ---------------------------------------------------------------- Macro
    Step "Installing the MicroStation macro"
    if (-not $OrganizationPath -or $OrganizationPath.Count -eq 0) {
        $OrganizationPath = @(Get-ChildItem "C:\ProgramData\Bentley\*\Configuration\Organization" -Directory -ErrorAction SilentlyContinue |
                              ForEach-Object FullName)
    }
    if ($OrganizationPath.Count -eq 0) { Fail "No MicroStation Configuration\Organization folder found. Pass -OrganizationPath." }
    $cfgText = @"
#----------------------------------------------------------------------
# VDOT Project Mapper (Robert Ratliff) - lets 'python load' find
# vdot_project_mapper.py from any WorkSet.
#----------------------------------------------------------------------
MS_PYTHONSEARCHDIRECTORIES > `$(_USTN_ORGANIZATION)Macros/
MS_PYTHONSCRIPTS           > `$(_USTN_ORGANIZATION)Macros/
MS_VBASEARCHDIRECTORIES    > `$(_USTN_ORGANIZATION)Macros/
"@
    foreach ($org in $OrganizationPath) {
        $mac = Join-Path $org "Macros"
        New-Item -ItemType Directory -Force $mac | Out-Null
        Copy-Item $script (Join-Path $mac "vdot_project_mapper.py") -Force
        if (Test-Path $mvbaNative) { Copy-Item $mvbaNative (Join-Path $mac "VDOT_Project_Mapper.mvba") -Force }
        [IO.File]::WriteAllText((Join-Path $org "vdot_project_mapper.cfg"), $cfgText)
        # Remove copies installed under the old "VDOT Map Georeferencer" names.
        foreach ($old in @("Macros\vdot_pdf_lrs_clipper.py", "Macros\VDOT_Georef.mvba", "Macros\VDOT_Georef_Native.mvba", "Macros\VDOT_Project_Mapper_Launcher.mvba", "vdot_pdf_georef.cfg")) {
            $op = Join-Path $org $old
            if (Test-Path $op) { Remove-Item $op -Force -ErrorAction SilentlyContinue }
        }
        Ok $mac
    }

    Step "Done"
    Write-Host @"
   Restart MicroStation, open a DGN that has a coordinate system, then key in:
       python load vdot_project_mapper.py
   or the native VBA version (no Python needed):
       vba run [VDOT_Project_Mapper]modMain.Main
   Install log: $log
"@
}
catch {
    Fail $_.Exception.Message
}
Stop-Transcript | Out-Null
exit 0
