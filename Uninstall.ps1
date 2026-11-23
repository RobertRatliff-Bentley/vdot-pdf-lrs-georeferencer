# Removes the VDOT Project Mapper macros and .cfg (Python libraries are left installed).
Get-ChildItem "C:\ProgramData\Bentley\*\Configuration\Organization" -Directory -ErrorAction SilentlyContinue | ForEach-Object {
    foreach ($f in @("Macros\vdot_project_mapper.py", "Macros\VDOT_Project_Mapper.mvba", "Macros\VDOT_Project_Mapper_Launcher.mvba", "vdot_project_mapper.cfg",
                    "Macros\vdot_pdf_lrs_clipper.py", "Macros\VDOT_Georef.mvba", "Macros\VDOT_Georef_Native.mvba", "vdot_pdf_georef.cfg")) {
        $p = Join-Path $_.FullName $f
        if (Test-Path $p) { Remove-Item $p -Force; Write-Host "Removed $p" }
    }
}
