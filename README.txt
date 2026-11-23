VDOT Project Mapper for MicroStation
Created by Robert Ratliff

REQUIREMENTS
  - MicroStation 2024+ or OpenRoads Designer 2025 (Python 3.12, shared PowerPlatformPython)
    The installer also puts the macro in C:\ProgramData\Bentley\OpenRoads Designer 2025.00\Configuration\Organization\Macros
  - Internet access when running the tool (VDOT LRS / boundary REST services)

INSTALL (no internet needed)
  1. Copy this whole folder to the computer.
  2. Double-click Install.bat and allow Administrator access.
     - Installs the bundled Python libraries (PyMuPDF, numpy, pyproj, shapely and their
       dependencies) into C:\ProgramData\Bentley\PowerPlatformPython.
     - Copies vdot_project_mapper.py and the VBA macro VDOT_Project_Mapper_Launcher.mvba into every
       C:\ProgramData\Bentley\<product>\Configuration\Organization\Macros folder and adds
       vdot_project_mapper.cfg so the macro is found from any WorkSet.
  3. Restart MicroStation.

  Custom/network Organization folder:
     powershell -ExecutionPolicy Bypass -File Install.ps1 -OrganizationPath "\\server\Bentley\Organization"

RUN
  Open a DGN with a geographic coordinate system and key in:
     python load vdot_project_mapper.py
  or run the VBA macro (same window and behaviour - it starts the Python tool):
     vba run [VDOT_Project_Mapper_Launcher]VdotProjectMapper.Main
     (or Utilities > Macro > Macros... > VDOT_Project_Mapper_Launcher > Main; this key-in can be put on a
      toolbar button or function key)
  VdotProjectMapper.bas is the VBA source of the macro.

NATIVE VBA VERSION (no Python required)
  VDOT_Project_Mapper.mvba is a stand-alone VBA rewrite of the tool with the same window,
  cards, PDF viewer tools (pan, adjust, tie points, block fence), auto-align, routes,
  high-resolution masked PNG tiles + hidden PDF placement, scale bar and north arrow. It does not need Python
  or the bundled libraries - it only needs MicroStation/OpenRoads and internet access.
     vba run [VDOT_Project_Mapper]modMain.Main
  Differences: the viewer zooms with the toolbar buttons (no mouse wheel); page 1 of the PDF
  is used. VDOT_Project_Mapper_src\ holds the VBA source modules.
  Log: %TEMP%\vdot_georef_native.log

UNINSTALL
  powershell -ExecutionPolicy Bypass -File Uninstall.ps1

LOGS
  %TEMP%\vdot_georef_install.log   (installer)
  %TEMP%\vdot_georef.log           (tool)

VDOT districts
  The type drop-down in the Jurisdiction card also lists 'Districts' (the 9 VDOT district
  boundaries from the VDOTAdministrativeBoundaries REST service) and one entry per district
  (Bristol, Salem, Lynchburg, Richmond, Hampton Roads, Fredericksburg, Culpeper, Staunton,
  Northern Virginia). Picking a district lists its boundary first, then its counties, cities
  and towns; load the district boundary to clip/align to a whole district map.

Scale bar & north arrow
  Tick 'Scale bar & north arrow' in the Output card to place a round-length scale bar
  (feet/miles or meters/km) and a true-north arrow below the county's lower-right corner,
  on the 'Annotation level' (default VDOT_Map_Annotation).

LRS route levels
  Routes are placed on one level per LRS class. The 'Routes level' box is the name prefix
  (default VDOT_LRS_Routes). Each level is created with by-level colour/weight:
    VDOT_LRS_Routes_Interstate       colour 3 (red)     weight 4
    VDOT_LRS_Routes_US_Primary       colour 1 (blue)    weight 3
    VDOT_LRS_Routes_State_Primary    colour 4 (yellow)  weight 3
    VDOT_LRS_Routes_Secondary        colour 2 (green)   weight 1
    VDOT_LRS_Routes_Ramps_Frontage   colour 6 (orange)  weight 1
    VDOT_LRS_Routes_Urban_Streets    colour 5 (magenta) weight 1
    VDOT_LRS_Routes_Other            colour 8 (grey)    weight 0
  Use Level Display to switch route classes on/off. Same in Python and native VBA versions.

Block fence (PDF viewer)
  Click '▭ Fence' in the PDF viewer toolbar and drag a rectangle on the map. After the map is aligned,
  Place clips the map image, LRS routes and scale-bar position to that block instead of the county
  boundary, and the block also becomes the active MicroStation fence. '✕ Clear fence' returns to county.
