# VDOT Project Mapper for MicroStation / OpenRoads Designer

Georeferences VDOT county PDF maps in MicroStation 2026 / OpenRoads Designer 25 using the VDOT LRS ArcGIS REST service: auto-aligns the PDF to the county boundary, clips it, and places the map image, boundary, per-class LRS route levels, scale bar and north arrow. The map image is rendered at a selectable **Image resolution** (150-400 DPI; 300 DPI default in Python, 200 DPI in VBA) as georeferenced PNG tiles cropped to the boundary, so large district maps stay sharp when zoomed in. **Delete placed drawing & rasters** removes everything on the tool's levels in the active model and detaches its raster attachments (undoable). **Get VDOT county maps** opens the [VDOT maps website](https://www.vdot.virginia.gov/travel-traffic/maps/) to download county PDFs. The jurisdiction drop-down also offers the **9 VDOT districts** (Bristol, Salem, Lynchburg, Richmond, Hampton Roads, Fredericksburg, Culpeper, Staunton, Northern Virginia) from the VDOT administrative boundaries REST service - pick a district to load its boundary or list its counties, cities and towns.

**LRS item types:** placed elements carry MicroStation Item Types (library `VDOT_LRS`) with LRS attributes - `LRS_Route` on every route (RouteName, CommonName, Category, RouteClass, RouteNumber, RouteType, Direction, Jurisdiction, RouteID, LRS_OBJECTID, Source) and `LRS_Boundary` on the boundary (Name, FullName, JurisType, VDOTCode, FIPS, District, LRS_OBJECTID, Source). Toggle with the **LRS item types** checkbox.

## Demo

https://github.com/user-attachments/assets/72b5b800-0210-4505-81ff-6975d6578522

## Versions
- `vdot_project_mapper.py` - Python / PyQt5 tool (run with `python load vdot_project_mapper.py`).
- `VDOT_Project_Mapper.mvba` - native VBA version, no Python needed (`vba run [VDOT_Project_Mapper]modMain.Main`). Source in `VDOT_Project_Mapper_src/`.

## Install
Download the offline installer zip (includes all Python wheels) from **Releases**, extract, and run `Install.bat`. See `README.txt` for details.
