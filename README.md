# VDOT PDF LRS Georeferencer for MicroStation / OpenRoads Designer

Georeferences VDOT county PDF maps in MicroStation 2026 / OpenRoads Designer 25 using the VDOT LRS ArcGIS REST service: auto-aligns the PDF to the county boundary, clips it, and places the map image, boundary, per-class LRS route levels, scale bar and north arrow.

## Versions
- `vdot_pdf_lrs_clipper.py` - Python / PyQt5 tool (run with `python load vdot_pdf_lrs_clipper.py`).
- `VDOT_Georef_Native.mvba` - native VBA version, no Python needed (`vba run [VDOT_Georef_Native]modMain.Main`). Source in `VDOT_Georef_Native_src/`.
- `VDOT_Georef.mvba` / `VdotGeoref.bas` - legacy VBA launcher for the Python tool.

## Install
Download the offline installer zip (includes all Python wheels) from **Releases**, extract, and run `Install.bat`. See `README.txt` for details.
