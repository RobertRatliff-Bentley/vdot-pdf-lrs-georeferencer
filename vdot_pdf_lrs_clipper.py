# -*- coding: utf-8 -*-
"""
VDOT PDF Map Georeferencer & Clipper for MicroStation
=====================================================
Created by Robert Ratliff  ·  Build 2026-11-18

Loads any VDOT county/city road-map PDF, pulls the jurisdiction boundary and
(optionally) LRS routes from VDOT's ArcGIS REST services, georeferences the PDF
to the boundary (automatic fit + manual tools), and places the result in the
active DGN using the model's geographic coordinate system (GCS):

  * jurisdiction boundary as closed shapes
  * LRS routes as line strings (colored by route category)
  * the original PDF page attached as a raster with the solved transform and
    clipped to the boundary (raster clip, optional hole masks / white transparency)

Requirements
  * MicroStation (CONNECT / 2023+) with Python, a model with a GCS assigned
    (e.g. VA83-SF / EPSG 2284) and internet access to services.arcgis.com.
  * PyQt5 + numpy (shipped with Bentley PowerPlatformPython); PySide6 (QtPdf)
    or PyMuPDF for PDF rendering; shapely is optional (route clipping).

Run
  Key-in:  python load "<path>\\vdot_pdf_lrs_clipper.py"   (or Python Manager)
"""

import sys
import os
import re
import json
import math
import time
import tempfile
import threading
import subprocess
import traceback
import urllib.parse
import urllib.request
import importlib


def _repair_stdlib():
    """MicroStation's Python host can wipe the globals of already-imported standard modules
    (e.g. contextlib.sys becomes None) after scripts are unloaded/reloaded, which breaks
    subprocess. Re-execute any module whose 'sys'/'os' global has been set to None."""
    fixed = []
    for _ in range(2):
        for name, mod in list(sys.modules.items()):
            d = getattr(mod, "__dict__", None)
            if not d or not str(d.get("__file__") or "").endswith((".py", ".pyc")):
                continue
            if any(k in d and d[k] is None for k in ("sys", "os", "_os")):
                try:
                    importlib.reload(mod)
                    fixed.append(name)
                except Exception:
                    pass
    return fixed


_repair_stdlib()

import numpy as np

MSTN_IMPORT_ERROR = ""
# Import order matches Bentley's samples: MSPyMstnPlatform depends on types
# registered by MSPyECObjects/MSPyDgnView; importing it without them fails and
# leaves pybind11 types half-registered until MicroStation restarts.
try:
    from MSPyBentley import *            # noqa: F401,F403
    from MSPyBentleyGeom import *        # noqa: F401,F403
    from MSPyECObjects import *          # noqa: F401,F403
    from MSPyDgnPlatform import *        # noqa: F401,F403
    from MSPyDgnView import *            # noqa: F401,F403
    from MSPyMstnPlatform import *       # noqa: F401,F403
    HAVE_MSTN = True
except Exception as _e:
    HAVE_MSTN = False
    MSTN_IMPORT_ERROR = "%s: %s" % (type(_e).__name__, _e)
    if "already registered" in str(_e):
        MSTN_IMPORT_ERROR += (" - MicroStation's Python modules were left half-loaded by an earlier "
                              "run. Restart MicroStation and run this script again.")

from PyQt5 import QtCore, QtGui, QtWidgets

try:
    from shapely.geometry import Polygon as _ShpPolygon, LineString as _ShpLine, MultiPolygon as _ShpMultiPolygon
    from shapely.prepared import prep as _shp_prep
    HAVE_SHAPELY = True
except Exception:
    HAVE_SHAPELY = False


APP_TITLE = "VDOT Map Georeferencer"
APP_VERSION = "1.0"
APP_AUTHOR = "Robert Ratliff"
BUILD_DATE = "2026-11-19"
APP_CREDIT = "Created by %s  ·  Build %s" % (APP_AUTHOR, BUILD_DATE)
# VDOT logo (vdot.virginia.gov), embedded as PNG so the macro stays a single file.
VDOT_LOGO_PNG_B64 = (
    "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"
)


def vdot_logo_pixmap(height=None):
    pm = QtGui.QPixmap()
    pm.loadFromData(QtCore.QByteArray.fromBase64(VDOT_LOGO_PNG_B64.encode("ascii")), "PNG")
    if height and not pm.isNull():
        pm = pm.scaledToHeight(int(height), QtCore.Qt.SmoothTransformation)
    return pm

JURIS_URL = "https://services.arcgis.com/p5v98VHDX9Atv3l7/arcgis/rest/services/VA_Jurisdictions/FeatureServer/0"
LRS_URL = "https://services.arcgis.com/p5v98VHDX9Atv3l7/arcgis/rest/services/LRS_Route_Master/FeatureServer/0"

JURIS_TYPES = {"CO": "County", "CI": "City", "TO": "Town"}

# group label -> (RTE_CATEGORY_NM values, default on, MicroStation color, weight, preview color)
ROUTE_GROUPS = [
    ("Interstate", ["Interstate"], True, 3, 4, "#ef4444"),
    ("US Primary", ["US Highway Primary"], True, 1, 3, "#3b82f6"),
    ("State Primary", ["State Highway Primary"], True, 4, 3, "#eab308"),
    ("Secondary", ["Secondary"], True, 2, 1, "#22c55e"),
    ("Ramps & Frontage", ["Interstate Ramp", "Interstate Frontage Road", "Non-Interstate Ramp",
                          "Non-Interstate Frontage Road", "Street IS Ramp", "Street Non-IS Ramp"], False, 6, 1, "#f97316"),
    ("Urban & Streets", ["Urban", "Urban Streets", "Streets", "County Roads"], False, 5, 1, "#a855f7"),
    ("Other", ["Federal Lands Roads", "Institutional Property", "Private Property", "School Roads",
               "Shared Use Path"], False, 8, 0, "#94a3b8"),
]


def route_level_name(prefix, group_label):
    """Per-class route level, e.g. VDOT_LRS_Routes_US_Primary."""
    return "%s_%s" % (prefix, group_label.replace(" & ", "_").replace(" ", "_"))

DEFAULT_LEVELS = {
    "boundary": "VDOT_LRS_Boundary",
    "routes": "VDOT_LRS_Routes",
    "raster": "VDOT_PDF_Map",
    "annotation": "VDOT_Map_Annotation",
}

PREVIEW_DPI = 48.0
WIN_ATTR = "_vdot_georef_win"
# Resolution of the masked county image attached in MicroStation (longest side capped).
MASK_IMAGE_DPI = 200.0
MASK_IMAGE_MAX_PX = 10000
# Initial bitmap-size guess for the PDF attachment; corrected from MicroStation after attaching.
PDF_ATTACH_DPI_GUESS = 300.0
BENTLEY_PYTHON = r"C:\ProgramData\Bentley\PowerPlatformPython\python\python.exe"


LOG_PATH = os.path.join(tempfile.gettempdir(), "vdot_georef.log")
_LOG_LOCK = threading.Lock()


def _log_file(msg):
    try:
        with _LOG_LOCK, open(LOG_PATH, "a", encoding="utf-8") as f:
            f.write("%s [%s] %s\n" % (time.strftime("%Y-%m-%d %H:%M:%S"), threading.current_thread().name, msg))
    except Exception:
        pass


def _log_console(msg):
    _log_file(msg)
    # Host consoles (MicroStation) are not thread-safe: only print from the main thread.
    if threading.current_thread() is not threading.main_thread():
        return
    try:
        print("[%s] %s" % (APP_TITLE, msg))
    except Exception:
        pass


def _excepthook(et, ev, tb):
    # PyQt5 aborts the host process (qFatal) on unhandled exceptions in slots unless a hook is set.
    _log_console("Unhandled error:\n" + "".join(traceback.format_exception(et, ev, tb)))
    if threading.current_thread() is threading.main_thread():
        w = getattr(sys, WIN_ATTR, None)
        try:
            if w is not None:
                w.log("Unexpected error: %s" % ev, "err")
        except Exception:
            pass


def _qt_message_handler(mode, ctx, msg):
    names = {QtCore.QtDebugMsg: "debug", QtCore.QtWarningMsg: "warning", QtCore.QtCriticalMsg: "critical",
             QtCore.QtFatalMsg: "FATAL", QtCore.QtInfoMsg: "info"}
    if mode != QtCore.QtDebugMsg:
        _log_file("Qt %s: %s" % (names.get(mode, mode), msg))


def install_safety_hooks():
    sys.excepthook = _excepthook
    try:
        threading.excepthook = lambda a: _log_console(
            "Thread error:\n" + "".join(traceback.format_exception(a.exc_type, a.exc_value, a.exc_traceback)))
    except Exception:
        pass
    try:
        QtCore.qInstallMessageHandler(_qt_message_handler)
    except Exception:
        pass


def ui_action(fn):
    """Wrap a UI entry point: never let an exception reach Qt, and never run re-entrantly from
    inside another operation's event pump (the call is deferred until the UI is idle instead)."""
    nargs = fn.__code__.co_argcount - 1

    def wrapper(self, *args, **kw):
        args = args[:nargs]
        if getattr(self, "_pumping", False):
            QtCore.QTimer.singleShot(150, lambda: wrapper(self, *args, **kw))
            return None
        try:
            return fn(self, *args, **kw)
        except Exception as ex:
            _log_console("%s failed:\n%s" % (fn.__name__, traceback.format_exc()))
            try:
                self.log("%s failed: %s" % (fn.__name__, ex), "err")
            except Exception:
                pass
            return None

    wrapper.__name__ = fn.__name__
    wrapper.__doc__ = fn.__doc__
    return wrapper


# =============================================================================
# REST client (VDOT ArcGIS services)
# =============================================================================
class VdotRestClient:
    def __init__(self, timeout=90):
        self.timeout = timeout

    def _query(self, layer_url, params, retries=3):
        p = dict(params)
        p.setdefault("f", "json")
        data = urllib.parse.urlencode(p).encode("utf-8")
        last = None
        for attempt in range(retries):
            try:
                req = urllib.request.Request(
                    layer_url + "/query", data=data,
                    headers={"User-Agent": "MicroStation-VDOT-Georef/1.0",
                             "Content-Type": "application/x-www-form-urlencoded"})
                with urllib.request.urlopen(req, timeout=self.timeout) as r:
                    js = json.loads(r.read().decode("utf-8"))
                if "error" in js:
                    raise RuntimeError("Service error: %s" % js["error"].get("message", js["error"]))
                return js
            except Exception as ex:
                last = ex
                time.sleep(1.0 + attempt)
        raise RuntimeError("Request failed: %s" % last)

    def list_jurisdictions(self):
        out, offset = [], 0
        while True:
            js = self._query(JURIS_URL, {
                "where": "1=1",
                "outFields": "OBJECTID,NAME,NAMELSAD,JURISTYPE,VDOT_CODE,STCOFIPS,DISTRICT",
                "returnGeometry": "false", "orderByFields": "NAMELSAD",
                "resultOffset": offset, "resultRecordCount": 1000})
            feats = js.get("features", [])
            out.extend(f["attributes"] for f in feats)
            if not js.get("exceededTransferLimit") or not feats:
                break
            offset += len(feats)
        return out

    def get_boundary(self, objectid, out_sr=4269):
        js = self._query(JURIS_URL, {
            "where": "OBJECTID=%d" % int(objectid), "outFields": "OBJECTID,NAMELSAD",
            "returnGeometry": "true", "outSR": out_sr, "geometryPrecision": 8})
        feats = js.get("features", [])
        if not feats:
            raise RuntimeError("Boundary not found (OBJECTID=%s)" % objectid)
        rings = [np.asarray(r, dtype=float)[:, :2] for r in feats[0]["geometry"]["rings"] if len(r) >= 4]
        return rings

    def get_routes(self, bbox, categories, out_sr=4269, progress=None, cancel=None):
        cats = ",".join("'%s'" % c.replace("'", "''") for c in categories)
        where = "RTE_CATEGORY_NM IN (%s)" % cats
        out, offset = [], 0
        while True:
            js = self._query(LRS_URL, {
                "where": where,
                "geometry": "%.8f,%.8f,%.8f,%.8f" % tuple(bbox),
                "geometryType": "esriGeometryEnvelope", "inSR": out_sr,
                "spatialRel": "esriSpatialRelIntersects",
                "outFields": "OBJECTID,RTE_NM,RTE_COMMON_NM,RTE_CATEGORY_NM",
                "returnGeometry": "true", "outSR": out_sr,
                "maxAllowableOffset": 0.00001, "geometryPrecision": 7,
                "orderByFields": "OBJECTID",
                "resultOffset": offset, "resultRecordCount": 2000})
            feats = js.get("features", [])
            for f in feats:
                g = f.get("geometry") or {}
                a = f["attributes"]
                for path in g.get("paths", []):
                    if len(path) >= 2:
                        out.append((a.get("RTE_CATEGORY_NM") or "", a.get("RTE_NM") or "",
                                    np.asarray(path, dtype=float)[:, :2]))
            offset += len(feats)
            if progress:
                progress(offset)
            if cancel and cancel():
                break
            if not js.get("exceededTransferLimit") or not feats:
                break
        return out


def clip_routes_to_rings(routes, rings):
    """Clip route polylines (lon/lat) to the boundary polygon using shapely."""
    if not HAVE_SHAPELY or not rings:
        return routes
    try:
        polys = _rings_to_shapely(rings)
        if polys is None:
            return routes
        pp = _shp_prep(polys)
        out = []
        for cat, name, pts in routes:
            ls = _ShpLine(pts)
            if pp.contains(ls):
                out.append((cat, name, pts))
                continue
            if not pp.intersects(ls):
                continue
            g = polys.intersection(ls)
            parts = getattr(g, "geoms", [g])
            for part in parts:
                if part.geom_type == "LineString" and part.length > 0:
                    out.append((cat, name, np.asarray(part.coords)[:, :2]))
        return out
    except Exception as ex:
        _log_console("Route clipping skipped: %s" % ex)
        return routes


def _ring_area(r):
    x, y = r[:, 0], r[:, 1]
    return 0.5 * float(np.dot(x, np.roll(y, -1)) - np.dot(y, np.roll(x, -1)))


def _rings_to_shapely(rings):
    """Esri rings: outer rings clockwise (negative signed area), holes counter-clockwise."""
    outers, holes = [], []
    for r in rings:
        (outers if _ring_area(r) < 0 else holes).append(r)
    if not outers:
        outers, holes = holes, []
    polys = []
    for o in outers:
        po = _ShpPolygon(o)
        hs = [h for h in holes if po.contains(_ShpPolygon(h).representative_point())]
        polys.append(_ShpPolygon(o, hs).buffer(0))
    return _ShpMultiPolygon([p for g in polys for p in getattr(g, "geoms", [g])]) if len(polys) > 1 else polys[0]


def split_rings(rings):
    outers, holes = [], []
    for r in rings:
        (outers if _ring_area(r) < 0 else holes).append(r)
    if not outers:
        outers, holes = holes, []
    return outers, holes


def simplify_polyline(pts, tol):
    """Douglas-Peucker (iterative); keeps first/last points."""
    n = len(pts)
    if n < 3 or tol <= 0:
        return pts
    keep = np.zeros(n, dtype=bool)
    keep[0] = keep[-1] = True
    stack = [(0, n - 1)]
    while stack:
        i0, i1 = stack.pop()
        if i1 <= i0 + 1:
            continue
        a, b = pts[i0], pts[i1]
        seg = pts[i0 + 1:i1]
        d = b - a
        L = math.hypot(d[0], d[1])
        if L == 0:
            dist = np.hypot(seg[:, 0] - a[0], seg[:, 1] - a[1])
        else:
            dist = np.abs(d[0] * (seg[:, 1] - a[1]) - d[1] * (seg[:, 0] - a[0])) / L
        k = int(np.argmax(dist))
        if dist[k] > tol:
            m = i0 + 1 + k
            keep[m] = True
            stack.append((i0, m))
            stack.append((m, i1))
    return pts[keep]


def guess_jurisdiction(pdf_path, juris_list):
    """Match '95_Washington.pdf' -> VDOT_CODE '095', fall back to name tokens."""
    stem = os.path.splitext(os.path.basename(pdf_path))[0]
    m = re.match(r"^\s*(\d{1,3})(?!\d)", stem)
    words = [w.lower() for w in re.findall(r"[A-Za-z]+", stem)]
    if m:
        code = m.group(1).zfill(3)
        hits = [j for j in juris_list if str(j.get("VDOT_CODE") or "").zfill(3) == code]
        if len(hits) > 1 and words:
            named = [j for j in hits if (j.get("NAME") or "").lower() in " ".join(words)]
            hits = named or hits
        if hits:
            hits.sort(key=lambda j: 0 if j.get("JURISTYPE") == "CO" else 1)
            return hits[0]
    if words:
        text = " ".join(words)
        isCity = "city" in words
        best = None
        for j in juris_list:
            nm = (j.get("NAME") or "").lower()
            if nm and re.search(r"\b%s\b" % re.escape(nm), text):
                score = len(nm) + (5 if (j.get("JURISTYPE") == "CI") == isCity else 0) + (2 if j.get("JURISTYPE") == "CO" else 0)
                if best is None or score > best[0]:
                    best = (score, j)
        if best:
            return best[1]
    return None


# =============================================================================
# PDF rendering (subprocess with PySide6.QtPdf, fallback PyMuPDF in-process)
# =============================================================================
_RENDER_HELPER = r'''
import sys, json
def main():
    pdf, out, dpi, page = sys.argv[1], sys.argv[2], float(sys.argv[3]), int(sys.argv[4])
    from PySide6.QtGui import QGuiApplication, QImage, QPainter, QColor
    from PySide6.QtCore import QSize
    from PySide6.QtPdf import QPdfDocument
    app = QGuiApplication.instance() or QGuiApplication([sys.argv[0]])
    doc = QPdfDocument(None)
    doc.load(pdf)
    if doc.status() != QPdfDocument.Status.Ready or doc.pageCount() < 1:
        print(json.dumps({"error": "Cannot open PDF (status %s)" % doc.status()})); return
    n = doc.pageCount()
    page = max(0, min(page, n - 1))
    sz = doc.pagePointSize(page)
    w = max(1, int(round(sz.width() * dpi / 72.0)))
    h = max(1, int(round(sz.height() * dpi / 72.0)))
    img = doc.render(page, QSize(w, h))
    canvas = QImage(w, h, QImage.Format.Format_RGB32)
    canvas.fill(QColor(255, 255, 255))
    p = QPainter(canvas); p.drawImage(0, 0, img); p.end()
    if not canvas.save(out, "PNG"):
        print(json.dumps({"error": "Cannot write " + out})); return
    print(json.dumps({"pages": n, "page": page, "w_pt": sz.width(), "h_pt": sz.height(), "width": w, "height": h}))
try:
    main()
except Exception as ex:
    print(json.dumps({"error": repr(ex)}))
'''


def _find_python_exe():
    cands = [os.path.join(sys.prefix, "python.exe"), os.path.join(sys.exec_prefix, "python.exe"), BENTLEY_PYTHON]
    if os.path.basename(sys.executable).lower().startswith("python"):
        cands.insert(0, sys.executable)
    for c in cands:
        if c and os.path.isfile(c):
            return c
    return None


def render_pdf_page(pdf_path, out_png, dpi, page=0, wait=None):
    """Render one PDF page to PNG at dpi. Returns dict(pages, w_pt, h_pt, width, height)."""
    _repair_stdlib()
    try:
        try:
            import pymupdf as fitz  # PyMuPDF renders in-process (no subprocess needed)
        except ImportError:
            import fitz
        doc = fitz.open(pdf_path)
        try:
            page = max(0, min(page, doc.page_count - 1))
            pg = doc[page]
            pix = pg.get_pixmap(matrix=fitz.Matrix(dpi / 72.0, dpi / 72.0), alpha=False)
            pix.save(out_png)
            return {"pages": doc.page_count, "page": page, "w_pt": pg.rect.width, "h_pt": pg.rect.height,
                    "width": pix.width, "height": pix.height}
        finally:
            doc.close()
    except ImportError:
        pass
    exe = _find_python_exe()
    if not exe:
        raise RuntimeError("Python executable for PDF rendering not found.")
    helper = os.path.join(tempfile.gettempdir(), "vdot_georef_render_helper.py")
    with open(helper, "w", encoding="utf-8") as f:
        f.write(_RENDER_HELPER)
    flags = 0x08000000 if os.name == "nt" else 0  # CREATE_NO_WINDOW
    env = dict(os.environ)
    for k in ("QT_PLUGIN_PATH", "QT_QPA_PLATFORM_PLUGIN_PATH", "PYTHONHOME", "PYTHONPATH"):
        env.pop(k, None)
    proc = subprocess.Popen([exe, helper, pdf_path, out_png, "%g" % dpi, str(int(page))],
                            stdout=subprocess.PIPE, stderr=subprocess.PIPE, creationflags=flags, env=env)
    while proc.poll() is None:
        if wait:
            wait()
        else:
            time.sleep(0.05)
    out, err = proc.communicate()
    txt = out.decode("utf-8", "replace").strip().splitlines()
    info = None
    for line in reversed(txt):
        try:
            info = json.loads(line)
            break
        except Exception:
            continue
    if info is None:
        raise RuntimeError("PDF render failed: %s" % err.decode("utf-8", "replace")[-800:])
    if "error" in info:
        raise RuntimeError("PDF render failed: %s" % info["error"])
    return info


def make_masked_image(pdf_path, page, pdf_info, T_for, mask_rings, white_transparent, tmpdir):
    """Render the PDF page and save a PNG whose pixels outside mask_rings (world/UOR rings,
    even-odd so holes cut out) and, optionally, near-white paper are fully transparent.
    T_for(w, h) maps top-left-origin pixels of a w x h image to UOR. Returns (path, (w, h))."""
    longest_in = max(pdf_info["w_pt"], pdf_info["h_pt"]) / 72.0
    dpi = min(MASK_IMAGE_DPI, MASK_IMAGE_MAX_PX / max(longest_in, 1e-6))
    raw = os.path.join(tmpdir, "page_%d.png" % int(time.time() * 1000))
    render_pdf_page(pdf_path, raw, dpi, page)
    src = QtGui.QImage(raw).convertToFormat(QtGui.QImage.Format_RGBA8888)
    if src.isNull():
        raise RuntimeError("Rendered page image could not be loaded.")
    w, h = src.width(), src.height()
    try:
        os.remove(raw)
    except OSError:
        pass
    ptr = src.constBits()
    ptr.setsize(src.byteCount())
    rgba = np.frombuffer(ptr, np.uint8).reshape(h, src.bytesPerLine())[:, :4 * w].reshape(h, w, 4).copy()
    del src
    if mask_rings:
        Minv = np.linalg.inv(T_for(w, h))
        path = QtGui.QPainterPath()
        path.setFillRule(QtCore.Qt.OddEvenFill)
        for r in mask_rings:
            r = np.asarray(r, float)
            px = (Minv @ np.vstack([r.T, np.ones(len(r))]))[:2].T
            path.addPolygon(QtGui.QPolygonF([QtCore.QPointF(float(x), float(y)) for x, y in px]))
            path.closeSubpath()
        m = QtGui.QImage(w, h, QtGui.QImage.Format_Grayscale8)
        m.fill(0)
        p = QtGui.QPainter(m)
        p.setRenderHint(QtGui.QPainter.Antialiasing, True)
        p.fillPath(path, QtGui.QColor(255, 255, 255))
        p.end()
        mp = m.constBits()
        mp.setsize(m.byteCount())
        alpha = np.frombuffer(mp, np.uint8).reshape(h, m.bytesPerLine())[:, :w].copy()
        rgba[..., 3] = np.minimum(rgba[..., 3], alpha)
        del m
        if not alpha.any():
            raise RuntimeError("The county boundary does not overlap the PDF page; check the alignment.")
    if white_transparent:
        rgba[(rgba[..., :3] >= 235).all(axis=2), 3] = 0
    rgba[rgba[..., 3] == 0, :3] = 0
    out = QtGui.QImage(rgba.data, w, h, 4 * w, QtGui.QImage.Format_RGBA8888)
    base = os.path.splitext(os.path.basename(pdf_path))[0]
    suffix = "_county" if mask_rings else "_map"
    cands = [os.path.join(os.path.dirname(pdf_path), base + suffix + ".png"),
             os.path.join(os.environ.get("LOCALAPPDATA") or tempfile.gettempdir(), "VDOT_Georef", base + suffix + ".png")]
    for c in cands:
        try:
            os.makedirs(os.path.dirname(c), exist_ok=True)
            if out.save(c, "PNG"):
                return c, (w, h)
        except OSError:
            continue
    raise RuntimeError("Could not save the masked image next to the PDF or in %s." % os.path.dirname(cands[-1]))


def qimage_to_gray(img):
    g = img.convertToFormat(QtGui.QImage.Format_Grayscale8)
    w, h, bpl = g.width(), g.height(), g.bytesPerLine()
    ptr = g.constBits()
    ptr.setsize(h * bpl)
    return np.frombuffer(ptr, np.uint8).reshape(h, bpl)[:, :w].copy()


# =============================================================================
# Image analysis / alignment math (numpy only)
# =============================================================================
def _dilate(a, conn8=True):
    m = a.copy()
    m[1:] |= a[:-1]
    m[:-1] |= a[1:]
    if conn8:
        n = m.copy()
        n[:, 1:] |= m[:, :-1]
        n[:, :-1] |= m[:, 1:]
        return n
    m[:, 1:] |= a[:, :-1]
    m[:, :-1] |= a[:, 1:]
    return m


def _erode(a, conn8=True):
    return ~_dilate(~a, conn8)


def dilate_n(a, n):
    for i in range(n):
        a = _dilate(a, i % 2 == 0)
    return a


def erode_n(a, n):
    for i in range(n):
        a = _erode(a, i % 2 == 0)
    return a


def distance_map(mask, cap):
    """Truncated (octagonal) distance to nearest True pixel."""
    dist = np.full(mask.shape, float(cap), dtype=np.float32)
    dist[mask] = 0.0
    cur = mask.copy()
    for d in range(1, int(cap)):
        nxt = _dilate(cur, d % 2 == 1)
        new = nxt & ~cur
        if not new.any():
            break
        dist[new] = d
        cur = nxt
    return dist


def long_line_mask(mask, L):
    """Pixels belonging to horizontal/vertical runs of length >= L (frames, tables)."""
    H, W = mask.shape
    out = np.zeros_like(mask)
    for axis in (1, 0):
        m = mask if axis == 1 else mask.T
        h, w = m.shape
        if w <= L:
            continue
        cs = np.zeros((h, w + 1), dtype=np.int32)
        np.cumsum(m, axis=1, out=cs[:, 1:])
        full = (cs[:, L:] - cs[:, :-L]) == L
        d = np.zeros((h, w + 1), dtype=np.int32)
        d[:, :w - L + 1] += full
        d[:, L:] -= full
        cov = np.cumsum(d, axis=1)[:, :w] > 0
        out |= cov if axis == 1 else cov.T
    return out


def _label_largest(cells, weight):
    """Largest 8-connected component of `cells`, scored by sum(weight)."""
    H, W = cells.shape
    lab = np.zeros((H, W), dtype=np.int32)
    flat = cells.ravel()
    labf = lab.ravel()
    wf = weight.ravel()
    best, best_score, cur = 0, -1.0, 0
    idx = np.flatnonzero(flat)
    for start in idx:
        if labf[start]:
            continue
        cur += 1
        labf[start] = cur
        stack = [start]
        score = 0.0
        while stack:
            p = stack.pop()
            score += wf[p]
            r, c = divmod(p, W)
            for dr in (-1, 0, 1):
                rr = r + dr
                if rr < 0 or rr >= H:
                    continue
                base = rr * W
                for dc in (-1, 0, 1):
                    cc = c + dc
                    if 0 <= cc < W:
                        q = base + cc
                        if flat[q] and not labf[q]:
                            labf[q] = cur
                            stack.append(q)
        if score > best_score:
            best, best_score = cur, score
    return lab == best if best else np.zeros_like(cells)


def _fill_holes(m):
    """Fill holes: everything not reachable from the border through background."""
    H, W = m.shape
    bg = ~m
    reach = np.zeros_like(m)
    reach[0, :] = bg[0, :]
    reach[-1, :] = bg[-1, :]
    reach[:, 0] |= bg[:, 0]
    reach[:, -1] |= bg[:, -1]
    for _ in range(4 * (H + W)):
        nxt = _dilate(reach, False) & bg
        if (nxt == reach).all():
            break
        reach = nxt
    return ~reach


def detect_map_region(gray, dpi):
    """Detect the mapped jurisdiction as the largest non-white blob (frames/tables removed)."""
    content = gray < 235
    content &= ~long_line_mask(content, max(20, int(dpi * 1.1)))
    f = max(1, int(round(dpi / 12.0)))
    H2, W2 = gray.shape[0] // f, gray.shape[1] // f
    dens = content[:H2 * f, :W2 * f].reshape(H2, f, W2, f).mean(axis=(1, 3))
    cells = dens > 0.06
    r = 2
    dil = dilate_n(cells, r)
    comp = _label_largest(dil, dens)
    closed = erode_n(comp, r) & comp
    filled = _fill_holes(closed)
    full = np.zeros(gray.shape, dtype=bool)
    full[:H2 * f, :W2 * f] = np.repeat(np.repeat(filled, f, axis=0), f, axis=1)
    return full


def bilinear(img, x, y, fill):
    H, W = img.shape
    x = x - 0.5
    y = y - 0.5
    ok = (x >= 0) & (y >= 0) & (x < W - 1) & (y < H - 1)
    out = np.full(x.shape, fill, dtype=np.float64)
    if ok.any():
        xs, ys = x[ok], y[ok]
        x0 = xs.astype(np.int64)
        y0 = ys.astype(np.int64)
        fx, fy = xs - x0, ys - y0
        v = (img[y0, x0] * (1 - fx) * (1 - fy) + img[y0, x0 + 1] * fx * (1 - fy) +
             img[y0 + 1, x0] * (1 - fx) * fy + img[y0 + 1, x0 + 1] * fx * fy)
        out[ok] = v
    return out


def apply_G(G, pts):
    return pts @ G[:, :2].T + G[:, 2]


def sim_to_G(p):
    tx, ty, ls, th = p
    s = math.exp(ls)
    c, sn = math.cos(th), math.sin(th)
    return np.array([[s * c, s * sn, tx], [s * sn, -s * c, ty]], dtype=float)


def G_to_sim(G):
    a, b = G[0, 0], G[0, 1]
    return np.array([G[0, 2], G[1, 2], math.log(max(1e-12, math.hypot(a, b))), math.atan2(b, a)])


def nelder_mead(f, x0, steps, iters=500, tol=1e-7, tick=None):
    n = len(x0)
    pts = [np.asarray(x0, float)]
    for i in range(n):
        x = np.asarray(x0, float).copy()
        x[i] += steps[i]
        pts.append(x)
    vals = [f(p) for p in pts]
    for it in range(iters):
        order = np.argsort(vals)
        pts = [pts[i] for i in order]
        vals = [vals[i] for i in order]
        if abs(vals[-1] - vals[0]) <= tol * (abs(vals[0]) + 1e-9):
            break
        if tick and it % 25 == 0:
            tick()
        c = np.mean(pts[:-1], axis=0)
        xr = c + (c - pts[-1])
        fr = f(xr)
        if fr < vals[0]:
            xe = c + 2.0 * (c - pts[-1])
            fe = f(xe)
            pts[-1], vals[-1] = (xe, fe) if fe < fr else (xr, fr)
        elif fr < vals[-2]:
            pts[-1], vals[-1] = xr, fr
        else:
            xc = c + 0.5 * (pts[-1] - c) if fr >= vals[-1] else c + 0.5 * (xr - c)
            fc = f(xc)
            if fc < min(fr, vals[-1]):
                pts[-1], vals[-1] = xc, fc
            else:
                for i in range(1, len(pts)):
                    pts[i] = pts[0] + 0.5 * (pts[i] - pts[0])
                    vals[i] = f(pts[i])
    i = int(np.argmin(vals))
    return pts[i], vals[i]


def densify(polys, spacing, max_pts=8000):
    out = []
    for p in polys:
        if len(p) < 2:
            continue
        seg = np.diff(p, axis=0)
        L = np.hypot(seg[:, 0], seg[:, 1])
        for i in range(len(seg)):
            k = max(1, int(L[i] / spacing))
            t = np.arange(k)[:, None] / k
            out.append(p[i] + t * seg[i])
        out.append(p[-1:])
    if not out:
        return np.zeros((0, 2))
    pts = np.vstack(out)
    if len(pts) > max_pts:
        pts = pts[np.linspace(0, len(pts) - 1, max_pts).astype(int)]
    return pts


class Aligner:
    """Solves G (2x3): normalized world -> preview pixel."""

    def __init__(self, gray, dpi):
        self.gray = gray
        self.dpi = dpi
        self._region = None
        self._edge_dist = None
        self._ink_dist = None

    def region(self):
        if self._region is None:
            self._region = detect_map_region(self.gray, self.dpi)
        return self._region

    def edge_dist(self):
        if self._edge_dist is None:
            r = self.region()
            edge = r & ~_erode(r, False)
            self._edge_dist = distance_map(edge, 40)
        return self._edge_dist

    def ink_dist(self):
        if self._ink_dist is None:
            dark = self.gray < 120
            dark &= ~long_line_mask(dark, max(20, int(self.dpi * 1.1)))
            self._ink_dist = distance_map(dark, 10)
        return self._ink_dist

    def initial_fit(self, world_rings):
        r = self.region()
        ys, xs = np.nonzero(r)
        if len(xs) < 50:
            raise RuntimeError("Could not detect the map area in the PDF.")
        px0, px1, py0, py1 = xs.min(), xs.max() + 1, ys.min(), ys.max() + 1
        allw = np.vstack(world_rings)
        wx0, wy0 = allw.min(axis=0)
        wx1, wy1 = allw.max(axis=0)
        s = math.sqrt(((px1 - px0) / (wx1 - wx0)) * ((py1 - py0) / (wy1 - wy0)))
        wc = np.array([(wx0 + wx1) / 2, (wy0 + wy1) / 2])
        pc = np.array([(px0 + px1) / 2, (py0 + py1) / 2])
        G = np.array([[s, 0, 0], [0, -s, 0]], dtype=float)
        G[:, 2] = pc - apply_G(G, wc[None, :])[0]
        return G

    def _objective(self, dmap, pts, cap, weights=None):
        def f(G):
            pp = apply_G(G, pts)
            d = bilinear(dmap, pp[:, 0], pp[:, 1], cap)
            return float(np.average(d, weights=weights))
        return f

    def refine(self, G, world_rings, route_lines=None, affine=False, stage1=True, tick=None):
        s = math.hypot(G[0, 0], G[0, 1])
        if stage1:
            pts = densify(world_rings, 2.0 / s, 5000)
            f = self._objective(self.edge_dist(), pts, 40.0)
            x, _ = nelder_mead(lambda p: f(sim_to_G(p)), G_to_sim(G), [12, 12, 0.03, 0.02], 400, tick=tick)
            x, _ = nelder_mead(lambda p: f(sim_to_G(p)), x, [4, 4, 0.01, 0.006], 300, tick=tick)
            G = sim_to_G(x)
            s = math.hypot(G[0, 0], G[0, 1])
        pts = densify(world_rings, 1.0 / s, 8000)
        w = np.ones(len(pts))
        if route_lines:
            rp = densify(route_lines, 2.0 / s, 8000)
            if len(rp):
                pts = np.vstack([pts, rp])
                w = np.concatenate([w, np.full(len(rp), 0.5 * len(w) / len(rp))])
        f = self._objective(self.ink_dist(), pts, 10.0, w)
        x, _ = nelder_mead(lambda p: f(sim_to_G(p)), G_to_sim(G), [3, 3, 0.006, 0.004], 400, tick=tick)
        x, _ = nelder_mead(lambda p: f(sim_to_G(p)), x, [1, 1, 0.002, 0.0015], 300, tick=tick)
        G = sim_to_G(x)
        if affine:
            s = math.hypot(G[0, 0], G[0, 1])
            st = [0.004 * s] * 4
            g0 = np.array([G[0, 0], G[0, 1], G[1, 0], G[1, 1], G[0, 2], G[1, 2]])
            fa = lambda q: f(np.array([[q[0], q[1], q[4]], [q[2], q[3], q[5]]]))
            q, _ = nelder_mead(fa, g0, st + [1.5, 1.5], 900, tick=tick)
            G = np.array([[q[0], q[1], q[4]], [q[2], q[3], q[5]]])
        return G

    def score(self, G, world_rings):
        s = math.hypot(G[0, 0], G[0, 1])
        pts = densify(world_rings, 1.0 / s, 8000)
        pp = apply_G(G, pts)
        d = bilinear(self.ink_dist(), pp[:, 0], pp[:, 1], 10.0)
        return float(d.mean()), float((d <= 2.0).mean())


def solve_from_pairs(world_pts, pix_pts, affine):
    """Least-squares G from tie-point pairs (normalized world -> pixel)."""
    W = np.asarray(world_pts, float)
    P = np.asarray(pix_pts, float)
    n = len(W)
    if affine and n >= 3:
        A = np.hstack([W, np.ones((n, 1))])
        cx, _, _, _ = np.linalg.lstsq(A, P[:, 0], rcond=None)
        cy, _, _, _ = np.linalg.lstsq(A, P[:, 1], rcond=None)
        G = np.vstack([cx, cy])
    elif n >= 2:
        # G = [[a, b, tx], [b, -a, ty]]  (similarity with Y flip)
        A = np.zeros((2 * n, 4))
        A[0::2, 0], A[0::2, 1], A[0::2, 2] = W[:, 0], W[:, 1], 1
        A[1::2, 0], A[1::2, 1], A[1::2, 3] = -W[:, 1], W[:, 0], 1
        b = P.reshape(-1)
        a, bb, tx, ty = np.linalg.lstsq(A, b, rcond=None)[0]
        G = np.array([[a, bb, tx], [bb, -a, ty]])
    else:
        return None, None
    res = apply_G(G, W) - P
    return G, float(np.sqrt((res ** 2).sum(axis=1).mean()))


# =============================================================================
# Scale bar sizing
# =============================================================================
def scale_bar_spec(target_master, master_unit):
    """Pick a round scale-bar length near target_master (in master units).

    Returns (length in master units, [label at 0, mid, end], unit name)."""
    u = (master_unit or "").strip().lower()
    if u in ("ft", "'", "feet", "foot", "us survey feet", "sft"):
        factor, name = (5280.0, "Miles") if target_master >= 2640 else (1.0, "Feet")
    elif u in ("m", "meters", "metres", "meter"):
        factor, name = (1000.0, "Kilometers") if target_master >= 1000 else (1.0, "Meters")
    else:
        factor, name = 1.0, (master_unit or "Units")
    t = max(1e-9, target_master / factor)
    mag = 10.0 ** np.floor(np.log10(t))
    nice = max(s * mag for s in (1.0, 2.0, 2.5, 5.0) if s * mag <= t * 1.0000001)

    def fmt(v):
        return ("%.3f" % v).rstrip("0").rstrip(".")
    return nice * factor, ["0", fmt(nice / 2.0), fmt(nice)], name


# =============================================================================
# Coordinate conversion (MicroStation GCS; pyproj fallback when run outside)
# =============================================================================
class CoordConverter:
    def __init__(self):
        self.ok = False
        self.label = "No GCS"
        self.epsg = 0
        self.uor_per_master = 1.0
        self.master_label = "units"
        self.datum_sr = 4269
        self._gcs = None
        self._proj = None
        if HAVE_MSTN:
            try:
                model = ISessionMgr.ActiveDgnModelRef
                self._gcs = DgnGCS.FromModel(model, True)
                if self._gcs is not None and self._gcs.IsValid():
                    self.ok = True
                    try:
                        self.epsg = int(self._gcs.GetEPSGCode(False))
                    except Exception:
                        pass
                    nm = ""
                    try:
                        nm = str(self._gcs.GetName())
                    except Exception:
                        pass
                    try:
                        dn = str(self._gcs.GetDatumName())
                        if "WGS84" in dn.replace("-", "").upper():
                            self.datum_sr = 4326
                    except Exception:
                        pass
                    self.label = "%s%s" % (nm or "GCS", (" · EPSG %d" % self.epsg) if self.epsg else "")
                try:
                    mi = model.GetModelInfo()
                    mu = mi.GetMasterUnit()
                    _, per_m = mu.GetConversionFactorFrom(
                        UnitDefinition.GetStandardUnit(StandardUnit.eMetricMeters))
                    self.uor_per_master = float(mi.GetUorPerMeter()) / float(per_m)
                    lab = str(mu.GetLabel())
                    self.master_label = {"'": "ft", '"': "in"}.get(lab, lab)
                except Exception:
                    pass
            except Exception as ex:
                _log_console("GCS error: %s" % ex)
        else:
            try:
                from pyproj import Transformer
                self._proj = Transformer.from_crs(4269, 2284, always_xy=True)
                self.ok = True
                self.epsg = 2284
                self.label = "Preview only · EPSG 2284 (pyproj)"
                self.master_label = "ft"
            except Exception:
                pass

    def to_world(self, ll):
        ll = np.asarray(ll, float)
        if self._proj is not None:
            x, y = self._proj.transform(ll[:, 0], ll[:, 1])
            return np.column_stack([x, y])
        out = np.empty_like(ll)
        gp = GeoPoint()
        pt = DPoint3d()
        for i in range(len(ll)):
            gp.longitude = float(ll[i, 0])
            gp.latitude = float(ll[i, 1])
            gp.elevation = 0.0
            self._gcs.UorsFromLatLong(pt, gp)
            out[i, 0] = pt.x
            out[i, 1] = pt.y
        return out

    def to_lonlat(self, w):
        w = np.asarray(w, float)
        if self._proj is not None:
            from pyproj import Transformer
            t = Transformer.from_crs(2284, 4269, always_xy=True)
            x, y = t.transform(w[:, 0], w[:, 1])
            return np.column_stack([x, y])
        out = np.empty_like(w)
        gp = GeoPoint()
        for i in range(len(w)):
            self._gcs.LatLongFromUors(gp, DPoint3d(float(w[i, 0]), float(w[i, 1]), 0.0))
            out[i, 0], out[i, 1] = gp.longitude, gp.latitude
        return out


# =============================================================================
# MicroStation placement
# =============================================================================
class MstnPlacer:
    def __init__(self):
        self.model = ISessionMgr.ActiveDgnModelRef
        self.dgn = self.model.GetDgnFile()
        self.is3d = bool(self.model.Is3d())

    def level_id(self, name, color=None, weight=None, description=None):
        lc = self.dgn.GetLevelCache()
        lh = lc.GetLevelByName(name)
        if not lh.IsValid():
            lh = lc.CreateLevel(name, LEVEL_NULL_CODE, LEVEL_NULL_ID)
            try:
                eh = EditLevelHandle(lh)
                if color is not None:
                    eh.SetByLevelColor(LevelDefinitionColor(color, self.dgn))
                if weight is not None:
                    eh.SetByLevelWeight(weight)
                if description:
                    eh.SetDescription(description)
            except Exception:
                pass
            lc.Write()
        return lh.GetLevelId()

    def _apply_props(self, eeh, level, color, weight):
        st = ElementPropertiesSetter()
        st.SetLevel(level)
        if color is not None:
            st.SetColor(color)
        if weight is not None:
            st.SetWeight(weight)
        st.Apply(eeh)

    def add_curve(self, pts, closed, level, color, weight):
        if len(pts) < 2:
            return False
        if closed and (pts[0] != pts[-1]).any():
            pts = np.vstack([pts, pts[:1]])
        arr = DPoint3dArray()
        for x, y in pts:
            arr.append(DPoint3d(float(x), float(y), 0.0))
        btype = CurveVector.eBOUNDARY_TYPE_Outer if closed else CurveVector.eBOUNDARY_TYPE_Open
        cv = CurveVector.CreateLinear(arr, btype, True)
        eeh = EditElementHandle()
        if DraftingElementSchema.ToElement(eeh, cv, None, self.is3d, self.model) != BentleyStatus.eSUCCESS:
            return False
        self._apply_props(eeh, level, color, weight)
        return eeh.AddToModel() == BentleyStatus.eSUCCESS

    def add_shape(self, pts, level, color, weight=0, fill=None):
        pts = np.asarray(pts, float)
        if (pts[0] != pts[-1]).any():
            pts = np.vstack([pts, pts[:1]])
        arr = DPoint3dArray()
        for x, y in pts:
            arr.append(DPoint3d(float(x), float(y), 0.0))
        cv = CurveVector.CreateLinear(arr, CurveVector.eBOUNDARY_TYPE_Outer, True)
        eeh = EditElementHandle()
        if DraftingElementSchema.ToElement(eeh, cv, None, self.is3d, self.model) != BentleyStatus.eSUCCESS:
            return False
        self._apply_props(eeh, level, color, weight)
        if fill is not None:
            try:
                eeh.GetHandler().AddSolidFill(eeh, int(fill), True)
            except Exception as ex:
                _log_console("fill failed: %s" % ex)
        return eeh.AddToModel() == BentleyStatus.eSUCCESS

    def add_text(self, text, x, y, height, level, color, just=None):
        model = self.model.GetDgnModel()
        tbp = TextBlockProperties.Create(model)
        pp = ParagraphProperties.Create(model)
        pp.SetJustification(just if just is not None else TextElementJustification.eCenterMiddle)
        rp = RunProperties.Create(DgnFontManager.GetDefaultTrueTypeFont(),
                                  DPoint2d(float(height), float(height)), model)
        tb = TextBlock(tbp, pp, rp, model)
        tb.AppendText(str(text))
        tb.SetUserOrigin(DPoint3d(float(x), float(y), 0.0))
        eeh = EditElementHandle()
        if TextElemHandler.CreateElement(eeh, None, tb) != TextBlockToElementResult.eTEXTBLOCK_TO_ELEMENT_RESULT_Success:
            return False
        self._apply_props(eeh, level, color, None)
        return eeh.AddToModel() == BentleyStatus.eSUCCESS

    @staticmethod
    def define_fence(ring_w):
        """Make the given world ring the active MicroStation fence (view 1)."""
        arr = DPoint3dArray()
        for x, y in np.asarray(ring_w, float)[:-1]:
            arr.append(DPoint3d(float(x), float(y), 0.0))
        vp = IViewManager.GetActiveViewSet().GetViewport(0)
        return FenceManager.GetManager().DefineByPoints(arr, vp) == 0

    def add_scale_and_north(self, rings_w, conv, level, color=0):
        """Scale bar and true-north arrow just below the bottom-right corner of the boundary."""
        allp = np.vstack([np.asarray(r, float) for r in rings_w])
        xmin, ymin = allp.min(axis=0)
        xmax, ymax = allp.max(axis=0)
        width = float(xmax - xmin)
        upm = max(1e-12, float(conv.uor_per_master))
        unit = conv.master_label
        L_uor, labels, unit_txt = scale_bar_spec(width / upm / 4.0, unit)
        L_uor *= upm
        hb = L_uor / 22.0                 # bar height
        th = hb * 1.3                     # label text height
        margin = max(width * 0.03, hb * 3)
        x0 = float(xmax) - L_uor
        y0 = float(ymin) - margin - hb
        n = 0
        seg = L_uor / 4.0
        for i in range(4):
            xa, xb = x0 + i * seg, x0 + (i + 1) * seg
            n += bool(self.add_shape([(xa, y0), (xb, y0), (xb, y0 + hb), (xa, y0 + hb)], level, color, 1,
                                     fill=color if i % 2 == 0 else None))
        ly = y0 + hb + th * 0.9
        for frac, lab in zip((0.0, 0.5, 1.0), labels):
            n += bool(self.add_text(lab, x0 + frac * L_uor, ly, th, level, color))
        n += bool(self.add_text(unit_txt, x0 + L_uor / 2.0, y0 - th * 0.9, th, level, color))

        # North arrow, left of the scale bar, rotated to true north at its location
        A = max(L_uor * 0.22, hb * 7)
        cx, cy = x0 - A * 0.9, y0 + hb / 2.0 + A * 0.25
        dx, dy = 0.0, 1.0
        try:
            ll = conv.to_lonlat(np.array([[cx, cy]]))[0]
            w2 = conv.to_world(np.array([[ll[0], ll[1] + 0.01]]))[0]
            v = np.array([w2[0] - cx, w2[1] - cy])
            if np.hypot(*v) > 0:
                dx, dy = v / np.hypot(*v)
        except Exception as ex:
            _log_console("true north: %s" % ex)

        def P(u, v):
            return (cx + u * dy + v * dx, cy - u * dx + v * dy)
        tip, notch = P(0, A / 2.0), P(0, -A / 5.0)
        left, right = P(-A * 0.28, -A / 2.0), P(A * 0.28, -A / 2.0)
        n += bool(self.add_shape([tip, notch, left], level, color, 1, fill=color))
        n += bool(self.add_shape([tip, right, notch], level, color, 1))
        nx, ny = P(0, A / 2.0 + th * 1.4)
        n += bool(self.add_text("N", nx, ny, th * 1.6, level, color))
        return n, "%s %s" % (labels[-1], unit_txt)

    @staticmethod
    def _ms_transform(T3, height_px):
        # T3 maps top-left-origin pixels (row down); MicroStation raster space has its
        # origin at the bottom-left pixel with y up, so flip rows: row = H - y.
        M = np.asarray(T3, dtype=float) @ np.array([[1.0, 0.0, 0.0],
                                                    [0.0, -1.0, float(height_px)],
                                                    [0.0, 0.0, 1.0]])
        return Transform.FromRowValues(M[0, 0], M[0, 1], 0.0, M[0, 2],
                                       M[1, 0], M[1, 1], 0.0, M[1, 2],
                                       0.0, 0.0, 1.0, 0.0)

    def _create_raster(self, path, T_for, wh, level, transparency_pct, page, visible):
        w, hgt = int(wh[0]), int(wh[1])
        T3 = T_for(w, hgt)
        ext = DPoint2d(w * math.hypot(T3[0, 0], T3[1, 0]), hgt * math.hypot(T3[0, 1], T3[1, 1]))
        moniker = DgnDocumentMoniker.CreateFromFileName(path, None)
        eeh = EditElementHandle()
        st = IRasterAttachmentEdit.CreateRasterAttachment(eeh, None, moniker, self._ms_transform(T3, hgt),
                                                          ext, self.model)
        if st != BentleyStatus.eSUCCESS:
            raise RuntimeError("CreateRasterAttachment failed for %s: %s" % (os.path.basename(path), st))
        h = eeh.GetRasterFrameHandler()
        opts = [lambda: h.SetGeoreferencePriority(eeh, GeoreferencePriority.eGeoreferencePriority_Attachment),
                lambda: h.SetGCSInheritedFromModelState(eeh, False),
                lambda: h.SetDisplayPriorityPlane(eeh, RasterDisplayPriorityPlane.eDisplayPriority_BackPlane),
                lambda: h.SetTransparencyState(eeh, True)]
        if page is not None:
            opts.append(lambda: h.SetPageNumber(eeh, int(page)))
        if transparency_pct > 0:
            opts.append(lambda: h.SetImageTransparencyLevel(eeh, int(round(255 * transparency_pct / 100.0))))
        if not visible:
            opts += [lambda v=v: h.SetViewState(eeh, v, False) for v in range(8)]
        for fn in opts:
            try:
                fn()
            except Exception as ex:
                _log_console("Raster attach option failed: %s" % ex)
        try:
            self._apply_props(eeh, level, None, None)
        except Exception:
            pass
        if eeh.AddToModel() != BentleyStatus.eSUCCESS:
            raise RuntimeError("Adding raster attachment %s failed." % os.path.basename(path))
        return eeh

    def _bitmap_size(self, ref):
        rast = DgnRasterCollection.GetRasters(self.model).Find(ref)
        src = rast.GetRasterSource() if rast is not None else None
        if src is not None and src.IsValidSource():
            w, hgt = int(src.GetBitmapWidth()), int(src.GetBitmapHeight())
            if w > 0 and hgt > 0:
                return w, hgt
        return None

    def attach_raster(self, path, T_for, wh, level, transparency_pct=0, page=None, visible=True):
        """Attach a georeferenced image or PDF page. T_for(w, h) maps top-left-origin pixels of a
        w x h bitmap to UOR. MicroStation chooses a PDF's bitmap size itself, so if it differs
        from wh the attachment is re-created at the real size. visible=False turns it off in
        all 8 views. Returns the bitmap size used."""
        eeh = self._create_raster(path, T_for, wh, level, transparency_pct, page, visible)
        size = (int(wh[0]), int(wh[1]))
        try:
            real = self._bitmap_size(eeh.GetElementRef())
        except Exception as ex:
            _log_console("Bitmap size query failed: %s" % ex)
            real = None
        if real is not None and real != size:
            if EditElementHandle(eeh.GetElementRef(), self.model).DeleteFromModel() == BentleyStatus.eSUCCESS:
                self._create_raster(path, T_for, real, level, transparency_pct, page, visible)
                size = real
        return size

    @staticmethod
    def fit_view():
        try:
            PyCadInputQueue.SendKeyin("FIT VIEW EXTENDED")
        except Exception:
            pass


# =============================================================================
# UI
# =============================================================================
STYLE_TEMPLATE = """
* { font-family: 'Segoe UI'; font-size: 10pt; color: %(text)s; }
QWidget#root { background: %(bg)s; }
QFrame#header { background: qlineargradient(x1:0,y1:0,x2:1,y2:0, stop:0 %(hdr1)s, stop:1 %(hdr2)s);
                border-bottom: 1px solid %(border)s; }
QLabel#title { font-size: 15pt; font-weight: 600; color: %(title)s; }
QLabel#subtitle { color: %(subtitle)s; }
QLabel#chip { background: %(chip_bg)s; color: %(chip_fg)s; border: 1px solid %(chip_bd)s; border-radius: 10px; padding: 3px 10px; }
QLabel#chipWarn { background: %(warn_bg)s; color: %(warn_fg)s; border: 1px solid %(warn_bd)s; border-radius: 10px; padding: 3px 10px; }
QScrollArea, QScrollArea > QWidget > QWidget { background: %(bg)s; border: none; }
QFrame#card { background: %(card)s; border: 1px solid %(border)s; border-radius: 12px; }
QLabel#cardTitle { font-size: 11pt; font-weight: 600; color: %(card_title)s; }
QLabel#badge { background: %(accent)s; color: %(on_accent)s; border-radius: 11px; font-weight: 700;
               min-width: 22px; max-width: 22px; min-height: 22px; max-height: 22px; qproperty-alignment: AlignCenter; }
QLabel#hint { color: %(hint)s; font-size: 9pt; }
QLabel#metric { color: %(metric)s; font-weight: 600; }
QLineEdit, QSpinBox, QDoubleSpinBox, QComboBox, QListWidget, QPlainTextEdit {
    background: %(input_bg)s; border: 1px solid %(input_bd)s; border-radius: 8px; padding: 5px 8px;
    selection-background-color: %(accent)s; selection-color: %(on_accent)s; }
QLineEdit:focus, QSpinBox:focus, QListWidget:focus { border: 1px solid %(accent)s; }
QListWidget::item { padding: 4px 6px; border-radius: 6px; }
QListWidget::item:selected { background: %(sel_bg)s; color: %(sel_fg)s; }
QListWidget::item:hover { background: %(hover)s; }
QPushButton { background: %(btn)s; border: 1px solid %(btn_bd)s; border-radius: 8px; padding: 7px 12px; }
QPushButton:hover { background: %(btn_hover)s; border-color: %(btn_hover_bd)s; }
QPushButton:pressed { background: %(btn_pressed)s; }
QPushButton:disabled { color: %(dis_fg)s; background: %(dis_bg)s; border-color: %(dis_bd)s; }
QPushButton:checked { background: %(sel_bg)s; border-color: %(accent)s; color: %(sel_fg)s; }
QPushButton#primary { background: qlineargradient(x1:0,y1:0,x2:1,y2:0, stop:0 %(accent)s, stop:1 %(accent_end)s);
                      color: %(on_accent)s; font-weight: 700; font-size: 11pt; border: none; padding: 11px; }
QPushButton#primary:hover { background: %(primary_hover)s; }
QPushButton#primary:disabled { background: %(primary_dis_bg)s; color: %(primary_dis_fg)s; }
QPushButton#accent { border-color: %(accent2_bd)s; color: %(accent2_fg)s; }
QPushButton#accent:hover { background: %(accent2_hover)s; }
QPushButton#tool { padding: 5px 10px; }
QCheckBox, QRadioButton { spacing: 8px; }
QCheckBox::indicator, QRadioButton::indicator { width: 16px; height: 16px; }
QCheckBox::indicator:unchecked { border: 1px solid %(chk_bd)s; border-radius: 4px; background: %(input_bg)s; }
QCheckBox::indicator:checked { border: 1px solid %(accent)s; border-radius: 4px; background: %(accent)s; }
QSlider::groove:horizontal { height: 4px; background: %(groove)s; border-radius: 2px; }
QSlider::handle:horizontal { background: %(accent)s; width: 14px; margin: -6px 0; border-radius: 7px; }
QSlider::sub-page:horizontal { background: %(accent)s; border-radius: 2px; }
QProgressBar { background: %(input_bg)s; border: 1px solid %(border)s; border-radius: 6px; height: 10px; text-align: center; color: transparent; }
QProgressBar::chunk { background: qlineargradient(x1:0,y1:0,x2:1,y2:0, stop:0 %(accent)s, stop:1 %(chunk2)s); border-radius: 6px; }
QFrame#statusbar { background: %(statusbar)s; border-top: 1px solid %(border)s; }
QFrame#viewbar { background: %(viewbar)s; border-bottom: 1px solid %(border)s; }
QGraphicsView { background: %(canvas)s; border: none; }
QSplitter::handle { background: %(border)s; }
QSplitter::handle:horizontal { width: 7px; }
QSplitter::handle:hover, QSplitter::handle:pressed { background: %(accent)s; }
QToolTip { background: %(tip_bg)s; color: %(tip_fg)s; border: 1px solid %(tip_bd)s; padding: 6px; }
QScrollBar:vertical { background: transparent; width: 10px; }
QScrollBar::handle:vertical { background: %(scroll)s; border-radius: 5px; min-height: 30px; }
QScrollBar::add-line, QScrollBar::sub-line { height: 0; }
"""

THEMES = {
    "dark": dict(
        text="#e5e7eb", bg="#15171c", hdr1="#1f2937", hdr2="#111827", border="#2a2f3a", title="#f9fafb",
        subtitle="#9ca3af", chip_bg="#1e3a2f", chip_fg="#6ee7b7", chip_bd="#2f6b54", warn_bg="#3b2a12",
        warn_fg="#fbbf24", warn_bd="#6b4e1f", card="#1d2027", card_title="#f3f4f6", accent="#f28c28",
        accent_end="#e8641b", on_accent="#111827", hint="#8b93a3", metric="#93c5fd", input_bg="#12141a",
        input_bd="#323846", sel_bg="#3a2a1a", sel_fg="#fdba74", hover="#222632", btn="#262b36",
        btn_bd="#343b4a", btn_hover="#2f3542", btn_hover_bd="#4b5563", btn_pressed="#1f2430",
        dis_fg="#5b6170", dis_bg="#1c1f26", dis_bd="#262a33", primary_hover="#f59e0b",
        primary_dis_bg="#3a3f4b", primary_dis_fg="#6b7280", accent2_bd="#0e7490", accent2_fg="#67e8f9",
        accent2_hover="#10343d", chk_bd="#4b5563", groove="#323846", chunk2="#22d3ee",
        statusbar="#111318", viewbar="#1a1d24", canvas="#0f1115", tip_bg="#111827", tip_fg="#e5e7eb",
        tip_bd="#374151", scroll="#2f3542",
        # inline (non-stylesheet) colours
        ok="#6ee7b7", warn="#fbbf24", err="#f87171", info="#cbd5e1", stamp="#6b7280",
        drop_line="#3b4252", drop_text="#9ca3af"),
    "light": dict(
        text="#1f2937", bg="#f3f4f6", hdr1="#ffffff", hdr2="#eef2f7", border="#d6dae1", title="#111827",
        subtitle="#6b7280", chip_bg="#dcfce7", chip_fg="#166534", chip_bd="#86efac", warn_bg="#fef3c7",
        warn_fg="#92400e", warn_bd="#fcd34d", card="#ffffff", card_title="#111827", accent="#f28c28",
        accent_end="#e8641b", on_accent="#111827", hint="#6b7280", metric="#1d4ed8", input_bg="#ffffff",
        input_bd="#cbd2dc", sel_bg="#fde7cf", sel_fg="#9a3412", hover="#f1f3f7", btn="#f8fafc",
        btn_bd="#cbd2dc", btn_hover="#eef1f5", btn_hover_bd="#9ca3af", btn_pressed="#e5e7eb",
        dis_fg="#a0a7b4", dis_bg="#f3f4f6", dis_bd="#e2e5ea", primary_hover="#f59e0b",
        primary_dis_bg="#e5e7eb", primary_dis_fg="#9ca3af", accent2_bd="#0e7490", accent2_fg="#0e7490",
        accent2_hover="#e0f7fb", chk_bd="#9ca3af", groove="#d1d5db", chunk2="#06b6d4",
        statusbar="#ffffff", viewbar="#f8fafc", canvas="#e5e7eb", tip_bg="#ffffff", tip_fg="#1f2937",
        tip_bd="#d1d5db", scroll="#c4cad4",
        ok="#15803d", warn="#b45309", err="#dc2626", info="#374151", stamp="#9ca3af",
        drop_line="#c4cad4", drop_text="#6b7280"),
}
_THEME = ["dark"]


def pal():
    return THEMES[_THEME[0]]


def style_for(theme):
    return STYLE_TEMPLATE % THEMES[theme]


class _GripHandle(QtWidgets.QSplitterHandle):
    """Splitter handle with a visible grip; double-click resets the sidebar width."""

    def __init__(self, orientation, parent):
        super().__init__(orientation, parent)
        self.setCursor(QtCore.Qt.SplitHCursor)
        self.setToolTip("Drag to resize the panel (double-click to reset)")

    def paintEvent(self, ev):
        super().paintEvent(ev)
        p = QtGui.QPainter(self)
        p.setRenderHint(QtGui.QPainter.Antialiasing)
        p.setPen(QtCore.Qt.NoPen)
        p.setBrush(QtGui.QColor(140, 140, 140, 220))
        cx, cy = self.width() / 2.0, self.height() / 2.0
        for dy in (-10, -5, 0, 5, 10):
            p.drawEllipse(QtCore.QPointF(cx, cy + dy), 1.5, 1.5)
        p.end()

    def mouseDoubleClickEvent(self, ev):
        sp = self.splitter()
        tot = sum(sp.sizes())
        sp.setSizes([420, max(400, tot - 420)])
        sp.splitterMoved.emit(420, 1)


class GripSplitter(QtWidgets.QSplitter):
    def createHandle(self):
        return _GripHandle(self.orientation(), self)


class Card(QtWidgets.QFrame):
    def __init__(self, step, title, parent=None):
        super().__init__(parent)
        self.setObjectName("card")
        lay = QtWidgets.QVBoxLayout(self)
        lay.setContentsMargins(14, 12, 14, 14)
        lay.setSpacing(8)
        hdr = QtWidgets.QHBoxLayout()
        b = QtWidgets.QLabel(str(step))
        b.setObjectName("badge")
        t = QtWidgets.QLabel(title)
        t.setObjectName("cardTitle")
        hdr.addWidget(b)
        hdr.addWidget(t)
        hdr.addStretch(1)
        self.status = QtWidgets.QLabel("")
        self.status.setObjectName("hint")
        hdr.addWidget(self.status)
        lay.addLayout(hdr)
        self.body = lay

    def set_done(self, text, ok=True):
        self._ok = ok
        self.status.setText(("✔ " if ok else "• ") + text)
        self.restyle()

    def restyle(self):
        ok = getattr(self, "_ok", None)
        if ok is not None:
            self.status.setStyleSheet("color: %s;" % (pal()["ok"] if ok else pal()["warn"]))


class MapView(QtWidgets.QGraphicsView):
    MODE_PAN, MODE_ADJUST, MODE_TIE, MODE_FENCE = 0, 1, 2, 3
    overlayChanged = QtCore.pyqtSignal()
    tiePicked = QtCore.pyqtSignal(object, object)  # world_norm (np 2), pixel (np 2)
    fileDropped = QtCore.pyqtSignal(str)
    fenceDrawn = QtCore.pyqtSignal(object)  # (x0, y0, x1, y1) preview pixels

    def __init__(self, owner):
        super().__init__()
        self.owner = owner
        self.setScene(QtWidgets.QGraphicsScene(self))
        self.setRenderHints(QtGui.QPainter.Antialiasing | QtGui.QPainter.SmoothPixmapTransform)
        self.setTransformationAnchor(QtWidgets.QGraphicsView.AnchorUnderMouse)
        self.setViewportUpdateMode(QtWidgets.QGraphicsView.FullViewportUpdate)
        self.setAcceptDrops(True)
        self.mode = self.MODE_PAN
        self.pixmap_item = None
        self.overlay = QtWidgets.QGraphicsItemGroup()
        self.overlay.setZValue(10)
        self.scene().addItem(self.overlay)
        self.pin_items = []
        self._drag = None
        self._pending_world = None
        self._pending_item = None
        self.fence_item = None
        self._fence_start = None
        self.set_mode(self.MODE_PAN)

    # ---- block fence (preview pixel space)
    def set_fence_rect(self, r):
        if r is None:
            if self.fence_item is not None:
                self.scene().removeItem(self.fence_item)
                self.fence_item = None
            return
        x0, y0, x1, y1 = r
        rect = QtCore.QRectF(min(x0, x1), min(y0, y1), abs(x1 - x0), abs(y1 - y0))
        if self.fence_item is None:
            self.fence_item = QtWidgets.QGraphicsRectItem()
            pen = QtGui.QPen(QtGui.QColor("#f59e0b"), 2.5, QtCore.Qt.DashLine)
            pen.setCosmetic(True)
            self.fence_item.setPen(pen)
            self.fence_item.setBrush(QtGui.QColor(245, 158, 11, 40))
            self.fence_item.setZValue(30)
            self.scene().addItem(self.fence_item)
        self.fence_item.setRect(rect)

    def _clamp_to_image(self, sp):
        if self.pixmap_item is None:
            return sp.x(), sp.y()
        br = self.pixmap_item.boundingRect()
        return (min(max(sp.x(), br.left()), br.right()), min(max(sp.y(), br.top()), br.bottom()))

    # ---- content
    def set_image(self, qimg):
        if self.pixmap_item is not None:
            self.scene().removeItem(self.pixmap_item)
        self.pixmap_item = self.scene().addPixmap(QtGui.QPixmap.fromImage(qimg))
        self.pixmap_item.setZValue(0)
        self.pixmap_item.setTransformationMode(QtCore.Qt.SmoothTransformation)
        self.scene().setSceneRect(QtCore.QRectF(-200, -200, qimg.width() + 400, qimg.height() + 400))
        self.fit()

    def set_pdf_opacity(self, v):
        if self.pixmap_item is not None:
            self.pixmap_item.setOpacity(v)

    def clear_overlay(self):
        for it in list(self.overlay.childItems()):
            self.overlay.removeFromGroup(it)
            self.scene().removeItem(it)

    def add_overlay_paths(self, polys, color, width, closed, z=0, dash=False):
        path = QtGui.QPainterPath()
        for p in polys:
            if len(p) < 2:
                continue
            path.moveTo(float(p[0, 0]), float(p[0, 1]))
            for x, y in p[1:]:
                path.lineTo(float(x), float(y))
            if closed:
                path.closeSubpath()
        pen = QtGui.QPen(QtGui.QColor(color), width)
        pen.setCosmetic(True)
        pen.setJoinStyle(QtCore.Qt.RoundJoin)
        if dash:
            pen.setStyle(QtCore.Qt.DashLine)
        it = QtWidgets.QGraphicsPathItem(path)
        it.setPen(pen)
        it.setZValue(z)
        self.overlay.addToGroup(it)
        return it

    def set_overlay_G(self, G):
        if G is None:
            self.overlay.setVisible(False)
            return
        self.overlay.setTransform(QtGui.QTransform(G[0, 0], G[1, 0], G[0, 1], G[1, 1], G[0, 2], G[1, 2]))

    def fit(self):
        if self.pixmap_item is not None:
            self.fitInView(self.pixmap_item, QtCore.Qt.KeepAspectRatio)

    def set_mode(self, mode):
        self.mode = mode
        self.setDragMode(QtWidgets.QGraphicsView.ScrollHandDrag if mode == self.MODE_PAN
                         else QtWidgets.QGraphicsView.NoDrag)
        self.viewport().setCursor(QtCore.Qt.OpenHandCursor if mode == self.MODE_PAN else
                                  QtCore.Qt.SizeAllCursor if mode == self.MODE_ADJUST else QtCore.Qt.CrossCursor)
        self.cancel_pending()

    def cancel_pending(self):
        self._pending_world = None
        if self._pending_item is not None:
            self.scene().removeItem(self._pending_item)
            self._pending_item = None

    def draw_pins(self, pairs, G):
        for it in self.pin_items:
            self.scene().removeItem(it)
        self.pin_items = []
        pen_l = QtGui.QPen(QtGui.QColor("#f472b6"), 1.5)
        pen_l.setCosmetic(True)
        for i, (w, p) in enumerate(pairs):
            src = apply_G(G, np.asarray([w]))[0] if G is not None else p
            ln = self.scene().addLine(src[0], src[1], p[0], p[1], pen_l)
            ln.setZValue(20)
            self.pin_items.append(ln)
            for (x, y), col in ((src, "#22d3ee"), (p, "#f472b6")):
                e = self._marker(x, y, col)
                self.pin_items.append(e)
            t = self.scene().addSimpleText(str(i + 1))
            t.setBrush(QtGui.QColor("#fde68a"))
            t.setFlag(QtWidgets.QGraphicsItem.ItemIgnoresTransformations, True)
            t.setPos(p[0] + 2, p[1] + 2)
            t.setZValue(22)
            self.pin_items.append(t)

    def _marker(self, x, y, col):
        e = QtWidgets.QGraphicsEllipseItem(-5, -5, 10, 10)
        pen = QtGui.QPen(QtGui.QColor("#111827"), 1.5)
        pen.setCosmetic(True)
        e.setPen(pen)
        e.setBrush(QtGui.QColor(col))
        e.setFlag(QtWidgets.QGraphicsItem.ItemIgnoresTransformations, True)
        e.setPos(x, y)
        e.setZValue(21)
        self.scene().addItem(e)
        return e

    # ---- events
    def wheelEvent(self, ev):
        mods = ev.modifiers()
        sp = self.mapToScene(ev.pos())
        delta = ev.angleDelta().y() / 120.0
        G = self.owner.G
        if G is not None and self.mode == self.MODE_ADJUST and mods & (QtCore.Qt.ControlModifier | QtCore.Qt.ShiftModifier):
            fine = 0.25 if mods & QtCore.Qt.AltModifier else 1.0
            if mods & QtCore.Qt.ControlModifier:
                f = 1.0 + 0.01 * delta * fine
                M = np.array([[f, 0, sp.x() * (1 - f)], [0, f, sp.y() * (1 - f)]])
            else:
                a = math.radians(0.25 * delta * fine)
                c, s = math.cos(a), math.sin(a)
                M = np.array([[c, -s, sp.x() - c * sp.x() + s * sp.y()], [s, c, sp.y() - s * sp.x() - c * sp.y()]])
            self.owner.compose_G(M)
            ev.accept()
            return
        f = 1.0015 ** ev.angleDelta().y()
        self.scale(f, f)

    def mousePressEvent(self, ev):
        if self.mode == self.MODE_FENCE and ev.button() == QtCore.Qt.LeftButton and self.pixmap_item is not None:
            self._fence_start = self._clamp_to_image(self.mapToScene(ev.pos()))
            self._fence_prev = self.fence_item.rect() if self.fence_item is not None else None
            self.set_fence_rect(self._fence_start + self._fence_start)
            ev.accept()
            return
        if ev.button() == QtCore.Qt.LeftButton and self.mode == self.MODE_ADJUST and self.owner.G is not None:
            self._drag = self.mapToScene(ev.pos())
            ev.accept()
            return
        if self.mode == self.MODE_TIE and self.owner.G is not None:
            if ev.button() == QtCore.Qt.RightButton:
                self.cancel_pending()
                return
            if ev.button() == QtCore.Qt.LeftButton:
                sp = self.mapToScene(ev.pos())
                p = np.array([sp.x(), sp.y()])
                if self._pending_world is None:
                    w = self.owner.nearest_overlay_vertex(p, 25.0 / max(1e-9, self.transform().m11()))
                    if w is None:
                        self.owner.toast("Click closer to the cyan overlay to pick a source point.")
                        return
                    self._pending_world = w
                    src = apply_G(self.owner.G, w[None, :])[0]
                    self._pending_item = self._marker(src[0], src[1], "#22d3ee")
                else:
                    w = self._pending_world
                    self.cancel_pending()
                    self.tiePicked.emit(w, p)
                return
        super().mousePressEvent(ev)

    def mouseMoveEvent(self, ev):
        if self._fence_start is not None:
            self.set_fence_rect(self._fence_start + self._clamp_to_image(self.mapToScene(ev.pos())))
            return
        if self._drag is not None:
            sp = self.mapToScene(ev.pos())
            d = sp - self._drag
            self._drag = sp
            self.owner.compose_G(np.array([[1, 0, d.x()], [0, 1, d.y()]]))
            return
        super().mouseMoveEvent(ev)

    def mouseReleaseEvent(self, ev):
        if self._fence_start is not None:
            x0, y0 = self._fence_start
            x1, y1 = self._clamp_to_image(self.mapToScene(ev.pos()))
            self._fence_start = None
            if abs(x1 - x0) * self.transform().m11() < 6 or abs(y1 - y0) * self.transform().m22() < 6:
                prev = self._fence_prev
                self.set_fence_rect(None if prev is None else
                                    (prev.left(), prev.top(), prev.right(), prev.bottom()))
                return
            self.set_fence_rect((x0, y0, x1, y1))
            self.fenceDrawn.emit((min(x0, x1), min(y0, y1), max(x0, x1), max(y0, y1)))
            return
        if self._drag is not None:
            self._drag = None
            self.owner.after_manual_change()
            return
        super().mouseReleaseEvent(ev)

    def keyPressEvent(self, ev):
        G = self.owner.G
        if G is not None and self.mode == self.MODE_ADJUST and ev.key() in (
                QtCore.Qt.Key_Left, QtCore.Qt.Key_Right, QtCore.Qt.Key_Up, QtCore.Qt.Key_Down):
            step = 0.25 if ev.modifiers() & QtCore.Qt.ShiftModifier else 1.0
            dx = {QtCore.Qt.Key_Left: -step, QtCore.Qt.Key_Right: step}.get(ev.key(), 0)
            dy = {QtCore.Qt.Key_Up: -step, QtCore.Qt.Key_Down: step}.get(ev.key(), 0)
            self.owner.compose_G(np.array([[1, 0, dx], [0, 1, dy]]))
            return
        if ev.key() == QtCore.Qt.Key_Escape:
            self.cancel_pending()
        super().keyPressEvent(ev)

    def dragEnterEvent(self, ev):
        if ev.mimeData().hasUrls():
            ev.acceptProposedAction()

    def dragMoveEvent(self, ev):
        if ev.mimeData().hasUrls():
            ev.acceptProposedAction()

    def dropEvent(self, ev):
        for u in ev.mimeData().urls():
            p = u.toLocalFile()
            if p.lower().endswith(".pdf"):
                self.fileDropped.emit(p)
                break

    def drawForeground(self, painter, rect):
        if self.pixmap_item is None:
            painter.save()
            painter.resetTransform()
            vr = self.viewport().rect()
            painter.setPen(QtGui.QPen(QtGui.QColor(pal()["drop_line"]), 2, QtCore.Qt.DashLine))
            painter.drawRoundedRect(vr.adjusted(30, 30, -30, -30), 16, 16)
            painter.setPen(QtGui.QColor(pal()["drop_text"]))
            f = painter.font()
            f.setPointSize(13)
            painter.setFont(f)
            painter.drawText(vr, QtCore.Qt.AlignCenter,
                             "Drop a VDOT map PDF here\nor use  Browse…  in step 1")
            painter.restore()


class GeorefWindow(QtWidgets.QWidget):
    def __init__(self):
        super().__init__(None, QtCore.Qt.Window)
        self.setObjectName("root")
        self.setWindowTitle("%s %s  —  %s" % (APP_TITLE, APP_VERSION, APP_CREDIT))
        icon_pm = vdot_logo_pixmap()
        if not icon_pm.isNull():
            side = max(icon_pm.width(), icon_pm.height())
            sq = QtGui.QPixmap(side, side)
            sq.fill(QtCore.Qt.transparent)
            qp = QtGui.QPainter(sq)
            qp.drawPixmap(0, (side - icon_pm.height()) // 2, icon_pm)
            qp.end()
            self.setWindowIcon(QtGui.QIcon(sq))
        self.resize(1480, 920)
        self._settings = QtCore.QSettings("VDOT", "MapGeoreferencer")
        theme = str(self._settings.value("theme", "dark"))
        _THEME[0] = theme if theme in THEMES else "dark"
        self.setStyleSheet(style_for(_THEME[0]))
        self._pumping = False
        self._busy_depth = 0
        self.client = VdotRestClient()
        self.conv = CoordConverter()
        self.juris = []
        self.cur_juris = None
        self.pdf_path = None
        self.pdf_info = None
        self.preview_png = None
        self.gray = None
        self.aligner = None
        self.rings_ll = []
        self.rings_w = []
        self.fence_px = None
        self.rings_n = []
        self.routes = []        # (cat, name, ll, world, norm)
        self.norm_c = np.zeros(2)
        self.norm_S = 1.0
        self.G = None
        self.G_auto = None
        self.pairs = []
        self.tmpdir = tempfile.mkdtemp(prefix="vdot_georef_")
        self._build()
        QtCore.QTimer.singleShot(250, self.load_jurisdictions)

    # ------------------------------------------------------------------ build
    def _build(self):
        root = QtWidgets.QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(0)

        hdr = QtWidgets.QFrame()
        hdr.setObjectName("header")
        hl = QtWidgets.QHBoxLayout(hdr)
        hl.setContentsMargins(20, 12, 20, 12)
        logo = QtWidgets.QLabel()
        logo.setObjectName("logo")
        logo.setPixmap(vdot_logo_pixmap(40))
        logo.setStyleSheet("QLabel#logo { background: #ffffff; border-radius: 8px; padding: 6px 10px; }")
        logo.setToolTip("Virginia Department of Transportation")
        logo.setAlignment(QtCore.Qt.AlignCenter)
        hl.addWidget(logo, 0, QtCore.Qt.AlignVCenter)
        hl.addSpacing(14)
        tcol = QtWidgets.QVBoxLayout()
        t = QtWidgets.QLabel(APP_TITLE)
        t.setObjectName("title")
        st = QtWidgets.QLabel("PDF road map  →  VDOT LRS boundary  →  georeferenced, clipped raster in MicroStation")
        st.setObjectName("subtitle")
        cr = QtWidgets.QLabel(APP_CREDIT)
        cr.setObjectName("hint")
        tcol.addWidget(t)
        tcol.addWidget(st)
        tcol.addWidget(cr)
        hl.addLayout(tcol)
        hl.addStretch(1)
        self.bt_theme = QtWidgets.QPushButton()
        self.bt_theme.setObjectName("tool")
        self.bt_theme.setCursor(QtCore.Qt.PointingHandCursor)
        self.bt_theme.clicked.connect(lambda: self.apply_theme("light" if _THEME[0] == "dark" else "dark"))
        self._update_theme_button()
        hl.addWidget(self.bt_theme)
        hl.addSpacing(8)
        self.gcs_chip = QtWidgets.QLabel(("🌐 " + self.conv.label) if self.conv.ok else "⚠ Active model has no GCS")
        self.gcs_chip.setObjectName("chip" if self.conv.ok else "chipWarn")
        self.gcs_chip.setToolTip("Coordinates from the VDOT services are reprojected with the active model's GCS.")
        hl.addWidget(self.gcs_chip)
        root.addWidget(hdr)

        split = GripSplitter(QtCore.Qt.Horizontal)
        split.setChildrenCollapsible(False)
        split.setHandleWidth(7)
        root.addWidget(split, 1)
        self.split = split

        # ---- sidebar
        side = QtWidgets.QScrollArea()
        side.setWidgetResizable(True)
        side.setHorizontalScrollBarPolicy(QtCore.Qt.ScrollBarAlwaysOff)
        side.setMinimumWidth(300)
        sw = QtWidgets.QWidget()
        sl = QtWidgets.QVBoxLayout(sw)
        sl.setContentsMargins(14, 14, 10, 14)
        sl.setSpacing(12)
        side.setWidget(sw)
        split.addWidget(side)
        self.sidebar = sw

        # 1 PDF
        c1 = Card(1, "Map PDF")
        self.card_pdf = c1
        row = QtWidgets.QHBoxLayout()
        self.ed_pdf = QtWidgets.QLineEdit()
        self.ed_pdf.setPlaceholderText("Select a VDOT county road map PDF…")
        self.ed_pdf.setReadOnly(True)
        b = QtWidgets.QPushButton("Browse…")
        b.clicked.connect(self.browse_pdf)
        row.addWidget(self.ed_pdf, 1)
        row.addWidget(b)
        c1.body.addLayout(row)
        row = QtWidgets.QHBoxLayout()
        row.addWidget(QtWidgets.QLabel("Page"))
        self.sp_page = QtWidgets.QSpinBox()
        self.sp_page.setRange(1, 1)
        self.sp_page.valueChanged.connect(lambda v: self.pdf_path and self.set_pdf(self.pdf_path, v - 1))
        row.addWidget(self.sp_page)
        row.addStretch(1)
        self.lb_pdf = QtWidgets.QLabel("")
        self.lb_pdf.setObjectName("hint")
        row.addWidget(self.lb_pdf)
        c1.body.addLayout(row)
        sl.addWidget(c1)

        # 2 jurisdiction
        c2 = Card(2, "Jurisdiction boundary")
        self.card_juris = c2
        row = QtWidgets.QHBoxLayout()
        self.ed_search = QtWidgets.QLineEdit()
        self.ed_search.setPlaceholderText("🔍  Search counties, cities, towns…")
        self.ed_search.textChanged.connect(self.filter_juris)
        self.cb_type = QtWidgets.QComboBox()
        self.cb_type.addItems(["All", "Counties", "Cities", "Towns"])
        self.cb_type.setCurrentIndex(1)
        self.cb_type.currentIndexChanged.connect(lambda _: self.filter_juris())
        row.addWidget(self.ed_search, 1)
        row.addWidget(self.cb_type)
        c2.body.addLayout(row)
        self.lst_juris = QtWidgets.QListWidget()
        self.lst_juris.setFixedHeight(150)
        self.lst_juris.itemDoubleClicked.connect(lambda _: self.load_boundary())
        c2.body.addWidget(self.lst_juris)
        row = QtWidgets.QHBoxLayout()
        self.lb_juris = QtWidgets.QLabel("Loading jurisdictions…")
        self.lb_juris.setObjectName("hint")
        self.lb_juris.setWordWrap(True)
        row.addWidget(self.lb_juris, 1)
        self.bt_boundary = QtWidgets.QPushButton("Load boundary")
        self.bt_boundary.clicked.connect(self.load_boundary)
        row.addWidget(self.bt_boundary)
        c2.body.addLayout(row)
        sl.addWidget(c2)

        # 3 LRS routes
        c3 = Card(3, "LRS routes")
        self.card_routes = c3
        grid = QtWidgets.QGridLayout()
        grid.setHorizontalSpacing(10)
        self.chk_groups = []
        for i, (label, cats, on, mcol, wt, pcol) in enumerate(ROUTE_GROUPS):
            cb = QtWidgets.QCheckBox(label)
            cb.setChecked(on)
            cb.setStyleSheet("QCheckBox { color: %s; }" % pcol)
            grid.addWidget(cb, i // 2, i % 2)
            self.chk_groups.append(cb)
        c3.body.addLayout(grid)
        row = QtWidgets.QHBoxLayout()
        self.chk_clip_routes = QtWidgets.QCheckBox("Clip to boundary")
        self.chk_clip_routes.setChecked(True)
        self.chk_clip_routes.setEnabled(HAVE_SHAPELY)
        self.chk_show_routes = QtWidgets.QCheckBox("Show in preview")
        self.chk_show_routes.setChecked(True)
        self.chk_show_routes.toggled.connect(self.rebuild_overlay)
        row.addWidget(self.chk_clip_routes)
        row.addWidget(self.chk_show_routes)
        row.addStretch(1)
        c3.body.addLayout(row)
        row = QtWidgets.QHBoxLayout()
        self.lb_routes = QtWidgets.QLabel("Not loaded")
        self.lb_routes.setObjectName("hint")
        row.addWidget(self.lb_routes, 1)
        self.bt_routes = QtWidgets.QPushButton("Fetch routes")
        self.bt_routes.clicked.connect(self.fetch_routes)
        row.addWidget(self.bt_routes)
        c3.body.addLayout(row)
        sl.addWidget(c3)

        # 4 Align
        c4 = Card(4, "Align PDF to boundary")
        self.card_align = c4
        row = QtWidgets.QHBoxLayout()
        self.rb_sim = QtWidgets.QRadioButton("Similarity")
        self.rb_aff = QtWidgets.QRadioButton("Affine")
        self.rb_sim.setChecked(True)
        self.rb_sim.setToolTip("Shift + uniform scale + rotation (recommended)")
        self.rb_aff.setToolTip("Adds independent X/Y scale and skew (projection differences)")
        row.addWidget(QtWidgets.QLabel("Model"))
        row.addWidget(self.rb_sim)
        row.addWidget(self.rb_aff)
        row.addStretch(1)
        self.chk_use_routes = QtWidgets.QCheckBox("Use routes")
        self.chk_use_routes.setToolTip("Also fit LRS routes to the road linework (after fetching routes)")
        row.addWidget(self.chk_use_routes)
        c4.body.addLayout(row)
        row = QtWidgets.QHBoxLayout()
        self.bt_auto = QtWidgets.QPushButton("✨ Auto-align")
        self.bt_auto.setObjectName("accent")
        self.bt_auto.clicked.connect(lambda: self.auto_align(full=True))
        self.bt_refine = QtWidgets.QPushButton("Refine")
        self.bt_refine.setToolTip("Fine-tune the current alignment to the map linework")
        self.bt_refine.clicked.connect(lambda: self.auto_align(full=False))
        self.bt_reset = QtWidgets.QPushButton("Reset")
        self.bt_reset.clicked.connect(self.reset_alignment)
        row.addWidget(self.bt_auto, 2)
        row.addWidget(self.bt_refine, 1)
        row.addWidget(self.bt_reset, 1)
        c4.body.addLayout(row)
        row = QtWidgets.QHBoxLayout()
        self.tool_group = QtWidgets.QButtonGroup(self)
        for i, (txt, tip) in enumerate((("✋ Pan", "Pan / zoom the preview (wheel)"),
                                        ("⤧ Adjust", "Drag to move overlay · Ctrl+wheel scale · Shift+wheel rotate · "
                                                     "arrows nudge · Alt = fine"),
                                        ("📍 Tie points", "Click the cyan boundary, then the matching spot on the PDF. "
                                                         "2+ pairs (3+ for affine) solve the fit. Right-click cancels."))):
            bt = QtWidgets.QPushButton(txt)
            bt.setObjectName("tool")
            bt.setCheckable(True)
            bt.setToolTip(tip)
            self.tool_group.addButton(bt, i)
            row.addWidget(bt)
        self.tool_group.button(0).setChecked(True)
        self.tool_group.buttonClicked[int].connect(self.set_tool)
        c4.body.addLayout(row)
        row = QtWidgets.QHBoxLayout()
        self.lb_fit = QtWidgets.QLabel("No alignment yet")
        self.lb_fit.setObjectName("metric")
        self.lb_fit.setWordWrap(True)
        row.addWidget(self.lb_fit, 1)
        self.bt_clear_pins = QtWidgets.QPushButton("Clear pins")
        self.bt_clear_pins.clicked.connect(self.clear_pins)
        row.addWidget(self.bt_clear_pins)
        c4.body.addLayout(row)
        sl.addWidget(c4)

        # 5 Output
        c5 = Card(5, "Place in MicroStation")
        self.card_out = c5
        grid = QtWidgets.QGridLayout()
        grid.setVerticalSpacing(6)
        self.chk_place_bnd = QtWidgets.QCheckBox("Boundary shapes")
        self.chk_place_routes = QtWidgets.QCheckBox("LRS routes")
        self.chk_place_raster = QtWidgets.QCheckBox("Map image")
        self.chk_clip = QtWidgets.QCheckBox("Mask outside county")
        self.chk_white = QtWidgets.QCheckBox("Transparent white")
        self.chk_holes = QtWidgets.QCheckBox("Mask holes")
        self.chk_pdf_off = QtWidgets.QCheckBox("Also attach PDF (off)")
        self.chk_scale_north = QtWidgets.QCheckBox("Scale bar && north arrow")
        self.chk_scale_north.setToolTip("Place a scale bar and a true-north arrow below the county's bottom-right corner")
        self.chk_pdf_off.setToolTip("Also attach the full original PDF page, georeferenced but turned off in all views")
        self.chk_place_raster.setToolTip("Render the PDF page to a PNG (saved next to the PDF) and attach it georeferenced")
        self.chk_white.setToolTip("Make white paper transparent so imagery/linework shows through")
        self.chk_clip.setToolTip("Make everything outside the county boundary transparent in the image")
        self.chk_holes.setToolTip("Exclude interior holes of the boundary (e.g. independent cities)")
        for cb, on in ((self.chk_place_bnd, True), (self.chk_place_routes, True), (self.chk_place_raster, True),
                       (self.chk_clip, True), (self.chk_white, False), (self.chk_holes, False),
                       (self.chk_pdf_off, True), (self.chk_scale_north, True)):
            cb.setChecked(on)
        grid.addWidget(self.chk_place_bnd, 0, 0)
        grid.addWidget(self.chk_place_routes, 0, 1)
        grid.addWidget(self.chk_place_raster, 1, 0)
        grid.addWidget(self.chk_clip, 1, 1)
        grid.addWidget(self.chk_white, 2, 0)
        grid.addWidget(self.chk_holes, 2, 1)
        grid.addWidget(self.chk_pdf_off, 3, 0)
        grid.addWidget(self.chk_scale_north, 3, 1)
        c5.body.addLayout(grid)
        self._fence_changed()
        form = QtWidgets.QFormLayout()
        form.setLabelAlignment(QtCore.Qt.AlignLeft)
        form.setHorizontalSpacing(12)
        trow = QtWidgets.QHBoxLayout()
        self.sl_transp = QtWidgets.QSlider(QtCore.Qt.Horizontal)
        self.sl_transp.setRange(0, 90)
        self.lb_transp = QtWidgets.QLabel("0 %")
        self.sl_transp.valueChanged.connect(lambda v: self.lb_transp.setText("%d %%" % v))
        trow.addWidget(self.sl_transp, 1)
        trow.addWidget(self.lb_transp)
        form.addRow("Raster transparency", trow)
        self.ed_lv = {}
        for key, lbl in (("boundary", "Boundary level"), ("routes", "Routes level"), ("raster", "Raster level"),
                         ("annotation", "Annotation level")):
            e = QtWidgets.QLineEdit(DEFAULT_LEVELS[key])
            self.ed_lv[key] = e
            if key == "routes":
                e.setToolTip("Level prefix. Routes go on one level per class, e.g. %s_Interstate, "
                             "_US_Primary, _State_Primary, _Secondary, _Ramps_Frontage, _Urban_Streets, _Other "
                             "(by-level colour and weight)." % DEFAULT_LEVELS[key])
            form.addRow(lbl, e)
        c5.body.addLayout(form)
        self.bt_place = QtWidgets.QPushButton("Place in MicroStation  ➜")
        self.bt_place.setObjectName("primary")
        self.bt_place.clicked.connect(self.place)
        c5.body.addWidget(self.bt_place)
        self.lb_out = QtWidgets.QLabel("")
        self.lb_out.setObjectName("hint")
        self.lb_out.setWordWrap(True)
        c5.body.addWidget(self.lb_out)
        sl.addWidget(c5)
        sl.addStretch(1)

        # ---- view area
        right = QtWidgets.QWidget()
        rl = QtWidgets.QVBoxLayout(right)
        rl.setContentsMargins(0, 0, 0, 0)
        rl.setSpacing(0)
        vb = QtWidgets.QFrame()
        vb.setObjectName("viewbar")
        vl = QtWidgets.QHBoxLayout(vb)
        vl.setContentsMargins(12, 6, 12, 6)
        for txt, fn in (("⤢ Fit", lambda: self.view.fit()), ("＋", lambda: self.view.scale(1.25, 1.25)),
                        ("－", lambda: self.view.scale(0.8, 0.8))):
            bt = QtWidgets.QPushButton(txt)
            bt.setObjectName("tool")
            bt.clicked.connect(fn)
            vl.addWidget(bt)
        vl.addSpacing(16)
        self.chk_overlay = QtWidgets.QCheckBox("Overlay")
        self.chk_overlay.setChecked(True)
        self.chk_overlay.toggled.connect(lambda v: self.view.overlay.setVisible(v and self.G is not None))
        vl.addWidget(self.chk_overlay)
        self.chk_region = QtWidgets.QCheckBox("Detected map area")
        self.chk_region.toggled.connect(self.toggle_region)
        vl.addWidget(self.chk_region)
        vl.addSpacing(16)
        self.bt_fence = QtWidgets.QPushButton("▭ Fence")
        self.bt_fence.setObjectName("tool")
        self.bt_fence.setCheckable(True)
        self.bt_fence.setToolTip("Drag a block fence on the PDF. When a fence is drawn, the map image, LRS routes\n"
                                 "and scale bar are clipped to it instead of the county boundary.")
        self.tool_group.addButton(self.bt_fence, MapView.MODE_FENCE)
        vl.addWidget(self.bt_fence)
        self.bt_fence_clear = QtWidgets.QPushButton("✕ Clear fence")
        self.bt_fence_clear.setObjectName("tool")
        self.bt_fence_clear.setToolTip("Remove the fence and clip to the county boundary again")
        self.bt_fence_clear.clicked.connect(self.clear_fence)
        vl.addWidget(self.bt_fence_clear)
        vl.addSpacing(16)
        vl.addWidget(QtWidgets.QLabel("PDF"))
        self.sl_pdf = QtWidgets.QSlider(QtCore.Qt.Horizontal)
        self.sl_pdf.setRange(10, 100)
        self.sl_pdf.setValue(100)
        self.sl_pdf.setFixedWidth(120)
        self.sl_pdf.valueChanged.connect(lambda v: self.view.set_pdf_opacity(v / 100.0))
        vl.addWidget(self.sl_pdf)
        vl.addStretch(1)
        self.lb_hint = QtWidgets.QLabel("Wheel = zoom · drag = pan")
        self.lb_hint.setObjectName("hint")
        vl.addWidget(self.lb_hint)
        rl.addWidget(vb)
        self.view = MapView(self)
        self.view.tiePicked.connect(self.add_pair)
        self.view.fileDropped.connect(lambda p: QtCore.QTimer.singleShot(0, lambda: self.set_pdf(p, 0)))
        self.view.fenceDrawn.connect(self.on_fence_drawn)
        rl.addWidget(self.view, 1)
        self.region_item = None
        split.addWidget(right)
        split.setStretchFactor(1, 1)
        split.setSizes([420, 1060])
        side.setMinimumWidth(max(300, sw.minimumSizeHint().width() + 14))
        try:
            sw0 = int(self._settings.value("sidebar_width", 0) or 0)
        except (TypeError, ValueError):
            sw0 = 0
        if sw0 >= 300:
            split.setSizes([sw0, max(400, 1480 - sw0)])
        split.splitterMoved.connect(
            lambda *_: self._settings.setValue("sidebar_width", split.sizes()[0]))

        # ---- status bar + log
        sb = QtWidgets.QFrame()
        sb.setObjectName("statusbar")
        sbl = QtWidgets.QVBoxLayout(sb)
        sbl.setContentsMargins(14, 6, 14, 6)
        sbl.setSpacing(4)
        row = QtWidgets.QHBoxLayout()
        self.lb_status = QtWidgets.QLabel("Ready")
        row.addWidget(self.lb_status, 1)
        self.progress = QtWidgets.QProgressBar()
        self.progress.setFixedWidth(260)
        self.progress.setRange(0, 1)
        self.progress.setValue(0)
        row.addWidget(self.progress)
        self.bt_log = QtWidgets.QPushButton("Log ▾")
        self.bt_log.setObjectName("tool")
        self.bt_log.setCheckable(True)
        row.addWidget(self.bt_log)
        sbl.addLayout(row)
        self.log_box = QtWidgets.QPlainTextEdit()
        self.log_box.setReadOnly(True)
        self.log_box.setFixedHeight(130)
        self.log_box.setVisible(False)
        self.bt_log.toggled.connect(self.log_box.setVisible)
        sbl.addWidget(self.log_box)
        root.addWidget(sb)
        self.update_state()

    # ------------------------------------------------------------------ helpers
    def _update_theme_button(self):
        dark = _THEME[0] == "dark"
        self.bt_theme.setText("☀  Light mode" if dark else "🌙  Dark mode")
        self.bt_theme.setToolTip("Switch to %s mode" % ("light" if dark else "dark"))

    def apply_theme(self, theme):
        _THEME[0] = theme if theme in THEMES else "dark"
        self.setStyleSheet(style_for(_THEME[0]))
        self._settings.setValue("theme", _THEME[0])
        self._update_theme_button()
        for c in self.findChildren(Card):
            c.restyle()
        self.lb_status.setStyleSheet("color:%s" % pal()["info"])
        self.view.viewport().update()

    def log(self, msg, level="info"):
        p = pal()
        col = p.get(level if level in ("ok", "warn", "err") else "info", p["info"])
        self.log_box.appendHtml('<span style="color:%s">%s</span>&nbsp; <span style="color:%s">%s</span>'
                                % (p["stamp"], time.strftime("%H:%M:%S"), col,
                                   str(msg).replace("&", "&amp;").replace("<", "&lt;")))
        self.lb_status.setText(msg)
        self.lb_status.setStyleSheet("color:%s" % col)
        _log_console(msg)
        if level == "err":
            self.bt_log.setChecked(True)

    def toast(self, msg):
        self.lb_status.setText(msg)
        self.lb_status.setStyleSheet("color:%s" % pal()["warn"])

    def pump(self):
        self._pumping = True
        try:
            QtWidgets.QApplication.processEvents(QtCore.QEventLoop.AllEvents, 30)
        finally:
            self._pumping = False

    def run_bg(self, fn, *args, **kw):
        """Run fn in a worker thread while keeping the UI responsive. MicroStation APIs must NOT be used in fn."""
        box = {}

        def worker():
            try:
                box["result"] = fn(*args, **kw)
            except Exception as ex:
                box["error"] = ex
                box["tb"] = traceback.format_exc()

        th = threading.Thread(target=worker, daemon=True)
        th.start()
        while th.is_alive():
            self.pump()
            th.join(0.03)
        if "error" in box:
            _log_console(box.get("tb", ""))
            raise box["error"]
        return box.get("result")

    def busy(self, msg, maximum=0):
        outer = self

        class _B:
            def __enter__(self_):
                outer._busy_depth += 1
                outer.sidebar.setEnabled(False)
                outer.view.setEnabled(False)
                QtWidgets.QApplication.setOverrideCursor(QtCore.Qt.BusyCursor)
                outer.progress.setRange(0, maximum)
                outer.progress.setValue(0)
                outer.log(msg)
                outer.pump()
                return self_

            def __exit__(self_, et, ev, tb):
                outer._busy_depth -= 1
                QtWidgets.QApplication.restoreOverrideCursor()
                if outer._busy_depth == 0:
                    outer.sidebar.setEnabled(True)
                    outer.view.setEnabled(True)
                    outer.progress.setRange(0, 1)
                    outer.progress.setValue(1 if et is None else 0)
                if et is not None:
                    outer.log("%s" % ev, "err")
                    _log_console("".join(traceback.format_exception(et, ev, tb)))
                outer.update_state()
                return et is not None and issubclass(et, Exception)
        return _B()

    def update_state(self):
        has_pdf = self.gray is not None
        has_bnd = len(self.rings_n) > 0
        has_G = self.G is not None
        self.bt_boundary.setEnabled(self.lst_juris.currentItem() is not None)
        self.bt_routes.setEnabled(has_bnd)
        self.bt_auto.setEnabled(has_pdf and has_bnd)
        self.bt_refine.setEnabled(has_pdf and has_bnd and has_G)
        self.bt_reset.setEnabled(has_G)
        for i in (1, 2):
            self.tool_group.button(i).setEnabled(has_G)
        self.bt_fence.setEnabled(bool(has_pdf))
        self.bt_clear_pins.setEnabled(bool(self.pairs))
        ok_place = has_bnd and self.conv.ok and HAVE_MSTN and (has_G or not self.chk_place_raster.isChecked())
        self.bt_place.setEnabled(bool(ok_place))
        if not HAVE_MSTN:
            self.bt_place.setToolTip("MicroStation API unavailable – " +
                                     (MSTN_IMPORT_ERROR or "run inside MicroStation to place elements."))
        self.chk_use_routes.setEnabled(bool(self.routes))
        if has_G:
            self.view.overlay.setVisible(self.chk_overlay.isChecked())

    # ------------------------------------------------------------------ jurisdictions
    @ui_action
    def load_jurisdictions(self):
        with self.busy("Loading Virginia jurisdictions from VDOT…"):
            self.juris = self.run_bg(self.client.list_jurisdictions)
            self.lb_juris.setText("%d jurisdictions available" % len(self.juris))
            self.log("Loaded %d jurisdictions." % len(self.juris), "ok")
            self.filter_juris()
            if self.pdf_path:
                self.autodetect_juris()

    @ui_action
    def filter_juris(self):
        txt = self.ed_search.text().strip().lower()
        typ = {1: "CO", 2: "CI", 3: "TO"}.get(self.cb_type.currentIndex())
        cur = self.lst_juris.currentItem().data(QtCore.Qt.UserRole)["OBJECTID"] if self.lst_juris.currentItem() else None
        self.lst_juris.clear()
        for j in self.juris:
            if typ and j.get("JURISTYPE") != typ:
                continue
            name = j.get("NAMELSAD") or j.get("NAME") or "?"
            code = j.get("VDOT_CODE") or ""
            if txt and txt not in name.lower() and txt not in str(code):
                continue
            it = QtWidgets.QListWidgetItem("%s   ·  %s %s" % (name, JURIS_TYPES.get(j.get("JURISTYPE"), ""),
                                                             ("#" + str(code)) if code else ""))
            it.setData(QtCore.Qt.UserRole, j)
            it.setToolTip("FIPS %s · District %s" % (j.get("STCOFIPS"), j.get("DISTRICT")))
            self.lst_juris.addItem(it)
            if cur is not None and j["OBJECTID"] == cur:
                self.lst_juris.setCurrentItem(it)
        self.update_state()

    def select_juris(self, j):
        if j.get("JURISTYPE") != {1: "CO", 2: "CI", 3: "TO"}.get(self.cb_type.currentIndex(), j.get("JURISTYPE")):
            self.cb_type.setCurrentIndex(0)
        self.ed_search.clear()
        self.filter_juris()
        for i in range(self.lst_juris.count()):
            it = self.lst_juris.item(i)
            if it.data(QtCore.Qt.UserRole)["OBJECTID"] == j["OBJECTID"]:
                self.lst_juris.setCurrentItem(it)
                self.lst_juris.scrollToItem(it, QtWidgets.QAbstractItemView.PositionAtCenter)
                break

    def autodetect_juris(self):
        j = guess_jurisdiction(self.pdf_path, self.juris)
        if j is None:
            self.lb_juris.setText("Could not detect jurisdiction from the file name – pick it from the list.")
            return False
        self.select_juris(j)
        self.lb_juris.setText("Auto-detected from file name: %s" % j.get("NAMELSAD"))
        self.log("Detected jurisdiction: %s (VDOT code %s)" % (j.get("NAMELSAD"), j.get("VDOT_CODE")), "ok")
        if not self.cur_juris or self.cur_juris["OBJECTID"] != j["OBJECTID"]:
            self.load_boundary()
        return True

    @ui_action
    def load_boundary(self):
        it = self.lst_juris.currentItem()
        if it is None:
            return
        j = it.data(QtCore.Qt.UserRole)
        if not self.conv.ok:
            self.log("The active model has no geographic coordinate system. Assign one (e.g. VA83-SF) first.", "err")
            return
        with self.busy("Downloading boundary for %s…" % j.get("NAMELSAD")):
            rings = self.run_bg(self.client.get_boundary, j["OBJECTID"], self.conv.datum_sr)
            self.cur_juris = j
            self.rings_ll = rings
            self.rings_w = [self.conv.to_world(r) for r in rings]
            allw = np.vstack(self.rings_w)
            lo, hi = allw.min(axis=0), allw.max(axis=0)
            self.norm_c = (lo + hi) / 2.0
            self.norm_S = float(max(hi - lo))
            self.rings_n = [(r - self.norm_c) / self.norm_S for r in self.rings_w]
            self.routes = []
            self.lb_routes.setText("Not loaded")
            self.card_routes.status.setText("")
            self.G = None
            self.G_auto = None
            self.pairs = []
            self.rebuild_overlay()
            nv = sum(len(r) for r in rings)
            self.card_juris.set_done(j.get("NAME") or "")
            self.log("Boundary loaded: %d ring(s), %d vertices." % (len(rings), nv), "ok")
            n_holes = len(split_rings(self.rings_w)[1])
            self.chk_holes.setToolTip(
                "Exclude interior holes of the boundary (e.g. independent cities). "
                + ("This boundary has %d hole(s)." % n_holes if n_holes else "This boundary has no holes."))
        if self.gray is not None and self.rings_n:
            if not self.try_restore_alignment():
                self.auto_align(full=True)

    # ------------------------------------------------------------------ routes
    def selected_categories(self):
        cats = []
        for cb, g in zip(self.chk_groups, ROUTE_GROUPS):
            if cb.isChecked():
                cats.extend(g[1])
        return cats

    @ui_action
    def fetch_routes(self):
        if not self.rings_ll:
            return False
        cats = self.selected_categories()
        if not cats:
            self.log("Select at least one route category.", "warn")
            return False
        allll = np.vstack(self.rings_ll)
        pad = 0.002
        bbox = (allll[:, 0].min() - pad, allll[:, 1].min() - pad, allll[:, 0].max() + pad, allll[:, 1].max() + pad)
        ok = False
        with self.busy("Downloading LRS routes…"):
            counter = {"n": 0}
            routes = self.run_bg(self.client.get_routes, bbox, cats, self.conv.datum_sr,
                                 lambda n: counter.__setitem__("n", n))
            n_raw = len(routes)
            if self.chk_clip_routes.isChecked():
                routes = self.run_bg(clip_routes_to_rings, routes, self.rings_ll)
            self.log("Projecting %d route parts…" % len(routes))
            out = []
            for i, (cat, name, ll) in enumerate(routes):
                w = self.conv.to_world(ll)
                out.append((cat, name, ll, w, (w - self.norm_c) / self.norm_S))
                if i % 200 == 0:
                    self.pump()
            self.routes = out
            nv = sum(len(r[2]) for r in out)
            self.lb_routes.setText("%d route parts · %d vertices (%d features fetched)" % (len(out), nv, counter["n"]))
            self.card_routes.set_done("%d" % len(out))
            self.log("LRS routes ready: %d parts from %d paths." % (len(out), n_raw), "ok")
            self.rebuild_overlay()
            ok = True
        return ok

    # ------------------------------------------------------------------ PDF
    @ui_action
    def browse_pdf(self):
        start = os.path.dirname(self.pdf_path) if self.pdf_path else os.path.expanduser("~")
        p, _ = QtWidgets.QFileDialog.getOpenFileName(self, "Open VDOT map PDF", start, "PDF files (*.pdf)")
        if p:
            self.set_pdf(p, 0)

    @ui_action
    def set_pdf(self, path, page=0):
        path = os.path.normpath(path)
        new_file = loaded = False
        with self.busy("Rendering preview of %s…" % os.path.basename(path)):
            out = os.path.join(self.tmpdir, "preview_%d.png" % int(time.time() * 1000))
            info = self.run_bg(render_pdf_page, path, out, PREVIEW_DPI, page)
            img = QtGui.QImage(out)
            if img.isNull():
                raise RuntimeError("Preview image could not be loaded.")
            new_file = path != self.pdf_path
            self.pdf_path, self.pdf_info, self.preview_png = path, info, out
            self.gray = qimage_to_gray(img)
            self.aligner = Aligner(self.gray, PREVIEW_DPI)
            self.ed_pdf.setText(path)
            self.ed_pdf.setToolTip(path)
            self.sp_page.blockSignals(True)
            self.sp_page.setRange(1, int(info["pages"]))
            self.sp_page.setValue(int(info["page"]) + 1)
            self.sp_page.blockSignals(False)
            self.lb_pdf.setText("%d page(s) · %.1f × %.1f in" % (info["pages"], info["w_pt"] / 72.0, info["h_pt"] / 72.0))
            self.view.set_image(img)
            self.view.set_pdf_opacity(self.sl_pdf.value() / 100.0)
            self.chk_region.setChecked(False)
            self.card_pdf.set_done(os.path.basename(path))
            self.log("PDF loaded: %s" % os.path.basename(path), "ok")
            if new_file:
                self.G = None
                self.pairs = []
                self.clear_fence()
            self.rebuild_overlay()
            loaded = True
        if not loaded:
            return
        if self.juris and new_file:
            if self.autodetect_juris():
                return
        if self.rings_n and self.G is None:
            if not self.try_restore_alignment():
                self.auto_align(full=True)

    @ui_action
    def toggle_region(self, on):
        if self.region_item is not None:
            self.view.scene().removeItem(self.region_item)
            self.region_item = None
        if on and self.aligner is not None:
            r = self.run_bg(self.aligner.region)
            h, w = r.shape
            rgba = np.zeros((h, w, 4), dtype=np.uint8)
            rgba[r] = (34, 211, 238, 70)
            img = QtGui.QImage(rgba.data, w, h, 4 * w, QtGui.QImage.Format_RGBA8888).copy()
            self.region_item = self.view.scene().addPixmap(QtGui.QPixmap.fromImage(img))
            self.region_item.setZValue(5)

    # ------------------------------------------------------------------ overlay & alignment
    @ui_action
    def rebuild_overlay(self):
        self.view.clear_overlay()
        if self.chk_show_routes.isChecked() and self.routes:
            by_group = {}
            for cat, name, ll, w, n in self.routes:
                gi = next((i for i, g in enumerate(ROUTE_GROUPS) if cat in g[1]), len(ROUTE_GROUPS) - 1)
                by_group.setdefault(gi, []).append(n)
            for gi, polys in by_group.items():
                self.view.add_overlay_paths(polys, ROUTE_GROUPS[gi][5], 1.2 if gi > 2 else 2.0, False, z=gi)
        if self.rings_n:
            self.view.add_overlay_paths(self.rings_n, "#0b1220", 5.0, True, z=20)
            self.view.add_overlay_paths(self.rings_n, "#22d3ee", 2.5, True, z=21)
        self.view.set_overlay_G(self.G)
        self.view.draw_pins(self.pairs, self.G)
        self.update_state()

    def compose_G(self, M):
        if self.G is None:
            return
        A = np.vstack([self.G, [0, 0, 1]])
        Mh = np.vstack([M, [0, 0, 1]])
        self.G = (Mh @ A)[:2]
        self.view.set_overlay_G(self.G)
        self.view.draw_pins(self.pairs, self.G)

    def after_manual_change(self):
        self.update_fit_label()

    def set_G(self, G, label=None):
        self.G = None if G is None else np.asarray(G, float)
        self.view.set_overlay_G(self.G)
        self.view.draw_pins(self.pairs, self.G)
        self.update_state()
        self.update_fit_label(label)

    def update_fit_label(self, prefix=None):
        if self.G is None or self.aligner is None:
            self.lb_fit.setText("No alignment yet")
            return
        mean_d, frac = self.aligner.score(self.G, self.rings_n)
        px_world = self.norm_S / math.sqrt(abs(np.linalg.det(self.G[:, :2])))
        ground = px_world * mean_d / max(1e-12, self.conv.uor_per_master) if HAVE_MSTN else px_world * mean_d
        txt = "%sFit: %.1f px mean · %d%% on linework · ≈ %.0f %s" % (
            (prefix + " · ") if prefix else "", mean_d, int(round(frac * 100)), ground, self.conv.master_label)
        self.lb_fit.setText(txt)
        self.card_align.set_done("aligned" if frac > 0.6 else "check fit", frac > 0.6)

    @ui_action
    def auto_align(self, full=True):
        if self.aligner is None or not self.rings_n:
            return
        affine = self.rb_aff.isChecked()
        route_lines = [r[4] for r in self.routes] if (self.chk_use_routes.isChecked() and self.routes) else None
        with self.busy("Auto-aligning PDF to boundary…" if full else "Refining alignment…"):
            if full or self.G is None:
                G0 = self.run_bg(self.aligner.initial_fit, self.rings_n)
                self.set_G(G0)
                self.pump()
                G = self.run_bg(self.aligner.refine, G0, self.rings_n, route_lines, affine, True)
            else:
                G = self.run_bg(self.aligner.refine, self.G, self.rings_n, route_lines, affine, False)
            self.G_auto = G.copy()
            self.set_G(G, "Auto" if full else "Refined")
            self.log("Alignment complete. %s" % self.lb_fit.text(), "ok")
            self.view.fit()

    @ui_action
    def reset_alignment(self):
        if self.G_auto is not None:
            self.set_G(self.G_auto.copy(), "Auto")
        else:
            self.auto_align(full=True)

    @ui_action
    def set_tool(self, i):
        self.view.set_mode(i)
        self.lb_hint.setText(["Wheel = zoom · drag = pan",
                              "Drag = move · Ctrl+wheel = scale · Shift+wheel = rotate · arrows = nudge",
                              "Click cyan boundary → click matching PDF spot · right-click = cancel",
                              "Drag a rectangle on the PDF to set the block fence · wheel = zoom"][i])
        self.view.setFocus()

    def nearest_overlay_vertex(self, p, tol):
        if self.G is None:
            return None
        polys = list(self.rings_n) + ([r[4] for r in self.routes] if self.chk_show_routes.isChecked() else [])
        best, bd = None, tol
        for poly in polys:
            pp = apply_G(self.G, poly)
            d = np.hypot(pp[:, 0] - p[0], pp[:, 1] - p[1])
            k = int(np.argmin(d))
            if d[k] < bd:
                best, bd = poly[k], d[k]
        return None if best is None else np.array(best)

    @ui_action
    def add_pair(self, w, p):
        self.pairs.append((np.asarray(w), np.asarray(p)))
        affine = self.rb_aff.isChecked()
        need = 3 if affine else 2
        if len(self.pairs) >= need:
            G, rms = solve_from_pairs([a for a, _ in self.pairs], [b for _, b in self.pairs], affine)
            if G is not None:
                self.set_G(G, "Tie points (RMS %.1f px)" % rms)
                self.log("Solved from %d tie points, RMS %.2f px." % (len(self.pairs), rms), "ok")
        else:
            self.view.draw_pins(self.pairs, self.G)
            self.toast("Tie point %d added – %d more needed." % (len(self.pairs), need - len(self.pairs)))
        self.update_state()

    @ui_action
    def clear_pins(self):
        self.pairs = []
        self.view.cancel_pending()
        self.view.draw_pins(self.pairs, self.G)
        self.update_state()

    # ------------------------------------------------------------------ sidecar (alignment persistence)
    def sidecar_path(self):
        return os.path.splitext(self.pdf_path)[0] + ".georef.json" if self.pdf_path else None

    def save_alignment(self):
        try:
            k = PREVIEW_DPI / 72.0
            wpt, hpt = self.pdf_info["w_pt"], self.pdf_info["h_pt"]
            grid = [(u, v) for u in np.linspace(0, wpt, 4) for v in np.linspace(0, hpt, 4)]
            pix = np.array(grid) * k
            Gh = np.vstack([self.G, [0, 0, 1]])
            inv = np.linalg.inv(Gh)
            wn = (pix @ inv[:2, :2].T) + inv[:2, 2]
            ll = self.conv.to_lonlat(wn * self.norm_S + self.norm_c)
            data = {"pdf": os.path.basename(self.pdf_path), "page": int(self.pdf_info["page"]),
                    "jurisdiction": self.cur_juris, "affine": bool(self.rb_aff.isChecked()),
                    "page_pt_to_lonlat": [[float(u), float(v), float(a), float(b)] for (u, v), (a, b) in zip(grid, ll)]}
            with open(self.sidecar_path(), "w", encoding="utf-8") as f:
                json.dump(data, f, indent=1)
        except Exception as ex:
            self.log("Could not save alignment sidecar: %s" % ex, "warn")

    def try_restore_alignment(self):
        sp = self.sidecar_path()
        if not sp or not os.path.isfile(sp) or not self.cur_juris:
            return False
        try:
            with open(sp, "r", encoding="utf-8") as f:
                data = json.load(f)
            if (data.get("jurisdiction") or {}).get("OBJECTID") != self.cur_juris.get("OBJECTID"):
                return False
            if int(data.get("page", 0)) != int(self.pdf_info["page"]):
                return False
            arr = np.array(data["page_pt_to_lonlat"], float)
            pix = arr[:, :2] * (PREVIEW_DPI / 72.0)
            wn = (self.conv.to_world(arr[:, 2:4]) - self.norm_c) / self.norm_S
            G, rms = solve_from_pairs(wn, pix, True)
            self.G_auto = G.copy()
            self.set_G(G, "Restored")
            self.log("Restored saved alignment from %s (residual %.2f px)." % (os.path.basename(sp), rms), "ok")
            self.view.fit()
            return True
        except Exception as ex:
            self.log("Ignoring alignment sidecar: %s" % ex, "warn")
            return False

    # ------------------------------------------------------------------ block fence
    def _fence_changed(self):
        fence = self.fence_px is not None
        self.chk_clip.setText("Mask outside fence" if fence else "Mask outside county")
        self.chk_clip.setToolTip("Make everything outside the %s transparent in the image"
                                 % ("block fence" if fence else "county boundary"))
        self.chk_holes.setEnabled(not fence)
        if hasattr(self, "bt_fence_clear"):
            self.bt_fence_clear.setEnabled(fence)

    def on_fence_drawn(self, r):
        self.fence_px = tuple(float(v) for v in r)
        self._fence_changed()
        dpi = float(PREVIEW_DPI)
        self.log("Block fence set: %.1f × %.1f in on the PDF. The image, routes and scale bar will be clipped "
                 "to it." % ((r[2] - r[0]) / dpi, (r[3] - r[1]) / dpi), "ok")

    def clear_fence(self):
        had = self.fence_px is not None
        self.fence_px = None
        self.view.set_fence_rect(None)
        self._fence_changed()
        if had:
            self.log("Block fence cleared – clipping to the county boundary.")

    def preview_to_world(self):
        """3×3 homogeneous transform: preview pixel -> world UOR (needs an alignment)."""
        H_g = np.vstack([self.G, [0, 0, 1]])
        H_w = np.array([[self.norm_S, 0, self.norm_c[0]], [0, self.norm_S, self.norm_c[1]], [0, 0, 1]])
        return H_w @ np.linalg.inv(H_g)

    def fence_world(self):
        if self.fence_px is None or self.G is None:
            return []
        x0, y0, x1, y1 = self.fence_px
        T = self.preview_to_world()
        c = np.array([[x0, y1, 1], [x1, y1, 1], [x1, y0, 1], [x0, y0, 1], [x0, y1, 1]], float)
        return [(T @ c.T).T[:, :2]]

    # ------------------------------------------------------------------ placement
    @ui_action
    def place(self):
        if not HAVE_MSTN:
            self.log("MicroStation is not available.", "err")
            return
        if self.fence_px is not None and self.G is None:
            self.log("Align the PDF before placing with a block fence.", "warn")
            return
        fence_w = self.fence_world()
        do_raster = self.chk_place_raster.isChecked()
        if do_raster and (self.G is None or self.pdf_path is None):
            self.log("Load a PDF and align it first.", "warn")
            return
        if self.chk_place_routes.isChecked() and not self.routes:
            if not self.fetch_routes():
                return
        with self.busy("Placing in MicroStation…", 4) as _:
            placer = MstnPlacer()
            lv = {k: (e.text().strip() or DEFAULT_LEVELS[k]) for k, e in self.ed_lv.items()}
            outers_w, holes_w = split_rings(self.rings_w)
            outers_n, holes_n = split_rings(self.rings_n)
            n_el = 0
            if self.chk_place_bnd.isChecked():
                lid = placer.level_id(lv["boundary"])
                for r in self.rings_w:
                    n_el += bool(placer.add_curve(r, True, lid, 3, 4))
                self.log("Placed %d boundary shape(s) on level %s." % (n_el, lv["boundary"]), "ok")
            self.progress.setValue(1)
            self.pump()
            if self.chk_place_routes.isChecked() and self.routes:
                routes = self.routes
                if fence_w:
                    if HAVE_SHAPELY:
                        fence_ll = [self.conv.to_lonlat(r) for r in fence_w]
                        clipped = self.run_bg(clip_routes_to_rings, [(c, nm, ll) for c, nm, ll, w, n in routes],
                                              fence_ll)
                        routes = [(c, nm, ll, self.conv.to_world(ll), None) for c, nm, ll in clipped]
                    else:
                        self.log("shapely is not installed, so routes were not clipped to the fence.", "warn")
                cnt = 0
                per_level = {}
                for i, (cat, name, ll, w, n) in enumerate(routes):
                    g = next((g for g in ROUTE_GROUPS if cat in g[1]), ROUTE_GROUPS[-1])
                    lname = route_level_name(lv["routes"], g[0])
                    if lname not in per_level:
                        per_level[lname] = [placer.level_id(lname, g[3], g[4], "VDOT LRS %s routes" % g[0]), 0]
                    if placer.add_curve(w, False, per_level[lname][0], g[3], g[4]):
                        cnt += 1
                        per_level[lname][1] += 1
                    if i % 250 == 0:
                        self.pump()
                self.log("Placed %d LRS route element(s) on per-class levels:%s" % (
                    cnt, "".join("\n    %s: %d" % (k, v[1]) for k, v in per_level.items() if v[1])), "ok")
            self.progress.setValue(2)
            self.pump()
            if do_raster:
                # preview pixel -> UOR, rescaled to the resolution of the attached image
                H_g = np.vstack([self.G, [0, 0, 1]])
                H_w = np.array([[self.norm_S, 0, self.norm_c[0]], [0, self.norm_S, self.norm_c[1]], [0, 0, 1]])
                T_pv = H_w @ np.linalg.inv(H_g)
                pv_h, pv_w = self.gray.shape[:2]

                def T_for(w, h):
                    return T_pv @ np.diag([pv_w / float(w), pv_h / float(h), 1.0])

                mask_rings = None
                if self.chk_clip.isChecked() and fence_w:
                    mask_rings = list(fence_w)
                elif self.chk_clip.isChecked():
                    mask_rings = list(outers_w)
                    if self.chk_holes.isChecked():
                        mask_rings += list(holes_w)
                        if not holes_w:
                            self.log("Mask holes: this boundary has no interior holes, so nothing was masked.", "warn")
                self.progress.setValue(3)
                self.pump()
                img_path, size = self.run_bg(
                    make_masked_image, self.pdf_path, int(self.pdf_info["page"]), self.pdf_info, T_for,
                    mask_rings, self.chk_white.isChecked(), self.tmpdir)
                lvl = placer.level_id(lv["raster"])
                placer.attach_raster(img_path, T_for, size, lvl, self.sl_transp.value())
                self.log("Map image attached: %s (%d ? %d px)%s" % (
                    os.path.basename(img_path), size[0], size[1],
                    ", everything outside the %s is transparent" % ("fence" if fence_w else "county")
                    if mask_rings else ""), "ok")
                if self.chk_pdf_off.isChecked():
                    k = PDF_ATTACH_DPI_GUESS / 72.0
                    guess = (max(1, int(round(self.pdf_info["w_pt"] * k))),
                             max(1, int(round(self.pdf_info["h_pt"] * k))))
                    psize = placer.attach_raster(self.pdf_path, T_for, guess, lvl, page=int(self.pdf_info["page"]),
                                                 visible=False)
                    self.log("Original PDF attached (%d ? %d px), turned off in all views ? "
                             "turn it on per view in Raster Manager." % psize, "ok")
                self.lb_out.setText("Image: %s" % img_path)
                self.save_alignment()
            if self.chk_scale_north.isChecked() and (fence_w or self.rings_w):
                try:
                    cnt, desc = placer.add_scale_and_north(fence_w or self.rings_w, self.conv,
                                                           placer.level_id(lv["annotation"]))
                    self.log("Placed scale bar (%s) and north arrow (%d elements) on level %s."
                             % (desc, cnt, lv["annotation"]), "ok")
                except Exception as ex:
                    self.log("Scale bar / north arrow failed: %s" % ex, "err")
            self.progress.setValue(4)
            if fence_w:
                try:
                    if MstnPlacer.define_fence(fence_w[0]):
                        self.log("Block fence also set as the active MicroStation fence.", "ok")
                except Exception as ex:
                    self.log("Could not set the MicroStation fence: %s" % ex, "warn")
            placer.fit_view()
            self.card_out.set_done("placed")
            self.log("Done – elements placed in %s." % ("model" if not HAVE_MSTN else
                                                       str(ISessionMgr.ActiveDgnModelRef.GetDgnModel().GetModelName())), "ok")

    def closeEvent(self, ev):
        try:
            import shutil
            shutil.rmtree(self.tmpdir, ignore_errors=True)
        except Exception:
            pass
        if getattr(sys, WIN_ATTR, None) is self:
            setattr(sys, WIN_ATTR, None)
        super().closeEvent(ev)


# =============================================================================
# Entry point (MicroStation cooperative event loop)
# =============================================================================
def _window_alive(w):
    try:
        return w is not None and w.isVisible()
    except RuntimeError:
        return False


def _keep_alive(token=None):
    if token is not None and getattr(sys, "_vdot_georef_loop", None) != token:
        return
    w = getattr(sys, WIN_ATTR, None)
    if not _window_alive(w):
        return
    try:
        if not w._pumping:
            QtWidgets.QApplication.processEvents()
            if HAVE_MSTN:
                PyCadInputQueue.PythonMainLoop()
    except Exception as ex:
        _log_console("keep_alive: %s" % ex)
    QtCore.QTimer.singleShot(100, lambda: _keep_alive(token))


def main():
    install_safety_hooks()
    _log_file("---- %s %s starting (MicroStation=%s, Python %s)" % (APP_TITLE, APP_VERSION, HAVE_MSTN,
                                                                    sys.version.split()[0]))
    if MSTN_IMPORT_ERROR:
        _log_file("MicroStation import failed: " + MSTN_IMPORT_ERROR)
    app = QtWidgets.QApplication.instance() or QtWidgets.QApplication(list(getattr(sys, "argv", None) or ["mstn"]))
    sys._vdot_georef_app = app
    w = getattr(sys, WIN_ATTR, None)
    if _window_alive(w) and type(w) is GeorefWindow:
        w.raise_()
        w.activateWindow()
        return w
    if _window_alive(w):
        # window from an earlier load of this script: its methods are stale code
        _log_file("Replacing window from a previous load of the script.")
        try:
            w.close()
            w.deleteLater()
        except Exception:
            pass
    # windows from older script versions (stored under the previous attribute name)
    old = getattr(sys, "_vdot_georef_window", None)
    if _window_alive(old):
        _log_file("Closing window from an older version of the script.")
        try:
            old.close()
            old.deleteLater()
        except Exception:
            pass
    sys._vdot_georef_window = None
    w = GeorefWindow()
    setattr(sys, WIN_ATTR, w)
    w.show()
    w.raise_()
    w.activateWindow()
    if MSTN_IMPORT_ERROR:
        w.log("Placing into MicroStation is disabled. " + MSTN_IMPORT_ERROR, "err")
    if HAVE_MSTN:
        token = object()
        sys._vdot_georef_loop = token
        QtCore.QTimer.singleShot(100, lambda: _keep_alive(token))
    else:
        app.exec_()
    return w


if not os.environ.get("VDOT_GEOREF_NO_AUTORUN"):
    main()
