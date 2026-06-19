#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
WO Flow Planning App
- Auto-loads station config from WO_Planning_App.xlsx (same folder as this script)
- Fetches Work Orders + QTY from SQL (v_MasterWOInfo / V_MasterRuncardInfo)
- Calculates estimated start/end dates per station per WO
- Production window: 09:00 – 00:00 (15 hrs/day)
- Manual start date/time override per WO
- Excel export
"""

import sys, os, math, re
from datetime import datetime, timedelta, time as dtime

import pandas as pd
import pyodbc

from PyQt5.QtCore import Qt, QThread, pyqtSignal, QAbstractTableModel, QModelIndex, QDate
from PyQt5.QtGui import QFont, QColor, QBrush, QIcon, QPixmap, QPainter, QPen, QLinearGradient
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QFileDialog, QMessageBox, QProgressBar,
    QTableView, QTableWidget, QTableWidgetItem, QHeaderView,
    QDialog, QFormLayout, QDateTimeEdit, QDialogButtonBox,
    QAbstractItemView, QStatusBar, QTabWidget,
)

# ── Paths ──────────────────────────────────────────────────────────────────────
# Works both as a .py script and as a PyInstaller frozen .exe
if getattr(sys, "frozen", False):
    APP_DIR = os.path.dirname(sys.executable)
else:
    APP_DIR = os.path.dirname(os.path.abspath(__file__))
STATION_EXCEL = os.path.join(APP_DIR, "WO_Planning_App.xlsx")

# ── DB ─────────────────────────────────────────────────────────────────────────
DB_CONN = (
    r"DRIVER=SQL Server;"
    r"SERVER=US_SQL01.USPL.HOME;"
    r"DATABASE=MES;"
    r"UID=LABVIEW;"
    r"PWD=LABVIEW;"
)

MATERIAL_NO      = "400454000035"
WO_FETCH_FROM    = "2026-05-01"   # WOs are always fetched from this fixed date onward
PROD_START_HOUR  = 9    # 09:00
PROD_END_HOUR    = 24   # midnight → 15 hrs/day
PROD_HOURS_DAY   = PROD_END_HOUR - PROD_START_HOUR  # 15

APP_AUTHOR       = "Rajavardhan Reddy Gogulamudi"

# ── Colors ─────────────────────────────────────────────────────────────────────
C_HEADER_BG  = QColor("#2C3E50")
C_HEADER_FG  = QColor("#FFFFFF")
C_DATE_BG    = QColor("#A8D5A2")   # green  – start date row
C_ENDDATE_BG = QColor("#F7DC6F")   # yellow – end date row
C_ACTUAL_BG  = QColor("#FAD7A0")   # orange – actual date row
C_INOUT_BG   = QColor("#AED6F1")   # blue   – units/day row

STATION_COLORS = [
    QColor("#D6EAF8"), QColor("#D5F5E3"), QColor("#FDEBD0"),
    QColor("#FCF3CF"), QColor("#EBDEF0"), QColor("#D0ECE7"),
    QColor("#FADBD8"), QColor("#E8DAEF"), QColor("#D1F2EB"),
    QColor("#FEF9E7"),
]

# ── SQL ─────────────────────────────────────────────────────────────────────────
WO_QUERY = """
SELECT WO, WOQTY, STARTDATE
FROM v_MasterWOInfo
WHERE MATERIALNO = '{material_no}'
  AND CAST(STARTDATE AS DATE) >= '{from_date}'
  AND CAST(STARTDATE AS DATE) <= '{to_date}'
ORDER BY STARTDATE, WO
"""

# One query per WO — gets Total + Passed for every OPERATION from TESTRESULT_800G_MASTER.
# TESTRESULT_800G_MASTER has: WO, COMPONENTID, OPERATION, TESTRESULT, LIV_MASTER_SID
MASTER_STATS_QUERY = """
WITH Latest AS (
    SELECT COMPONENTID, OPERATION, TESTRESULT,
           ROW_NUMBER() OVER (
               PARTITION BY COMPONENTID, OPERATION
               ORDER BY LIV_MASTER_SID DESC
           ) AS rn
    FROM TESTRESULT_800G_MASTER
    WHERE WO = '{wo}'
)
SELECT
    OPERATION,
    COUNT(*)  AS Total,
    SUM(CASE WHEN UPPER(LTRIM(RTRIM(TESTRESULT))) = 'PASS' THEN 1 ELSE 0 END) AS Passed
FROM Latest
WHERE rn = 1
GROUP BY OPERATION
"""

# FW Writing has no PASS/FAIL column — presence of record means pass.
FW_ACTUAL_QUERY = """
SELECT COUNT(DISTINCT T.COMPONENTID) AS Total
FROM TestResult_800G_2XFR4_FWWRITE_TEST T
INNER JOIN TESTRESULT_800G_MASTER M
       ON M.COMPONENTID = RTRIM(LTRIM(T.COMPONENTID))
WHERE M.WO = '{wo}'
"""

# TP2/TP3 is three sub-operations (RT, LT, HT) stored as separate OPERATION rows.
# Total = components that entered any TP2/TP3 op.
# Passed = components whose latest result is PASS for ALL three temperature ops.
TP2TP3_ACTUAL_QUERY = """
WITH Latest AS (
    SELECT COMPONENTID, OPERATION, TESTRESULT,
           ROW_NUMBER() OVER (
               PARTITION BY COMPONENTID, OPERATION
               ORDER BY LIV_MASTER_SID DESC
           ) AS rn
    FROM TESTRESULT_800G_MASTER
    WHERE WO = '{wo}'
      AND (OPERATION LIKE '%TP2%' OR OPERATION LIKE '%TP3%')
),
PerComp AS (
    SELECT COMPONENTID,
           COUNT(*)  AS TotalOps,
           SUM(CASE WHEN UPPER(LTRIM(RTRIM(TESTRESULT))) = 'PASS' THEN 1 ELSE 0 END) AS PassOps
    FROM Latest WHERE rn = 1
    GROUP BY COMPONENTID
)
SELECT
    COUNT(*)  AS Total,
    SUM(CASE WHEN PassOps = TotalOps THEN 1 ELSE 0 END) AS Passed
FROM PerComp
"""


# ══════════════════════════════════════════════════════════════════════════════
# UPH text parser
# Handles formats found in the Excel:
#   "60 devices/ 1 hr"     → 60.0 UPH
#   "16 devices/ 1hr"      → 16.0 UPH
#   "7 devices/ 1 hr"      → 7.0  UPH
#   "64 devices/ 2hrs"     → 32.0 UPH
#   "12 hr/ device"        → slots_qty / 12  UPH  (burn-in: 144 slots / 12 hr = 12)
#   plain number "32"      → 32.0 UPH
# ══════════════════════════════════════════════════════════════════════════════
def parse_uph(text: str, slots_qty: int = 1) -> float:
    t = str(text).strip().lower()
    # "N devices / X hr(s)"
    m = re.search(r'(\d+(?:\.\d+)?)\s*devices?\s*/\s*(\d+(?:\.\d+)?)\s*hrs?', t)
    if m:
        return float(m.group(1)) / float(m.group(2))
    # "X hr(s) / device"  → use slots as parallel capacity
    m = re.search(r'(\d+(?:\.\d+)?)\s*hrs?\s*/\s*devices?', t)
    if m:
        hrs_per_device = float(m.group(1))
        return slots_qty / hrs_per_device if hrs_per_device > 0 else 0.0
    # plain numeric fallback
    try:
        return float(t)
    except ValueError:
        return 0.0


def parse_batch_hrs(text: str) -> float:
    """Returns soak hours per device if the UPH string is 'X hr/device' style, else 0."""
    t = str(text).strip().lower()
    m = re.search(r'(\d+(?:\.\d+)?)\s*hrs?\s*/\s*devices?', t)
    return float(m.group(1)) if m else 0.0


# ══════════════════════════════════════════════════════════════════════════════
# Station Excel loader  (auto-reads STATION_EXCEL every time)
# ══════════════════════════════════════════════════════════════════════════════
def load_stations() -> tuple[list[dict], str]:
    """
    Returns (stations_list, error_message).
    stations_list entries: {name, station_qty, slots_qty, uph, devices_per_day}
    UPH is already the total effective throughput for the station group.
    """
    if not os.path.isfile(STATION_EXCEL):
        return [], f"Station file not found:\n{STATION_EXCEL}"
    try:
        df = pd.read_excel(STATION_EXCEL, header=0)
        df.columns = [str(c).strip() for c in df.columns]

        # flexible column detection
        def find_col(keywords, df_cols):
            kws = [k.lower() for k in keywords]
            for col in df_cols:
                cl = col.lower().replace(" ", "_").replace(" ", "")
                if all(k in cl for k in kws):
                    return col
            for col in df_cols:
                cl = col.lower()
                if any(k in cl for k in kws):
                    return col
            return None

        cols = list(df.columns)
        name_col  = find_col(["station", "name"], cols) or find_col(["name"], cols) or find_col(["station"], cols)
        st_qty_col = find_col(["station", "qty"], cols) or find_col(["station", "quantity"], cols)
        sl_qty_col = find_col(["slot", "qty"], cols) or find_col(["slot", "quantity"], cols) or find_col(["slots"], cols)
        uph_col   = find_col(["units", "hour"], cols) or find_col(["uph"], cols)

        if not name_col:
            return [], f"Cannot find Station Name column.\nFound: {cols}"
        if not uph_col:
            return [], f"Cannot find UPH column.\nFound: {cols}"

        stations = []
        for _, row in df.iterrows():
            name = str(row[name_col]).strip()
            if not name or name.lower() in ("nan", "none", ""):
                continue
            try:
                sq = int(row[st_qty_col]) if st_qty_col and pd.notna(row.get(st_qty_col)) else 1
            except Exception:
                sq = 1
            try:
                sl = int(row[sl_qty_col]) if sl_qty_col and pd.notna(row.get(sl_qty_col)) else sq
            except Exception:
                sl = sq
            uph_raw = row[uph_col] if uph_col else 0
            uph       = parse_uph(str(uph_raw), slots_qty=sl)
            batch_hrs = parse_batch_hrs(str(uph_raw))

            stations.append({
                "name": name,
                "station_qty": max(1, sq),
                "slots_qty": max(1, sl),
                "uph": uph,
                "batch_hrs": batch_hrs,          # > 0 → soak station (e.g. Burn-in)
                "devices_per_day": round(uph * PROD_HOURS_DAY),
                "uph_raw": str(uph_raw),
            })

        if not stations:
            return [], "No valid station rows found in the Excel file."
        return stations, ""
    except Exception as e:
        return [], str(e)


# ══════════════════════════════════════════════════════════════════════════════
# Programmatic app icon — drawn at runtime so no external .ico is needed
# ══════════════════════════════════════════════════════════════════════════════
def make_app_icon() -> QIcon:
    """Generates a 'WO' badge icon (multiple sizes for taskbar / window)."""
    icon = QIcon()
    for size in (16, 24, 32, 48, 64, 128, 256):
        pm = QPixmap(size, size)
        pm.fill(Qt.transparent)
        p = QPainter(pm)
        p.setRenderHint(QPainter.Antialiasing, True)
        p.setRenderHint(QPainter.TextAntialiasing, True)

        # Rounded-square gradient background (brand navy → teal)
        grad = QLinearGradient(0, 0, size, size)
        grad.setColorAt(0.0, QColor("#2C3E50"))
        grad.setColorAt(1.0, QColor("#16A085"))
        p.setBrush(grad)
        p.setPen(QPen(QColor("#1A252F"), max(1, size // 32)))
        radius = size * 0.22
        p.drawRoundedRect(1, 1, size - 2, size - 2, radius, radius)

        # 'WO' monogram centered
        p.setPen(QColor("#FFFFFF"))
        font = QFont("Segoe UI", int(size * 0.42), QFont.Black)
        p.setFont(font)
        p.drawText(pm.rect(), Qt.AlignCenter, "WO")

        p.end()
        icon.addPixmap(pm)
    return icon


# ══════════════════════════════════════════════════════════════════════════════
# Ensure station Excel exists next to the exe / script
# ══════════════════════════════════════════════════════════════════════════════
def ensure_station_excel() -> bool:
    """
    Returns True when STATION_EXCEL is ready to use.

    When frozen by PyInstaller (--add-data "WO_Planning_App.xlsx;."), the file
    is bundled inside the exe.  On the very first run it is copied out to
    APP_DIR so the user can edit it; after that the external copy is used.
    """
    if os.path.isfile(STATION_EXCEL):
        return True

    if getattr(sys, "frozen", False):
        import shutil
        meipass = getattr(sys, "_MEIPASS", "")
        src = os.path.join(meipass, "WO_Planning_App.xlsx")
        if os.path.isfile(src):
            shutil.copy2(src, STATION_EXCEL)
            return True

    return False


# ══════════════════════════════════════════════════════════════════════════════
# Scheduling logic
# ══════════════════════════════════════════════════════════════════════════════
def next_prod_time(dt: datetime) -> datetime:
    # Hours 0-8 are before shift (or just past midnight) → snap to 09:00 same day
    if dt.hour < PROD_START_HOUR:
        return dt.replace(hour=PROD_START_HOUR, minute=0, second=0, microsecond=0)
    # Hours 9-23 are within the shift → use as-is
    return dt


def schedule_station(start_dt: datetime, qty: int, uph: float,
                     batch_hrs: float = 0.0, slots: int = 1) -> dict:
    if uph <= 0:
        return {"start": start_dt, "end": start_dt, "days_needed": 0, "units_per_day": 0}

    units_per_day = round(uph * PROD_HOURS_DAY)
    dt = next_prod_time(start_dt)

    # ── Batch / soak station (e.g. Burn-in "12 hr/device") ────────────────────
    # All devices that fit within 'slots' go in together and soak for batch_hrs.
    # The process runs continuously overnight — it does NOT pause at shift end.
    if batch_hrs > 0:
        batches     = math.ceil(qty / slots) if slots > 0 else 1
        total_hours = batches * batch_hrs
        return {
            "start":         dt,
            "end":           dt + timedelta(hours=total_hours),
            "days_needed":   max(1, math.ceil(total_hours / 24)),
            "units_per_day": units_per_day,
        }

    # ── Flow / throughput station ──────────────────────────────────────────────
    total_hours   = qty / uph
    station_start = dt
    remaining     = total_hours

    while remaining > 0:
        # end of today's shift
        if PROD_END_HOUR == 24:
            eod = (dt + timedelta(days=1)).replace(hour=0, minute=0, second=0, microsecond=0)
        else:
            eod = dt.replace(hour=PROD_END_HOUR, minute=0, second=0, microsecond=0)
        avail = (eod - dt).total_seconds() / 3600
        if avail <= 0:
            dt = (dt + timedelta(days=1)).replace(hour=PROD_START_HOUR, minute=0, second=0, microsecond=0)
            continue
        used = min(remaining, avail)
        dt += timedelta(hours=used)
        remaining -= used
        if remaining > 0:
            # When PROD_END_HOUR==24, dt lands on 00:00 of the NEXT calendar day.
            # Adding another timedelta(days=1) would skip that day entirely, so
            # only advance by one day if we are NOT already at midnight.
            if dt.hour == 0 and dt.minute == 0 and dt.second == 0:
                dt = dt.replace(hour=PROD_START_HOUR, minute=0, second=0, microsecond=0)
            else:
                dt = (dt + timedelta(days=1)).replace(hour=PROD_START_HOUR, minute=0, second=0, microsecond=0)

    days_needed = math.ceil(total_hours / PROD_HOURS_DAY)
    return {"start": station_start, "end": dt, "days_needed": max(1, days_needed), "units_per_day": units_per_day}


def build_plan(work_orders: list, stations: list) -> list:
    """
    Batch scheduling: WOs that start on the same date are batched together.
    The station processes the combined quantity of the whole batch at once.
    All WOs in the same batch share identical station start/end times.

    Batches from different days queue at each station sequentially —
    the next batch can only enter a station once the previous batch has left.
    """
    from collections import defaultdict

    # ── Group WOs by planned start DATE ───────────────────────────────────
    groups: dict = defaultdict(list)
    for wo in work_orders:
        groups[wo["start_dt"].date()].append(wo)

    station_free_at: dict = {st["name"]: None for st in stations}
    plan = []

    for date_key in sorted(groups.keys()):
        batch      = groups[date_key]
        total_qty  = sum(wo["qty"] for wo in batch)
        # Batch is ready when the earliest planned start in the group is reached
        batch_ready = min(wo["start_dt"] for wo in batch)

        current_dt   = batch_ready
        batch_sched  = []   # one schedule entry per station, shared by all WOs in batch

        for st in stations:
            # Batch enters station when BOTH: batch is ready AND station is free
            station_free = station_free_at[st["name"]]
            entry_dt = station_free if (station_free and station_free > current_dt) else current_dt
            entry_dt = next_prod_time(entry_dt)

            # Schedule using combined quantity of all WOs in the batch
            sched = schedule_station(entry_dt, total_qty, st["uph"],
                                     batch_hrs=st.get("batch_hrs", 0.0),
                                     slots=st.get("slots_qty", 1))
            station_free_at[st["name"]] = sched["end"]

            batch_sched.append({
                "name":            st["name"],
                "uph":             st["uph"],
                "uph_raw":         st["uph_raw"],
                "devices_per_day": sched["units_per_day"],
                "start":           sched["start"],
                "end":             sched["end"],
                "days_needed":     sched["days_needed"],
                "batch_qty":       total_qty,   # total devices processed together
            })
            current_dt = sched["end"]

        # Every WO in the batch gets the same station schedule
        for wo in batch:
            plan.append({
                "wo":       wo["wo"],
                "qty":      wo["qty"],
                "start_dt": wo["start_dt"],
                "stations": list(batch_sched),  # shallow copy per WO
                "actual":   {},
            })

    return plan


# ══════════════════════════════════════════════════════════════════════════════
# DB fetch worker
# ══════════════════════════════════════════════════════════════════════════════
class FetchWorker(QThread):
    finished = pyqtSignal(object)
    error    = pyqtSignal(str)
    progress = pyqtSignal(int)

    def __init__(self, to_date: str):
        super().__init__()
        self.to_date = to_date

    def run(self):
        try:
            self.progress.emit(20)
            conn = pyodbc.connect(DB_CONN, timeout=15)
            self.progress.emit(50)
            df = pd.read_sql(
                WO_QUERY.format(material_no=MATERIAL_NO,
                                from_date=WO_FETCH_FROM,
                                to_date=self.to_date),
                conn,
            )
            conn.close()
            self.progress.emit(90)
            rows = [
                {
                    "wo":          str(r["WO"]),
                    "qty":         int(r["WOQTY"]) if pd.notna(r["WOQTY"]) else 0,
                    "create_date": r["STARTDATE"],
                }
                for _, r in df.iterrows()
            ]
            self.progress.emit(100)
            self.finished.emit(rows)
        except Exception as e:
            self.error.emit(str(e))


# ══════════════════════════════════════════════════════════════════════════════
# Actual data worker — fetches real tested/passed counts per WO per station
# ══════════════════════════════════════════════════════════════════════════════
class ActualDataWorker(QThread):
    """Runs station queries for every WO and emits {wo: {station_name: (total, passed)}}."""
    finished = pyqtSignal(object)
    error    = pyqtSignal(str)
    progress = pyqtSignal(int, int)   # current, total

    def __init__(self, plan: list, stations: list):
        super().__init__()
        self._plan     = plan
        self._stations = stations

    def _fw_station(self):
        for s in self._stations:
            sn = s["name"].upper()
            if "FW" in sn or "FIRMWARE" in sn or "FWWRITE" in sn:
                return s
        return None

    def _tp2tp3_station(self):
        for s in self._stations:
            sn = s["name"].upper()
            if "TP2" in sn or "TP3" in sn:
                return s
        return None

    def run(self):
        try:
            conn   = pyodbc.connect(DB_CONN, timeout=30)
            total  = len(self._plan)
            result = {}
            fw_st      = self._fw_station()
            tp2tp3_st  = self._tp2tp3_station()

            for idx, wo_data in enumerate(self._plan):
                wo = wo_data["wo"]
                result[wo] = {}

                # ── Step 1: all non-FW stations from TESTRESULT_800G_MASTER ──
                try:
                    df = pd.read_sql(MASTER_STATS_QUERY.format(wo=wo), conn)
                    for _, row in df.iterrows():
                        op     = str(row["OPERATION"]).strip()
                        tested = int(row["Total"])
                        passed = int(row["Passed"])
                        # Match OPERATION to station name (exact first, then fuzzy)
                        matched = None
                        for s in self._stations:
                            if s["name"].strip() == op:
                                matched = s["name"]
                                break
                        if matched is None:
                            op_up = op.upper()
                            for s in self._stations:
                                if s["name"].upper().replace(" ", "_") == op_up.replace(" ", "_"):
                                    matched = s["name"]
                                    break
                        if matched:
                            result[wo][matched] = (tested, passed)
                except Exception:
                    pass

                # ── Step 2: FW Writing — no PASS/FAIL, count = all pass ──────
                if fw_st:
                    try:
                        fw_df = pd.read_sql(FW_ACTUAL_QUERY.format(wo=wo), conn)
                        count = int(fw_df["Total"].iloc[0]) if len(fw_df) else 0
                        result[wo][fw_st["name"]] = (count, count)
                    except Exception:
                        result[wo][fw_st["name"]] = (0, 0)

                # ── Step 3: TP2/TP3 — aggregate RT + LT + HT into one result ─
                if tp2tp3_st:
                    try:
                        tp_df  = pd.read_sql(TP2TP3_ACTUAL_QUERY.format(wo=wo), conn)
                        tested = int(tp_df["Total"].iloc[0])  if len(tp_df) else 0
                        passed = int(tp_df["Passed"].iloc[0]) if len(tp_df) else 0
                        result[wo][tp2tp3_st["name"]] = (tested, passed)
                    except Exception:
                        result[wo][tp2tp3_st["name"]] = (0, 0)

                self.progress.emit(idx + 1, total)

            conn.close()
            self.finished.emit(result)
        except Exception as e:
            self.error.emit(str(e))


# ══════════════════════════════════════════════════════════════════════════════
# Table model
# ══════════════════════════════════════════════════════════════════════════════
ROW_LABELS  = ["Start Date", "End Date", "Actual Date",
               # "Units/Day",     # reserved – uncomment to re-enable
               "WO In / Out", "Yield %"]
ROWS_PER_WO = len(ROW_LABELS)


class PlanTableModel(QAbstractTableModel):
    def __init__(self, plan: list, stations: list):
        super().__init__()
        self._plan     = plan
        self._stations = stations
        self._rows     = []
        for wi, _ in enumerate(plan):
            for ri, lbl in enumerate(ROW_LABELS):
                self._rows.append((wi, ri, lbl))

    def rowCount(self, _=QModelIndex()):   return len(self._rows)
    def columnCount(self, _=QModelIndex()): return 3 + len(self._stations)

    def headerData(self, section, orientation, role=Qt.DisplayRole):
        if orientation == Qt.Horizontal:
            if role == Qt.DisplayRole:
                hdrs = ["Testing WO#", "QTY In", ""] + [s["name"] for s in self._stations]
                return hdrs[section] if section < len(hdrs) else ""
            if role == Qt.BackgroundRole: return QBrush(C_HEADER_BG)
            if role == Qt.ForegroundRole: return QBrush(C_HEADER_FG)
            if role == Qt.FontRole:       return QFont("Segoe UI", 9, QFont.Bold)
            if role == Qt.TextAlignmentRole: return Qt.AlignCenter
        return None

    def _in_qty(self, wo_data: dict, actual: dict, st_idx: int) -> int:
        """Return the input quantity for station st_idx.
        First station → WO qty.  Others → previous station's passed count."""
        if st_idx == 0:
            return wo_data["qty"]
        prev_name = wo_data["stations"][st_idx - 1]["name"]
        _, prev_passed = actual.get(prev_name, (0, 0))
        return prev_passed

    def _yield_color(self, tested, passed):
        """Return background QColor for a yield cell."""
        if tested == 0:
            return QColor("#D5D8DC")          # grey – no data yet
        pct = passed / tested * 100
        if pct >= 95:  return QColor("#A9DFBF")   # green
        if pct >= 80:  return QColor("#FAD7A0")   # orange
        return QColor("#F1948A")                  # red

    def data(self, index, role=Qt.DisplayRole):
        if not index.isValid():
            return None
        row, col = index.row(), index.column()
        wi, ri, lbl = self._rows[row]
        wo_data = self._plan[wi]
        st_idx  = col - 3

        # ── actual data helpers ────────────────────────────────────────────
        actual = wo_data.get("actual", {})   # {station_name: (tested, passed)}

        if role == Qt.DisplayRole:
            if col == 0: return wo_data["wo"] if ri == 0 else ""
            if col == 1: return str(wo_data["qty"]) if ri == 0 else ""
            if col == 2: return lbl
            if st_idx < 0 or st_idx >= len(wo_data["stations"]): return ""
            st = wo_data["stations"][st_idx]
            if ri == 0: return st["start"].strftime("%m/%d  %H:%M")
            if ri == 1: return st["end"].strftime("%m/%d  %H:%M")
            if ri == 2: return ""   # QDateTimeEdit widget is embedded here
            # if ri == 3: return str(st["devices_per_day"])  # Units/Day – reserved
            # ri == 3 : WO In / Out
            # In  = previous station's passed count (WO qty for first station)
            # Out = this station's passed count from DB
            if ri == 3:
                _, passed = actual.get(st["name"], (0, 0))
                in_qty = self._in_qty(wo_data, actual, st_idx)
                if in_qty == 0 and passed == 0: return "— / —"
                return f"{in_qty}  /  {passed}"
            # ri == 4 : Yield %  = Out / In × 100
            if ri == 4:
                _, passed = actual.get(st["name"], (0, 0))
                in_qty = self._in_qty(wo_data, actual, st_idx)
                if in_qty == 0: return "—"
                return f"{passed / in_qty * 100:.1f}%"

        if role == Qt.BackgroundRole:
            if col == 0: return QBrush(QColor("#FADBD8"))
            if col == 1: return QBrush(QColor("#D6EAF8"))
            if col == 2:
                row_bg = [C_DATE_BG, C_ENDDATE_BG, C_ACTUAL_BG,
                          # C_INOUT_BG,         # Units/Day – reserved
                          QColor("#AED6F1"), QColor("#A9DFBF")]
                return QBrush(row_bg[ri])
            if st_idx < 0 or st_idx >= len(wo_data["stations"]):
                return None
            st = wo_data["stations"][st_idx]
            if ri == 4:   # yield – colour by performance
                _, passed = actual.get(st["name"], (0, 0))
                in_qty = self._in_qty(wo_data, actual, st_idx)
                return QBrush(self._yield_color(in_qty, passed))
            if ri == 3:   # in/out
                return QBrush(QColor("#D6EAF8"))
            sc = STATION_COLORS[st_idx % len(STATION_COLORS)]
            tint = [1.0, 0.88, 0.78, 0.65, 0.60][ri]  # Units/Day tint removed
            def t(c): return int(c + (255 - c) * (1 - tint))
            return QBrush(QColor(t(sc.red()), t(sc.green()), t(sc.blue())))

        if role == Qt.TextAlignmentRole: return Qt.AlignCenter
        if role == Qt.FontRole:
            f = QFont("Segoe UI", 9)
            if col in (0, 1): f.setBold(True)
            if ri == 4: f.setBold(True)   # yield bold
            return f
        if role == Qt.ForegroundRole:
            if col == 0: return QBrush(QColor("#C0392B"))
            if ri == 4:
                if st_idx >= 0 and st_idx < len(wo_data["stations"]):
                    _, passed = actual.get(
                        wo_data["stations"][st_idx]["name"], (0, 0))
                    in_qty = self._in_qty(wo_data, actual, st_idx)
                    if in_qty > 0 and passed / in_qty < 0.80:
                        return QBrush(QColor("#922B21"))  # dark red for bad yield
            return QBrush(QColor("#1A252F"))
        return None


# ══════════════════════════════════════════════════════════════════════════════
# Start time override dialog
# ══════════════════════════════════════════════════════════════════════════════
class StartTimeDialog(QDialog):
    def __init__(self, wo_list: list, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Update WO Start Date / Time")
        self.setMinimumWidth(500)
        layout = QVBoxLayout(self)

        info = QLabel("Set a custom start date & time for each Work Order.\n"
                      "Leave unchanged to keep the original create date at 09:00.")
        info.setWordWrap(True)
        layout.addWidget(info)

        scroll_widget = QWidget()
        form = QFormLayout(scroll_widget)
        form.setLabelAlignment(Qt.AlignRight)
        self._editors = {}
        for wo in wo_list:
            dte = QDateTimeEdit()
            dte.setDisplayFormat("yyyy-MM-dd  HH:mm")
            dte.setCalendarPopup(True)
            sdt = wo["start_dt"]
            if not isinstance(sdt, datetime):
                sdt = datetime.now()
            dte.setDateTime(sdt.replace(tzinfo=None))
            self._editors[wo["wo"]] = dte
            form.addRow(f"WO  {wo['wo']}:", dte)
        layout.addWidget(scroll_widget)

        btn = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        btn.accepted.connect(self.accept)
        btn.rejected.connect(self.reject)
        layout.addWidget(btn)

    def get_overrides(self) -> dict:
        return {wo: dte.dateTime().toPyDateTime() for wo, dte in self._editors.items()}


# ══════════════════════════════════════════════════════════════════════════════
# Main window
# ══════════════════════════════════════════════════════════════════════════════
class WOPlanningApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(f"WO Flow Planning — 800G FR4   •   by {APP_AUTHOR}")
        self.setWindowIcon(make_app_icon())
        self.resize(1440, 740)
        self._wo_rows:    list = []
        self._stations:   list = []
        self._plan:       list = []
        self._actual_dates: dict = {}   # {(wo, station_name): datetime} – user-edited actual dates
        self._build_ui()
        self._auto_load_stations()

    # ── build UI ───────────────────────────────────────────────────────────
    def _build_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        root = QVBoxLayout(central)
        root.setSpacing(6)
        root.setContentsMargins(8, 8, 8, 8)

        # ── title ──────────────────────────────────────────────────────────
        title_bar = QWidget()
        title_bar.setStyleSheet("background:#2C3E50; border-radius:4px;")
        tb_layout = QHBoxLayout(title_bar)
        tb_layout.setContentsMargins(14, 8, 14, 8)

        title = QLabel(
            f"WO Flow Planning Dashboard — 800G FR4  (MATERIALNO: {MATERIAL_NO})"
        )
        title.setStyleSheet("color:white; font-size:15px; font-weight:bold; background:transparent;")
        tb_layout.addWidget(title)
        tb_layout.addStretch()

        author_lbl = QLabel(f"by {APP_AUTHOR}")
        author_lbl.setStyleSheet(
            "color:#AED6F1; font-size:11px; font-style:italic;"
            "font-weight:normal; background:transparent;"
        )
        tb_layout.addWidget(author_lbl)

        root.addWidget(title_bar)

        # ── control bar ────────────────────────────────────────────────────
        ctrl = QHBoxLayout()
        ctrl.setSpacing(8)

        _today = QDate.currentDate()
        self._from_date_edit = QDateTimeEdit(_today)   # defaults to today
        self._from_date_edit.setDisplayFormat("yyyy-MM-dd")
        self._from_date_edit.setCalendarPopup(True)
        self._from_date_edit.setFixedWidth(130)
        ctrl.addWidget(QLabel(f"Up to Date  (from {WO_FETCH_FROM}):"))
        ctrl.addWidget(self._from_date_edit)

        def btn(label, color, slot, enabled=True):
            b = QPushButton(label)
            b.setFixedHeight(32)
            b.setStyleSheet(
                f"background:{color}; color:white; font-weight:bold; border-radius:4px;"
                f"padding:0 12px;"
            )
            b.clicked.connect(slot)
            b.setEnabled(enabled)
            return b

        self._btn_fetch    = btn("Fetch WOs from DB",  "#2980B9", self._fetch_wo)
        self._btn_reload   = btn("Reload Stations",     "#27AE60", self._reload_stations)
        self._btn_calc     = btn("Calculate Plan",       "#8E44AD", self._calculate_plan,  False)
        self._btn_export = btn("Export to Excel", "#16A085", self._export_excel, False)

        for b in (self._btn_fetch, self._btn_reload, self._btn_calc, self._btn_export):
            ctrl.addWidget(b)
        ctrl.addStretch()
        root.addLayout(ctrl)

        # ── info strip ─────────────────────────────────────────────────────
        info_row = QHBoxLayout()
        self._lbl_wo_count = QLabel("WOs: —")
        self._lbl_st_count = QLabel("Stations: —")
        self._lbl_win      = QLabel(
            f"Production: {PROD_START_HOUR:02d}:00 – 00:00  ({PROD_HOURS_DAY} hrs/day)"
        )
        self._lbl_file     = QLabel(f"Station file: {os.path.basename(STATION_EXCEL)}")
        for lbl in (self._lbl_wo_count, self._lbl_st_count, self._lbl_win, self._lbl_file):
            lbl.setStyleSheet("color:#555; font-size:10px;")
        for w in (self._lbl_wo_count, QLabel("|"), self._lbl_st_count,
                  QLabel("|"), self._lbl_win, QLabel("|"), self._lbl_file):
            info_row.addWidget(w)
        info_row.addStretch()
        root.addLayout(info_row)

        # ── progress bar ───────────────────────────────────────────────────
        self._progress = QProgressBar()
        self._progress.setFixedHeight(6)
        self._progress.setTextVisible(False)
        self._progress.setStyleSheet("QProgressBar::chunk{background:#3498DB;}")
        self._progress.hide()
        root.addWidget(self._progress)

        # ── tab widget ─────────────────────────────────────────────────────
        self._tabs = QTabWidget()
        self._tabs.setStyleSheet("""
            QTabWidget::pane {
                border: 1px solid #BDC3C7;
                border-radius: 3px;
            }
            QTabBar::tab {
                background: #D5D8DC;
                color: #2C3E50;
                font-weight: bold;
                font-size: 10px;
                padding: 7px 18px;
                margin-right: 2px;
                border-top-left-radius: 4px;
                border-top-right-radius: 4px;
            }
            QTabBar::tab:selected {
                background: #2C3E50;
                color: white;
            }
            QTabBar::tab:hover:!selected {
                background: #ABB2B9;
            }
        """)
        root.addWidget(self._tabs)

        # ── Tab 1 : Fetched Work Orders ────────────────────────────────────
        tab_wo = QWidget()
        vbox_wo = QVBoxLayout(tab_wo)
        vbox_wo.setContentsMargins(6, 6, 6, 6)

        lbl_wo = QLabel(
            "Work orders fetched from the database.  "
            "Edit  'Planned Start'  to override the start date/time, "
            "then click  'Calculate Plan'  to refresh."
        )
        lbl_wo.setWordWrap(True)
        lbl_wo.setStyleSheet("color:#555; font-size:10px; padding:2px 0;")
        vbox_wo.addWidget(lbl_wo)

        # 4 columns — the 4th hosts an inline QDateTimeEdit per WO
        self._wo_table = QTableWidget(0, 4)
        self._wo_table.setHorizontalHeaderLabels(
            ["Work Order #", "Quantity", "DB Start Date", "Planned Start  (editable)"]
        )
        self._wo_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        self._wo_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        self._wo_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeToContents)
        self._wo_table.horizontalHeader().setSectionResizeMode(3, QHeaderView.Stretch)
        self._wo_table.verticalHeader().setVisible(False)
        self._wo_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self._wo_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self._wo_table.setAlternatingRowColors(True)
        self._wo_table.setStyleSheet("""
            QTableWidget {
                border: none; font-size: 11px;
                alternate-background-color: #EBF5FB;
                gridline-color: #D5D8DC;
            }
            QHeaderView::section {
                background: #2980B9; color: white;
                font-weight: bold; font-size: 11px;
                padding: 6px; border: 1px solid #1A6B99;
            }
            QTableWidget::item { padding: 4px; }
        """)
        vbox_wo.addWidget(self._wo_table)
        self._tabs.addTab(tab_wo, "📋  Work Orders")

        # ── Tab 2 : Station Configuration ─────────────────────────────────
        tab_st = QWidget()
        vbox_st = QVBoxLayout(tab_st)
        vbox_st.setContentsMargins(6, 6, 6, 6)

        st_info_row = QHBoxLayout()
        lbl_st = QLabel(
            f"Station config from  <b>{os.path.basename(STATION_EXCEL)}</b>.  "
            f"Click <b>Edit</b> to modify, then <b>Save</b> to write back to file."
        )
        lbl_st.setStyleSheet("color:#555; font-size:10px; padding:2px 0;")
        lbl_st.setWordWrap(True)
        st_info_row.addWidget(lbl_st, stretch=1)

        self._btn_edit_st = QPushButton("✏  Edit Stations")
        self._btn_edit_st.setFixedHeight(30)
        self._btn_edit_st.setStyleSheet(
            "background:#E67E22; color:white; font-weight:bold; border-radius:4px; padding:0 12px;")
        self._btn_edit_st.clicked.connect(self._toggle_station_edit)
        st_info_row.addWidget(self._btn_edit_st)

        self._btn_save_st = QPushButton("💾  Save Stations")
        self._btn_save_st.setFixedHeight(30)
        self._btn_save_st.setStyleSheet(
            "background:#27AE60; color:white; font-weight:bold; border-radius:4px; padding:0 12px;")
        self._btn_save_st.clicked.connect(self._save_stations)
        self._btn_save_st.setEnabled(False)
        st_info_row.addWidget(self._btn_save_st)
        vbox_st.addLayout(st_info_row)

        self._st_table = QTableWidget(0, 5)
        self._st_table.setHorizontalHeaderLabels(
            ["Station Name", "Station Qty", "Slots Qty", "UPH  (e.g. '60 devices/1hr')",
             "Devices / Day  (auto-calculated)"]
        )
        self._st_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        self._st_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        self._st_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeToContents)
        self._st_table.horizontalHeader().setSectionResizeMode(3, QHeaderView.Stretch)
        self._st_table.horizontalHeader().setSectionResizeMode(4, QHeaderView.ResizeToContents)
        self._st_table.verticalHeader().setVisible(False)
        self._st_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self._st_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self._st_table.setAlternatingRowColors(True)
        self._st_table.setStyleSheet("""
            QTableWidget {
                border: none; font-size: 11px;
                alternate-background-color: #EAFAF1;
                gridline-color: #D5D8DC;
            }
            QHeaderView::section {
                background: #27AE60; color: white;
                font-weight: bold; font-size: 11px;
                padding: 6px; border: 1px solid #1E8449;
            }
            QTableWidget::item { padding: 4px; }
            QTableWidget::item:selected { background: #AED6F1; color: #1A252F; }
        """)
        vbox_st.addWidget(self._st_table)
        self._tabs.addTab(tab_st, "🏭  Station Config")

        # ── Tab 3 : WO Flow Plan ───────────────────────────────────────────
        tab_plan = QWidget()
        vbox_plan = QVBoxLayout(tab_plan)
        vbox_plan.setContentsMargins(6, 6, 6, 6)

        lbl_plan = QLabel("Estimated flow plan per station.  "
                          "Click  'Calculate Plan'  to generate.")
        lbl_plan.setStyleSheet("color:#555; font-size:10px; padding:2px 0;")
        vbox_plan.addWidget(lbl_plan)

        self._table = QTableView()
        self._table.setAlternatingRowColors(False)
        self._table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self._table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
        self._table.horizontalHeader().setMinimumSectionSize(100)
        self._table.verticalHeader().setVisible(False)
        self._table.setShowGrid(True)
        self._table.setStyleSheet("""
            QTableView {
                border: none; font-size: 10px;
                gridline-color: #BDC3C7;
            }
            QHeaderView::section {
                background: #2C3E50; color: white;
                font-weight: bold; font-size: 10px;
                padding: 6px; border: 1px solid #1A252F;
            }
        """)
        vbox_plan.addWidget(self._table)
        self._tabs.addTab(tab_plan, "📅  WO Flow Plan")

        # ── status bar ─────────────────────────────────────────────────────
        self._status = QStatusBar()
        self.setStatusBar(self._status)
        self._status.showMessage("Ready.")

    # ── auto load stations on startup ──────────────────────────────────────
    def _auto_load_stations(self):
        if not ensure_station_excel():
            msg = (
                "Station config file not found.\n\n"
                f"Expected location:\n{STATION_EXCEL}\n\n"
                "Place  WO_Planning_App.xlsx  in the same folder as this application,\n"
                "then click  'Reload Stations'."
            )
            self._lbl_st_count.setText("Stations: MISSING")
            self._status.showMessage(f"Station file not found: {STATION_EXCEL}")
            QMessageBox.warning(self, "Station File Missing", msg)
            return

        stations, err = load_stations()
        if err:
            self._lbl_st_count.setText("Stations: ERROR")
            self._status.showMessage(f"Station load error: {err}")
            QMessageBox.warning(self, "Station File Warning", err)
        else:
            self._stations = stations
            self._lbl_st_count.setText(f"Stations: {len(stations)}")
            self._populate_station_table(stations)
            self._tabs.setCurrentIndex(1)   # jump to Station Config tab
            self._status.showMessage(
                f"Loaded {len(stations)} stations from {os.path.basename(STATION_EXCEL)}"
            )
            self._check_ready()

    def _populate_station_table(self, stations: list):
        self._st_table.setRowCount(0)
        row_colors = ["#FDFEFE", "#EBF5FB"]
        for i, st in enumerate(stations):
            r = self._st_table.rowCount()
            self._st_table.insertRow(r)
            bg = QColor(row_colors[i % 2])

            cells = [
                st["name"],
                str(st["station_qty"]),
                str(st["slots_qty"]),
                st["uph_raw"],
                str(st["devices_per_day"]),
            ]
            for c, text in enumerate(cells):
                item = QTableWidgetItem(text)
                item.setTextAlignment(Qt.AlignCenter)
                item.setBackground(QBrush(bg))
                if c == 0:
                    item.setForeground(QBrush(QColor("#1A5276")))
                    item.setFont(QFont("Segoe UI", 10, QFont.Bold))
                elif c == 4:
                    item.setForeground(QBrush(QColor("#1E8449")))
                    item.setFont(QFont("Segoe UI", 10, QFont.Bold))
                self._st_table.setItem(r, c, item)

    def _reload_stations(self):
        self._auto_load_stations()
        if self._plan:
            self._calculate_plan()

    # ── station edit / save ────────────────────────────────────────────────
    def _toggle_station_edit(self):
        currently_editing = self._st_table.editTriggers() != QAbstractItemView.NoEditTriggers
        if currently_editing:
            # Turn off edit mode without saving
            self._st_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
            self._btn_edit_st.setText("✏  Edit Stations")
            self._btn_edit_st.setStyleSheet(
                "background:#E67E22; color:white; font-weight:bold; border-radius:4px; padding:0 12px;")
            self._btn_save_st.setEnabled(False)
            self._status.showMessage("Edit cancelled — no changes saved.")
        else:
            # Enable editing on cols 0-3 (not col 4 which is auto-calculated)
            self._st_table.setEditTriggers(QAbstractItemView.DoubleClicked | QAbstractItemView.SelectedClicked)
            for r in range(self._st_table.rowCount()):
                # Lock column 4 (Devices/Day — calculated)
                item = self._st_table.item(r, 4)
                if item:
                    item.setFlags(item.flags() & ~Qt.ItemIsEditable)
                    item.setBackground(QBrush(QColor("#D5D8DC")))
            self._btn_edit_st.setText("✖  Cancel Edit")
            self._btn_edit_st.setStyleSheet(
                "background:#C0392B; color:white; font-weight:bold; border-radius:4px; padding:0 12px;")
            self._btn_save_st.setEnabled(True)
            self._status.showMessage(
                "Edit mode ON — double-click any cell to edit. Click 'Save Stations' when done.")

    def _save_stations(self):
        try:
            rows = []
            for r in range(self._st_table.rowCount()):
                def cell(c): return (self._st_table.item(r, c).text()
                                     if self._st_table.item(r, c) else "")
                rows.append({
                    "Station Name":    cell(0),
                    "Station Qty":     cell(1),
                    "Slots Qty":       cell(2),
                    "Units per hour":  cell(3),
                    "Devices per day": cell(4),
                })
            df = pd.DataFrame(rows)
            df.to_excel(STATION_EXCEL, index=False)

            # Turn off edit mode
            self._st_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
            self._btn_edit_st.setText("✏  Edit Stations")
            self._btn_edit_st.setStyleSheet(
                "background:#E67E22; color:white; font-weight:bold; border-radius:4px; padding:0 12px;")
            self._btn_save_st.setEnabled(False)

            # Reload from file so UPH parsing is fresh
            self._auto_load_stations()
            if self._plan:
                self._calculate_plan()
            QMessageBox.information(self, "Saved",
                                    f"Station config saved to:\n{STATION_EXCEL}")
            self._status.showMessage("Station config saved and reloaded.")
        except Exception as e:
            QMessageBox.critical(self, "Save Error", f"Could not save station config:\n{e}")

    # ── fetch WOs ──────────────────────────────────────────────────────────
    def _fetch_wo(self):
        to_date = self._from_date_edit.date().toString("yyyy-MM-dd")
        self._progress.setRange(0, 100)
        self._progress.setValue(0)
        self._progress.show()
        self._btn_fetch.setEnabled(False)
        self._status.showMessage(
            f"Connecting to database…  (fetching {WO_FETCH_FROM} → {to_date})"
        )

        self._worker = FetchWorker(to_date)
        self._worker.finished.connect(self._on_fetch_done)
        self._worker.error.connect(self._on_fetch_error)
        self._worker.progress.connect(self._progress.setValue)
        self._worker.start()

    def _on_fetch_done(self, rows: list):
        # Preserve any planned-start dates the user already changed in the table
        saved_overrides: dict = {}
        for r in range(self._wo_table.rowCount()):
            wo_item   = self._wo_table.item(r, 0)
            dte_widget = self._wo_table.cellWidget(r, 3)
            if wo_item and dte_widget:
                saved_overrides[wo_item.text()] = dte_widget.dateTime().toPyDateTime()

        self._progress.hide()
        self._btn_fetch.setEnabled(True)
        self._wo_rows = rows
        self._lbl_wo_count.setText(f"WOs: {len(rows)}")
        self._populate_wo_table(rows, saved_overrides)
        self._check_ready()

        if not rows:
            self._status.showMessage("No work orders found for the selected date range.")
            self._tabs.setCurrentIndex(0)
            return

        # If stations are already loaded, auto-calculate and jump straight to the plan
        if self._stations:
            self._status.showMessage(f"Fetched {len(rows)} work orders — calculating plan…")
            self._calculate_plan()   # also switches to Flow Plan tab on success
        else:
            self._tabs.setCurrentIndex(0)
            self._status.showMessage(
                f"Fetched {len(rows)} work orders.  "
                "Load station Excel then click 'Calculate Plan'."
            )

    def _populate_wo_table(self, rows: list, overrides: dict = None):
        if overrides is None:
            overrides = {}
        self._wo_table.setRowCount(0)
        for row in rows:
            r = self._wo_table.rowCount()
            self._wo_table.insertRow(r)

            # Use the user's previously set date if available, else fall back to DB date
            planned_dt = overrides.get(row["wo"]) or self._safe_start_date(row["create_date"])
            cd = row["create_date"]
            if cd is None or (hasattr(pd, "isna") and pd.isna(cd)):
                date_str = "—"
            elif hasattr(cd, "strftime"):
                date_str = cd.strftime("%Y-%m-%d")
            else:
                date_str = str(cd)[:10]

            wo_item  = QTableWidgetItem(row["wo"])
            qty_item = QTableWidgetItem(str(row["qty"]))
            dt_item  = QTableWidgetItem(date_str)

            for item in (wo_item, qty_item, dt_item):
                item.setTextAlignment(Qt.AlignCenter)
                item.setFlags(item.flags() & ~Qt.ItemIsEditable)

            wo_item.setForeground(QBrush(QColor("#C0392B")))
            wo_item.setFont(QFont("Segoe UI", 10, QFont.Bold))
            qty_item.setFont(QFont("Segoe UI", 10, QFont.Bold))
            qty_item.setForeground(QBrush(QColor("#1A5276")))

            self._wo_table.setItem(r, 0, wo_item)
            self._wo_table.setItem(r, 1, qty_item)
            self._wo_table.setItem(r, 2, dt_item)

            # Inline editable QDateTimeEdit for planned start
            dte = QDateTimeEdit()
            dte.setDisplayFormat("yyyy-MM-dd  HH:mm")
            dte.setCalendarPopup(True)
            dte.setDateTime(planned_dt.replace(tzinfo=None))
            dte.setStyleSheet(
                "QDateTimeEdit {"
                "  background:#FDFEFE; border:1px solid #AED6F1;"
                "  border-radius:3px; padding:2px 6px; font-size:11px; font-weight:bold;"
                "  color:#154360;"
                "}"
                "QDateTimeEdit::drop-down { width:20px; }"
            )
            dte.setMinimumHeight(28)
            self._wo_table.setCellWidget(r, 3, dte)

        self._wo_table.resizeRowsToContents()

    def _on_fetch_error(self, msg: str):
        self._progress.hide()
        self._btn_fetch.setEnabled(True)
        QMessageBox.critical(self, "DB Error", f"Failed to fetch WOs:\n{msg}")
        self._status.showMessage("DB fetch failed.")

    # ── calculate ──────────────────────────────────────────────────────────
    def _check_ready(self):
        ready = bool(self._wo_rows) and bool(self._stations)
        self._btn_calc.setEnabled(ready)

    def _safe_start_date(self, cd) -> datetime:
        """Convert any create_date value (Timestamp, date, NaT, str, None) to a datetime at 09:00."""
        try:
            if cd is None or (hasattr(pd, "isna") and pd.isna(cd)):
                return datetime.combine(datetime.today().date(), dtime(PROD_START_HOUR, 0))
        except Exception:
            pass
        try:
            if isinstance(cd, datetime):
                return datetime.combine(cd.date(), dtime(PROD_START_HOUR, 0))
            if hasattr(cd, "date"):          # pandas Timestamp
                return datetime.combine(cd.date(), dtime(PROD_START_HOUR, 0))
            if isinstance(cd, str):
                # Normalise separators: handles YYYY-MM-DD and YYYY/MM/DD
                s = cd[:10].replace("/", "-")
                return datetime.combine(datetime.strptime(s, "%Y-%m-%d").date(),
                                        dtime(PROD_START_HOUR, 0))
            # datetime.date object
            return datetime.combine(cd, dtime(PROD_START_HOUR, 0))
        except Exception:
            # Last resort — log to status bar when possible rather than silently using today
            return datetime.combine(datetime.today().date(), dtime(PROD_START_HOUR, 0))

    def _calculate_plan(self):
        if not self._wo_rows or not self._stations:
            return
        try:
            wo_inputs = []
            for r, wo in enumerate(self._wo_rows):
                # Read the planned start directly from the inline QDateTimeEdit in column 3
                dte_widget = self._wo_table.cellWidget(r, 3)
                if dte_widget is not None:
                    sdt = dte_widget.dateTime().toPyDateTime()
                else:
                    sdt = self._safe_start_date(wo["create_date"])
                wo_inputs.append({"wo": wo["wo"], "qty": wo["qty"], "start_dt": sdt})

            self._plan = build_plan(wo_inputs, self._stations)
            # Initialise empty actual data so model renders before DB fetch
            for wo_data in self._plan:
                wo_data["actual"] = {}

            # Clear stale actual-date overrides so every fresh calculation
            # defaults all Actual Date widgets to the new End Dates.
            self._actual_dates = {}
            self._refresh_plan_table()
            self._btn_export.setEnabled(True)
            self._tabs.setCurrentIndex(2)
            self._status.showMessage(
                f"Plan calculated — {len(self._plan)} WOs × {len(self._stations)} stations."
                "  Fetching actual test data…"
            )
            self._start_actual_fetch()
        except Exception as e:
            import traceback
            QMessageBox.critical(self, "Plan Error",
                                 f"Failed to calculate plan:\n\n{e}\n\n{traceback.format_exc()}")
            self._status.showMessage("Plan calculation failed — see error dialog.")

    def _save_actual_date_widgets(self):
        """Persist current Actual Date widget values before the model is rebuilt."""
        mdl = self._table.model()
        if not self._plan or mdl is None:
            return
        for wi, wo_data in enumerate(self._plan):
            for si, st in enumerate(wo_data["stations"]):
                idx = mdl.index(wi * ROWS_PER_WO + 2, 3 + si)
                widget = self._table.indexWidget(idx)
                if widget is not None:
                    self._actual_dates[(wo_data["wo"], st["name"])] = \
                        widget.dateTime().toPyDateTime()

    def _refresh_plan_table(self):
        model = PlanTableModel(self._plan, self._stations)
        self._table.setModel(model)
        for wi in range(len(self._plan)):
            sr = wi * ROWS_PER_WO
            self._table.setSpan(sr, 0, ROWS_PER_WO, 1)
            self._table.setSpan(sr, 1, ROWS_PER_WO, 1)
        self._table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)

        # Embed an editable QDateTimeEdit in every Actual Date cell (ri == 2)
        for wi, wo_data in enumerate(self._plan):
            for si, st in enumerate(wo_data["stations"]):
                default_dt = self._actual_dates.get(
                    (wo_data["wo"], st["name"]), st["end"]
                )
                dte = QDateTimeEdit()
                dte.setDisplayFormat("MM/dd  HH:mm")
                dte.setCalendarPopup(True)
                dte.setDateTime(default_dt.replace(tzinfo=None))
                dte.setStyleSheet(
                    "QDateTimeEdit {"
                    "  background:#FEF5E7; border:1px solid #F0B27A;"
                    "  border-radius:2px; padding:1px 4px; font-size:9px;"
                    "}"
                    "QDateTimeEdit::drop-down { width:14px; }"
                )
                dte.setMinimumHeight(24)
                idx = model.index(wi * ROWS_PER_WO + 2, 3 + si)
                self._table.setIndexWidget(idx, dte)

        self._table.resizeRowsToContents()

    def _start_actual_fetch(self):
        self._actual_worker = ActualDataWorker(self._plan, self._stations)
        self._actual_worker.finished.connect(self._on_actual_done)
        self._actual_worker.error.connect(self._on_actual_error)
        self._actual_worker.progress.connect(self._on_actual_progress)
        self._actual_worker.start()

    def _on_actual_progress(self, done: int, total: int):
        self._status.showMessage(
            f"Fetching actual test data… {done}/{total} station queries complete.")

    def _on_actual_done(self, result: dict):
        # Save any user edits to Actual Date widgets before the table is rebuilt
        self._save_actual_date_widgets()
        # result: {wo: {station_name: (tested, passed)}}
        for wo_data in self._plan:
            wo_data["actual"] = result.get(wo_data["wo"], {})
        self._refresh_plan_table()
        self._status.showMessage(
            f"Plan ready — {len(self._plan)} WOs × {len(self._stations)} stations."
            "  In/Out and Yield updated from database.")

    def _on_actual_error(self, msg: str):
        self._status.showMessage(f"Could not fetch actual test data: {msg}")

    # ── override dialog ────────────────────────────────────────────────────

    # ── export ─────────────────────────────────────────────────────────────
    def _export_excel(self):
        if not self._plan:
            return
        path, _ = QFileDialog.getSaveFileName(
            self, "Save Plan", "WO_Flow_Plan.xlsx", "Excel Files (*.xlsx)"
        )
        if not path:
            return
        try:
            rows = []
            for wi, wo_data in enumerate(self._plan):
                actual = wo_data.get("actual", {})
                for ri, lbl in enumerate(ROW_LABELS):
                    row = {
                        "WO#":  wo_data["wo"] if ri == 0 else "",
                        "QTY":  wo_data["qty"] if ri == 0 else "",
                        "Info": lbl,
                    }
                    for si, st in enumerate(wo_data["stations"]):
                        _, passed = actual.get(st["name"], (0, 0))
                        # in_qty: WO qty for first station, else previous station's passed
                        if si == 0:
                            in_qty = wo_data["qty"]
                        else:
                            _, in_qty = actual.get(wo_data["stations"][si-1]["name"], (0, 0))
                        if ri == 0:   val = st["start"].replace(tzinfo=None)
                        elif ri == 1: val = st["end"].replace(tzinfo=None)
                        elif ri == 2:
                            # Read datetime from the embedded QDateTimeEdit widget
                            mdl = self._table.model()
                            if mdl:
                                w_idx = mdl.index(wi * ROWS_PER_WO + 2, 3 + si)
                                w = self._table.indexWidget(w_idx)
                                val = w.dateTime().toPyDateTime().replace(tzinfo=None) if w \
                                      else st["end"].replace(tzinfo=None)
                            else:
                                val = st["end"].replace(tzinfo=None)
                        # elif ri == 3: val = st["devices_per_day"]  # Units/Day – reserved
                        elif ri == 3: val = f"{in_qty} / {passed}" if (in_qty or passed) else "— / —"
                        elif ri == 4:
                            val = round(passed / in_qty, 4) if in_qty else None
                        else:         val = ""
                        row[st["name"]] = val
                    rows.append(row)

            df_out = pd.DataFrame(rows)
            with pd.ExcelWriter(path, engine="openpyxl") as writer:
                df_out.to_excel(writer, index=False, sheet_name="WO Flow Plan")
                ws = writer.sheets["WO Flow Plan"]

                from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
                from openpyxl.utils import get_column_letter

                thin   = Side(style="thin", color="BDC3C7")
                border = Border(left=thin, right=thin, top=thin, bottom=thin)
                center = Alignment(horizontal="center", vertical="center", wrap_text=True)
                hdr_fill = PatternFill("solid", fgColor="2C3E50")
                hdr_font = Font(color="FFFFFF", bold=True, size=10)
                row_fills = [
                    PatternFill("solid", fgColor="A8D5A2"),   # Start Date   – green
                    PatternFill("solid", fgColor="F7DC6F"),   # End Date     – yellow
                    PatternFill("solid", fgColor="FAD7A0"),   # Actual Date  – orange
                    # PatternFill("solid", fgColor="AED6F1"), # Units/Day    – reserved
                    PatternFill("solid", fgColor="D6EAF8"),   # WO In/Out    – light blue
                    PatternFill("solid", fgColor="A9DFBF"),   # Yield %      – light green
                ]

                for cell in ws[1]:
                    cell.fill = hdr_fill; cell.font = hdr_font
                    cell.alignment = center; cell.border = border

                for wi in range(len(self._plan)):
                    actual = self._plan[wi].get("actual", {})
                    st_names = [s["name"] for s in self._stations]
                    for ri in range(ROWS_PER_WO):
                        er = wi * ROWS_PER_WO + ri + 2
                        for ci in range(1, len(df_out.columns) + 1):
                            cell = ws.cell(row=er, column=ci)
                            cell.alignment = center; cell.border = border
                            if ci > 2:
                                st_col_idx = ci - 4   # columns: WO#, QTY, Info, then stations
                                if ri == 4 and 0 <= st_col_idx < len(st_names):
                                    t2, p2 = actual.get(st_names[st_col_idx], (0, 0))
                                    if t2 == 0:
                                        cell.fill = PatternFill("solid", fgColor="D5D8DC")
                                    elif p2 / t2 >= 0.95:
                                        cell.fill = PatternFill("solid", fgColor="A9DFBF")
                                    elif p2 / t2 >= 0.80:
                                        cell.fill = PatternFill("solid", fgColor="FAD7A0")
                                    else:
                                        cell.fill = PatternFill("solid", fgColor="F1948A")
                                    cell.font = Font(bold=True, size=10)
                                    cell.number_format = "0.0%"
                                else:
                                    cell.fill = row_fills[ri]
                                # Date / time format for Start Date, End Date, Actual Date
                                if ri in (0, 1, 2) and 0 <= st_col_idx < len(st_names):
                                    cell.number_format = "MM/DD/YY HH:MM"
                    sr = wi * ROWS_PER_WO + 2
                    er = sr + ROWS_PER_WO - 1
                    ws.merge_cells(start_row=sr, start_column=1, end_row=er, end_column=1)
                    ws.merge_cells(start_row=sr, start_column=2, end_row=er, end_column=2)
                    for ci in (1, 2):
                        c = ws.cell(row=sr, column=ci)
                        c.fill = PatternFill("solid", fgColor="FADBD8")
                        c.font = Font(color="C0392B", bold=True, size=10)
                        c.alignment = center

                for col in ws.columns:
                    w = max(len(str(c.value or "")) for c in col)
                    ws.column_dimensions[get_column_letter(col[0].column)].width = max(13, min(w + 4, 30))

            QMessageBox.information(self, "Exported", f"Saved to:\n{path}")
            self._status.showMessage(f"Exported → {os.path.basename(path)}")
        except Exception as e:
            QMessageBox.critical(self, "Export Error", str(e))


# ══════════════════════════════════════════════════════════════════════════════
if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    app.setFont(QFont("Segoe UI", 9))
    app.setApplicationName("WO Flow Planning")
    app.setWindowIcon(make_app_icon())
    win = WOPlanningApp()
    win.show()
    sys.exit(app.exec_())
