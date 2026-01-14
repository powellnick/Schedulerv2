import re, io, math
import pandas as pd
import streamlit as st
from PIL import Image, ImageDraw, ImageFont
from collections import defaultdict

import gspread
from google.oauth2.service_account import Credentials

from datetime import datetime
try:
    from zoneinfo import ZoneInfo
except Exception:
    ZoneInfo = None

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]
VAN_HISTORY_SHEET_NAME = "SMSO_VanHistory"

def clean_van_value(v):
    if pd.isna(v):
        return None
    s = str(v).strip()
    m = re.fullmatch(r"(\d+)\.0", s)
    if m:
        s = m.group(1)
    return s or None

def reset_van_history_sheet():
    client = get_gs_client()
    if client is None:
        return
    try:
        sh = client.open(VAN_HISTORY_SHEET_NAME)
        ws = sh.sheet1
    except Exception as e:
        st.error(f"Error opening sheet '{VAN_HISTORY_SHEET_NAME}' to reset: {e}")
        return

    headers = [
        'Transporter Id', 'Driver name',
        'Van 1', 'Fq 1',
        'Van 2', 'Fq 2',
        'Van 3', 'Fq 3',
        'Van 4', 'Fq 4',
        'Van 5', 'Fq 5',
    ]
    try:
        ws.clear()
        ws.update('A1', [headers])
        st.success("Van history Google Sheet cleared")
    except Exception as e:
        st.error(f"Error clearing van history sheet: {e}")

def get_gs_client():
    try:
        info = st.secrets["gcp_service_account"]
    except Exception:
        st.error("No [gcp_service_account] block found in Streamlit secrets.")
        return None
    try:
        creds = Credentials.from_service_account_info(info, scopes=SCOPES)
        return gspread.authorize(creds)
    except Exception as e:
        st.error(f"Error creating Google Sheets client: {e}")
        return None

def load_van_memory_from_sheet():
    client = get_gs_client()
    if client is None:
        return {}

    try:
        sh = client.open(VAN_HISTORY_SHEET_NAME)
        ws = sh.sheet1
        data = ws.get_all_records()
    except Exception:
        return {}

    if not data:
        return {}

    df_hist = pd.DataFrame(data)
    if "Transporter Id" not in df_hist.columns:
        return {}

    memory = {}
    for _, row in df_hist.iterrows():
        tid = row.get("Transporter Id")
        if pd.isna(tid):
            continue
        tid = str(tid).strip()
        vans = {}
        for i in range(1, 6):
            vcol = f"Van {i}"
            fcol = f"Fq {i}"
            van = row.get(vcol)
            freq = row.get(fcol)
            if pd.isna(van):
                continue
            v_clean = clean_van_value(van)
            if not v_clean:
                continue
            try:
                freq = int(freq)
            except Exception:
                continue
            vans[v_clean] = vans.get(v_clean, 0) + freq
        if vans:
            memory[tid] = vans
    return memory

def save_van_memory_to_sheet(memory, routes_df, transporter_col):
    client = get_gs_client()
    if client is None:
        st.warning("Van history not saved: no Google Sheets client.")
        return

    try:
        sh = client.open(VAN_HISTORY_SHEET_NAME)
        ws = sh.sheet1
    except Exception as e:
        st.error(f"Error opening sheet '{VAN_HISTORY_SHEET_NAME}': {e}")
        return

    df_hist = van_memory_to_df(memory, routes_df, transporter_col)
    ws.clear()
    if df_hist.empty:
        st.warning("Van history DataFrame is empty; nothing to write.")
        return

    rows = [df_hist.columns.tolist()] + df_hist.astype(str).values.tolist()
    try:
        ws.update(rows)
        st.success(f"Van history saved to Google Sheet ({len(df_hist)} rows).")
    except Exception as e:
        st.error(f"Error updating sheet: {e}")

st.set_page_config(page_title="SMSOLauncher", layout="wide")

def make_export_xlsx(df, launcher_name: str) -> bytes:
    export_cols = [
        'Order', 'Driver name', "CX #'s", 'Van', 'Staging Location', 'Pad', 'Time'
    ]
    out = df.copy().reset_index(drop=True)
    if 'Order' in out.columns:
        out['Order'] = range(1, len(out) + 1)
    else:
        out.insert(0, 'Order', out.index + 1)
    out.rename(columns={'CX': "CX #'s"}, inplace=True)

    for col in ['Front', 'Back', 'D Side', 'P Side']:
        if col not in out.columns:
            out[col] = ''
    out = out[['Order','Pad','Time','Driver name',"CX #'s",'Van','Staging Location','Front','Back','D Side','P Side']]

    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        out.to_excel(writer, index=False, sheet_name='Schedule')
        ws = writer.sheets['Schedule']

        ws.insert_rows(1)
        ws.merge_cells('H1:K1')
        ws['H1'] = 'Van Pictures'

        ws.freeze_panes = "A3"

        widths = {
            'A': 6,   # Order
            'B': 6,   # Pad
            'C': 8,   # Time
            'D': 32,  # Driver name
            'E': 8,   # CX #'s
            'F': 10,  # Van
            'G': 18,  # Staging Location
            'H': 10,  # Front (Van Pictures)
            'I': 10,  # Back (Van Pictures)
            'J': 10,  # D Side (Van Pictures)
            'K': 10,  # P Side (Van Pictures)
        }
        for col, w in widths.items():
            ws.column_dimensions[col].width = w

    buffer.seek(0)
    return buffer.getvalue()


def load_edited_schedule(file):
    """Load an edited Schedule export (from Download Excel) and normalize columns.

    The exported schedule includes an extra top row for the merged "Van Pictures" header,
    so we attempt reading with header=0 first, then fall back to header=1.
    """
    def _try_read(header_row: int):
        return pd.read_excel(file, sheet_name=0, header=header_row)

    try:
        df = _try_read(0)
    except Exception as e:
        st.error(f"Could not read edited schedule file: {e}")
        return None

    expected = {
        'Pad', 'Time', 'Driver name', "CX #'s", 'Van', 'Staging Location',
        'Front', 'Back', 'D Side', 'P Side'
    }

    if not expected.issubset(set(map(str, df.columns))):
        try:
            df = _try_read(1)
        except Exception as e:
            st.error(f"Could not read edited schedule file with header=1: {e}")
            return None

    df.columns = [str(c).strip() for c in df.columns]

    missing = expected - set(df.columns)
    if missing:
        st.error(f"Edited schedule is missing columns: {', '.join(sorted(missing))}")
        return None

    df = df.rename(columns={"CX #'s": "CX"})

    if 'Van' in df.columns:
        df['Van'] = df['Van'].apply(clean_van_value)

    for col in ['Pad', 'Time', 'Driver name', 'CX', 'Van', 'Staging Location']:
        if col not in df.columns:
            df[col] = None

    return df

def extract_time_range_start(s):
    m = re.search(r'(\d{1,2}:\d{2})', s or '')
    return m.group(1) if m else None

def extract_time(s):
    if not isinstance(s,str): return None
    m = re.search(r'(\d{1,2}:\d{2}\s*[ap]m)', s, flags=re.I)
    if m: return m.group(1).lower()
    m = re.search(r'(\d{1,2}:\d{2})\s*(AM|PM)', s, flags=re.I)
    if m: return (m.group(1)+m.group(2).lower())
    return None

def extract_pad(s):
    if not isinstance(s,str): return None
    m = re.search(r'Pad\s*([1-3])', s, flags=re.I)
    if m: return int(m.group(1))
    m = re.search(r'\bP\s*([1-3])\b', s, flags=re.I)
    if m: return int(m.group(1))
    return None

def find_transporter_col(df: pd.DataFrame):
    """Return the column name that looks like a transporter id column, or None."""
    for col in df.columns:
        key = str(col).strip().lower().replace(" ", "")
        if key in ("transporterid", "transporter_id"):
            return col
    return None

def read_van_list_file(uploaded_file):
    """Read a DownVans/AvailableVans-style xlsx and return a set of cleaned van strings."""
    if uploaded_file is None:
        return set()
    try:
        df = pd.read_excel(uploaded_file)
    except Exception as e:
        st.error(f"Could not read vans file: {e}")
        return set()

    van_col = None
    for col in df.columns:
        key = str(col).strip().lower()
        if key in ("van", "van#", "van #", "vans"):
            van_col = col
            break
    if van_col is None and len(df.columns) > 0:
        van_col = df.columns[0]

    vans = set()
    if van_col is not None:
        for v in df[van_col]:
            v_clean = clean_van_value(v)
            if v_clean:
                vans.add(v_clean)
    return vans

def parse_routes(file):
    df = pd.read_excel(file, sheet_name=0)
    df = df[df['Route code'].astype(str).str.startswith('CX', na=False)].copy()
    df['CX'] = df['Route code'].str.extract(r'(CX\d+)')
    if 'Van' not in df.columns:
        df['Van'] = None
    else:
        df['Van'] = df['Van'].apply(clean_van_value)

    df['Driver name'] = df['Driver name'].astype(str).str.split(r'\s*\|\s*')
    df = df.explode('Driver name').reset_index(drop=True)
    df = df[df['Driver name'].str.len() > 0]

    cols = ['CX', 'Driver name', 'Van']
    tcol = find_transporter_col(df)
    if tcol and tcol not in cols:
        cols.append(tcol)
    return df[cols]

def parse_zonemap(file):
    z = pd.read_excel(file, sheet_name=0, header=None)
    rows, cols = z.shape
    time_above = [None] * rows
    pad_above  = [None] * rows

    def _coerce_time_text(x):
        if isinstance(x, str):
            return x
        if isinstance(x, (pd.Timestamp, )):
            return x.strftime('%I:%M %p')
        return str(x)

    last_time = None
    last_pad  = None
    for rr in range(rows):
        for cc in range(cols):
            cell = z.iat[rr, cc]
            txt = _coerce_time_text(cell)
            t = extract_time(txt) or extract_time_range_start(txt)
            if t:
                last_time = t
            p = extract_pad(txt)
            if p is not None:
                last_pad = p
        time_above[rr] = last_time
        pad_above[rr]  = last_pad

    out = []
    for r in range(rows):
        for c in range(cols):
            val = z.iat[r,c]
            if isinstance(val,str) and 'SMSO' in val and 'CX' in val:
                cx_m = re.search(r'(CX\d+)', val)
                if not cx_m: 
                    continue
                cx = cx_m.group(1)
                staging = None
                stg_col = None
                for dc in [1,2,-1]:
                    cc = c+dc
                    if 0 <= cc < cols:
                        sval = z.iat[r,cc]
                        if isinstance(sval,str) and sval.startswith('STG'):
                            staging = sval.replace("STG.", "") if isinstance(sval, str) else sval
                            stg_col = cc
                            break
                pad=None; time=None

                if stg_col is None:
                    for dc in [2,3,-2]:
                        cc2 = c+dc
                        if 0 <= cc2 < cols:
                            sval = z.iat[r, cc2]
                            if isinstance(sval, str) and sval.startswith('STG'):
                                staging = sval.replace('STG.', '') if isinstance(sval, str) else sval
                                stg_col = cc2
                                break

                if stg_col is not None:
                    for rr in range(r-1, max(-1, r-30), -1):
                        neighbor_cols = [stg_col, stg_col+1, stg_col-1, stg_col+2]
                        for cc2 in neighbor_cols:
                            if 0 <= cc2 < cols:
                                raw = z.iat[rr, cc2]
                                txt = _coerce_time_text(raw)
                                t = extract_time(txt) or extract_time_range_start(txt)
                                if t:
                                    time = t
                                    break
                        if time:
                            break

                for rr in range(r-1, max(-1, r-30), -1):
                    if pad is None:
                        txt1 = z.iat[rr, c]
                        if isinstance(txt1, str):
                            p = extract_pad(txt1)
                            if p is not None:
                                pad = p
                                break
                    if pad is None and stg_col is not None:
                        txt2 = z.iat[rr, stg_col]
                        if isinstance(txt2, str):
                            p = extract_pad(txt2)
                            if p is not None:
                                pad = p
                                break

                if pad is None:
                    pad = pad_above[r]
                if time is None:
                    time = time_above[r]
                out.append({'CX':cx, 'Pad':pad, 'Time':time, 'Staging Location': staging})
    return pd.DataFrame(out).drop_duplicates(subset=['CX'])

def update_van_memory(memory, routes_df, transporter_col):
    if transporter_col is None or 'Van' not in routes_df.columns:
        return memory or {}

    if memory is None:
        memory = {}

    for _, row in routes_df.iterrows():
        tid = row.get(transporter_col)
        van = row.get('Van')
        if pd.isna(tid) or van is None:
            continue
        tid = str(tid).strip()
        van_clean = clean_van_value(van)
        if not van_clean:
            continue

        if tid not in memory:
            memory[tid] = {}
        memory[tid][van_clean] = memory[tid].get(van_clean, 0) + 1

    return memory


def assign_vans_from_memory(routes_df, transporter_col, memory, down_vans=None, available_vans=None):
    if transporter_col is None or not memory:
        return routes_df

    if down_vans is None:
        down_vans = set()
    if available_vans is None:
        available_vans = set()

    if 'Van' not in routes_df.columns:
        routes_df['Van'] = pd.NA

    assigned_vans = set()

    for i, row in routes_df.iterrows():
        v = row.get('Van')
        if v is None or (isinstance(v, float) and pd.isna(v)):
            continue

        v_str = clean_van_value(v)
        if not v_str:
            continue

        if v_str in down_vans:
            routes_df.at[i, 'Van'] = pd.NA
            continue

        if available_vans and v_str not in available_vans:
            routes_df.at[i, 'Van'] = pd.NA
            continue

        assigned_vans.add(v_str)

    for idx, row in routes_df.iterrows():
        existing_van = row.get('Van')
        if existing_van is not None and not (isinstance(existing_van, float) and pd.isna(existing_van)):
            v_str = clean_van_value(existing_van)
            if v_str:
                assigned_vans.add(v_str)
                continue

        tid = row.get(transporter_col)
        if pd.isna(tid):
            continue
        tid = str(tid).strip()
        prefs_dict = memory.get(tid)
        if not prefs_dict:
            continue

        prefs = sorted(prefs_dict.items(), key=lambda x: -x[1])[:5]

        chosen = None
        for v, _f in prefs:
            if v in down_vans:
                continue
            if available_vans and v not in available_vans:
                continue
            if v not in assigned_vans:
                chosen = v
                break
        if chosen is None:
            continue

        routes_df.at[idx, 'Van'] = chosen
        assigned_vans.add(chosen)

    return routes_df


def van_memory_to_df(memory, routes_df, transporter_col):
    rows = []
    if not memory or transporter_col is None:
        return pd.DataFrame(columns=[
            'Transporter Id', 'Driver name',
            'Van 1', 'Fq 1', 'Van 2', 'Fq 2', 'Van 3', 'Fq 3', 'Van 4', 'Fq 4', 'Van 5', 'Fq 5'
        ])

    name_map = {}
    if transporter_col in routes_df.columns and 'Driver name' in routes_df.columns:
        for _, row in routes_df.iterrows():
            tid = row.get(transporter_col)
            if pd.isna(tid):
                continue
            tid = str(tid).strip()
            if tid not in name_map:
                name_map[tid] = str(row['Driver name'])

    for tid, vans in memory.items():
        prefs = sorted(vans.items(), key=lambda x: -x[1])[:5]
        entry = {
            'Transporter Id': tid,
            'Driver name': name_map.get(tid, ''),
        }
        for i in range(5):
            vcol = f'Van {i+1}'
            fcol = f'Fq {i+1}'
            if i < len(prefs):
                v, f = prefs[i]
                entry[vcol] = v
                entry[fcol] = f
            else:
                entry[vcol] = ''
                entry[fcol] = 0
        rows.append(entry)

    return pd.DataFrame(rows)

def time_to_minutes(t):
    if not isinstance(t,str): return 1_000_000
    m = re.match(r'(\d{1,2}):(\d{2})', t)
    return int(m.group(1))*60 + int(m.group(2)) if m else 1_000_000

def shift_time_str(t, delta_min=5):
    if not isinstance(t, str):
        return t
    m = re.match(r"\s*(\d{1,2}):(\d{2})", t)
    if not m:
        return t
    h = int(m.group(1))
    mi = int(m.group(2))
    total = h*60 + mi + int(delta_min)
    new_h = (total // 60) % 24
    new_m = total % 60
    return f"{new_h}:{new_m:02d}"

def render_schedule(df, launcher=""):
    left_pad_w = 200
    idx_col_w = 60
    name_w = 360
    cx_w = 90
    van_w = 80
    pic_w = 70
    stg_w = 140
    header_h = 120
    row_h = 38
    gap = 2

    groups = []
    for (t,p), sub in df.groupby(['Time','Pad']):
        groups.append((t, p, sub.sort_values('Driver name')))
    groups.sort(key=lambda x: (time_to_minutes(x[0]), (x[1] if x[1] is not None else 9)))

    single_row_h = int(row_h * 1.6)
    total_rows = sum(len(g[2]) for g in groups)
    singletons = sum(1 for _, _, sub in groups if len(sub) == 1)
    extra_height = singletons * (single_row_h - row_h)

    width = left_pad_w + idx_col_w + name_w + cx_w + van_w + 4*pic_w + stg_w + 40
    height = header_h + total_rows*(row_h+gap) + extra_height + 40

    img = Image.new("RGB", (width, height), (245,245,245))
    d = ImageDraw.Draw(img)
    try:
        font_title = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf", 30)
        font = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf", 20)
        font_bold = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf", 20)
        font_small_bold = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf", 16)
        font_date = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf", 20)
    except:
        font_title=font=font_bold=font_small_bold=None

    d.rectangle([0,0,width,header_h], fill=(255,255,255), outline=(0,0,0))
    d.text((16,16),"Launcher:", fill=(0,0,0), font=font_title)
    d.text((16,55), launcher or "", fill=(0,0,0), font=font_title)
    d.rectangle([left_pad_w,0,width, header_h], outline=(0,0,0))
    d.text((left_pad_w+idx_col_w+10, 16), "DRIVER NAME", fill=(0,0,0), font=font_title)

    x0 = left_pad_w+idx_col_w+name_w
    if ZoneInfo is not None:
        _tz = ZoneInfo("America/Los_Angeles")
        date_str = datetime.now(_tz).strftime("%m/%d/%Y")
    else:
        date_str = datetime.now().strftime("%m/%d/%Y")

    pics_x_start = x0 + cx_w + van_w + stg_w
    pics_x_end   = pics_x_start + 4*pic_w
    date_w = d.textlength(date_str, font=font_small_bold)
    d.text(((pics_x_start + pics_x_end)/2 - date_w/2, 20), date_str, fill=(0,0,0), font=font_date)

    def cell(x1,w,label):
        d.rectangle([x1, 62, x1+w, header_h-8], fill=(255,235,150), outline=(0,0,0))
        d.text((x1+10, 70), label, fill=(0,0,0), font=font_small_bold)
    cell(x0, cx_w, "CX #’s")
    cell(x0+cx_w, van_w, "Van")
    cell(x0+cx_w+van_w, stg_w, "Staging\nLocation")
    cell(x0+cx_w+van_w+stg_w, pic_w, "Front")
    cell(x0+cx_w+van_w+stg_w+pic_w, pic_w, "Back")
    cell(x0+cx_w+van_w+stg_w+2*pic_w, pic_w, "D Side")
    cell(x0+cx_w+van_w+stg_w+3*pic_w, pic_w, "P Side")

    pad_colors = {1:(74,120,206), 2:(226,40,216), 3:(73,230,54)}

    y = header_h
    idx = 1
    for (t, p, sub) in groups:
        cur_row_h = single_row_h if len(sub) == 1 else row_h

        if len(sub) == 1:
            block_h = cur_row_h
        else:
            block_h = len(sub) * (cur_row_h + gap)

        col = pad_colors.get(int(p) if p==p else None, (220,220,220))
        d.rectangle([0, y, left_pad_w, y+block_h], fill=col, outline=(0,0,0))

        label = f"Pad {int(p)}\n{t}" if p==p else f"{t}"
        lines = label.split("\n")
        yy = y + block_h/2 - len(lines)*12
        for line in lines:
            w = d.textlength(line, font=font_bold)
            d.text((left_pad_w/2 - w/2, yy), line, fill=(0,0,0), font=font_bold)
            yy += 24

        row_y = y
        for _, row in sub.iterrows():
            base_color = pad_colors.get(int(p), (220,220,220))
            row_color = tuple(int(c + (255 - c) * 0.6) for c in base_color)
            
            d.rectangle([left_pad_w, row_y, left_pad_w+idx_col_w, row_y+cur_row_h], fill=row_color, outline=(0,0,0))
            w = d.textlength(str(idx), font=font_bold)
            d.text((left_pad_w + (idx_col_w - w)/2, row_y+8), str(idx), fill=(0,0,0), font=font_bold)            
            
            # name
            d.rectangle([left_pad_w+idx_col_w, row_y, left_pad_w+idx_col_w+name_w, row_y+cur_row_h], fill=row_color, outline=(0,0,0))
            d.text((left_pad_w+idx_col_w+8, row_y+8), str(row['Driver name']), fill=(0,0,0), font=font)

            # CX
            x = left_pad_w+idx_col_w+name_w
            d.rectangle([x, row_y, x+cx_w, row_y+cur_row_h], fill=row_color, outline=(0,0,0))
            d.text((x+8, row_y+8), str(row['CX']).replace('CX',''), fill=(0,0,0), font=font_bold)

            # Van
            x += cx_w
            d.rectangle([x, row_y, x+van_w, row_y+cur_row_h], fill=row_color, outline=(0,0,0))
            d.text((x+8, row_y+8), "" if pd.isna(row['Van']) else str(row['Van']), fill=(0,0,0), font=font_bold)

            # Staging
            x += van_w
            d.rectangle([x, row_y, x+stg_w, row_y+cur_row_h], fill=row_color, outline=(0,0,0))
            d.text((x+8, row_y+8), str(row['Staging Location']), fill=(0,0,0), font=font_bold)

            # Van Pictures
            x = left_pad_w + idx_col_w + name_w + cx_w + van_w + stg_w  # starting x for pictures
            for _ in range(4):
                d.rectangle([x, row_y, x+pic_w, row_y+cur_row_h], fill=row_color, outline=(0,0,0))
                x += pic_w

            row_y += cur_row_h + gap
            idx += 1

        y = row_y

    return img


st.title("SMSO Schedule Builder")
launcher = st.text_input("Launcher name", value="", placeholder="Enter launcher name")

col1, col2 = st.columns(2)
with col1:
    routes_file = st.file_uploader("Upload Routes file (e.g., Routes_DWS4_... .xlsx)", type=["xlsx"], key="routes")
with col2:
    zonemap_file = st.file_uploader("Upload ZoneMap file (.xlsx)", type=["xlsx"], key="zonemap")

col3, col4 = st.columns(2)
with col3:    
    downvans_file = st.file_uploader("Upload Down Vans file (optional, .xlsx)", type=["xlsx"], key="downvans")
with col4:
    availablevans_file = st.file_uploader("Upload Vans Available file (optional, .xlsx)", type=["xlsx"], key="availablevans")

edited_schedule_file = st.file_uploader(
    "Re-upload edited Schedule (from Download Excel)",
    type=["xlsx"],
    key="edited_schedule"
)

with st.expander("Clear van history cache"):
    clear_pw = st.text_input("Enter password to clear van history", type="password", key="clear_van_pw")
    if st.button("Clear van history", key="clear_van_btn"):
        if clear_pw == "SMSOclear":
            st.session_state['van_memory'] = {}
            reset_van_history_sheet()
        else:
            st.error("Incorrect password.")


if edited_schedule_file is not None:
    df_display = load_edited_schedule(edited_schedule_file)
    if df_display is None:
        st.stop()

    down_vans_set = read_van_list_file(downvans_file)
    available_vans_set = read_van_list_file(availablevans_file)

    if 'Van' in df_display.columns:
        def _van_allowed(v):
            v = clean_van_value(v)
            if not v:
                return None
            if v in down_vans_set:
                return None
            if available_vans_set and v not in available_vans_set:
                return None
            return v
        df_display['Van'] = df_display['Van'].apply(_van_allowed)

    img = render_schedule(df_display, launcher=launcher)
    st.image(img, caption="Final Schedule")

    png_buf = io.BytesIO()
    img.save(png_buf, format="PNG")
    png_buf.seek(0)
    st.download_button(
        "Download PNG",
        data=png_buf.getvalue(),
        file_name="schedule.png",
        mime="image/png",
    )

    xlsx_bytes = make_export_xlsx(df_display, launcher_name=launcher or "")
    st.download_button(
        "Download Excel",
        data=xlsx_bytes,
        file_name="schedule.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

elif routes_file and zonemap_file:
    routes = parse_routes(routes_file)
    zonemap = parse_zonemap(zonemap_file)

    down_vans_set = read_van_list_file(downvans_file)
    available_vans_set = read_van_list_file(availablevans_file)

    if 'van_memory' not in st.session_state:
        st.session_state['van_memory'] = load_van_memory_from_sheet()
    van_memory = st.session_state['van_memory']

    transporter_col = find_transporter_col(routes)

    has_any_van = 'Van' in routes.columns and routes['Van'].notna().any()

    if transporter_col is not None:
        if has_any_van:
            van_memory = update_van_memory(van_memory, routes, transporter_col)
            routes = assign_vans_from_memory(routes, transporter_col, van_memory, down_vans_set, available_vans_set)
            st.session_state['van_memory'] = van_memory
            save_van_memory_to_sheet(van_memory, routes, transporter_col)
        else:
            routes = assign_vans_from_memory(routes, transporter_col, van_memory, down_vans_set, available_vans_set)

    df = routes.merge(zonemap, on='CX', how='left')

    df['Time'] = df['Time'].apply(lambda t: shift_time_str(t, -5))

    df['Pad'] = df['Pad'].fillna(9)

    df['__t'] = df['Time'].apply(lambda t: (int(t.split(':')[0])*60 + int(t.split(':')[1][:2])) if isinstance(t,str) and ':' in t else 999_999)
    df.sort_values(['__t','Pad','Driver name'], inplace=True)
    df.drop(columns='__t', inplace=True)

    df_display = df.copy()
    df_display['Time'] = df_display['Time'].fillna('—')

    img = render_schedule(df_display, launcher=launcher)
    st.image(img, caption="Final Schedule")

    png_buf = io.BytesIO()
    img.save(png_buf, format="PNG")
    png_buf.seek(0)
    st.download_button(
        "Download PNG",
        data=png_buf.getvalue(),
        file_name="schedule.png",
        mime="image/png",
    )

    xlsx_bytes = make_export_xlsx(df_display, launcher_name=launcher or "")
    st.download_button(
        "Download Excel",
        data=xlsx_bytes,
        file_name="schedule.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    if 'van_memory' in st.session_state and st.session_state['van_memory']:
        transporter_col = find_transporter_col(routes)
        if transporter_col is not None:
            van_hist_df = van_memory_to_df(st.session_state['van_memory'], routes, transporter_col)
            hist_buf = io.BytesIO()
            with pd.ExcelWriter(hist_buf, engine='openpyxl') as writer:
                van_hist_df.to_excel(writer, index=False, sheet_name='VanHistory')
            hist_buf.seek(0)
            st.download_button(
                "Download Van history",
                data=hist_buf.getvalue(),
                file_name="van_history.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
