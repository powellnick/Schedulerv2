import re, io, math
import pandas as pd
import streamlit as st
from PIL import Image, ImageDraw, ImageFont
from collections import defaultdict

from datetime import datetime
try:
    from zoneinfo import ZoneInfo
except Exception:
    ZoneInfo = None

st.set_page_config(page_title="SMSOLauncher", layout="wide")


# --- Helper to find transporter-id column ---
def find_transporter_col(df: pd.DataFrame):
    """Return the column name that looks like a transporter id column, or None."""
    for col in df.columns:
        key = str(col).strip().lower().replace(" ", "")
        if key in ("transporterid", "transporter_id"):
            return col
    return None

def make_export_xlsx(df, launcher_name: str) -> bytes:
    export_cols = [
        'Order', 'Driver name', "CX #'s", 'Van', 'Staging Location', 'Pad', 'Time'
    ]
    out = df.copy().reset_index(drop=True)
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

def parse_routes(file):
    df = pd.read_excel(file, sheet_name=0)
    df = df[df['Route code'].astype(str).str.startswith('CX', na=False)].copy()
    df['CX'] = df['Route code'].str.extract(r'(CX\d+)')
    if 'Van' not in df.columns:
        df['Van'] = None
    else:
        df['Van'] = df['Van'].astype(str).str.strip()

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


# --- Van memory helpers ---
def parse_van_memory(file):
    """Read van history into {transporter_id: [(van, freq), ...]} and return (memory, df, tid_col)."""
    mem_df = pd.read_excel(file, sheet_name=0)
    tid_col = find_transporter_col(mem_df)
    if tid_col is None:
        return {}, mem_df, None

    # detect Van/Fq columns, e.g. 'Van 1', 'Fq 1'
    van_cols = {}
    fq_cols = {}
    for col in mem_df.columns:
        m_v = re.match(r'\s*van\s*(\d+)', str(col), flags=re.I)
        if m_v:
            van_cols[int(m_v.group(1))] = col
        m_f = re.match(r'\s*fq\s*(\d+)', str(col), flags=re.I)
        if m_f:
            fq_cols[int(m_f.group(1))] = col

    memory = {}
    for _, row in mem_df.iterrows():
        tid = row[tid_col]
        if pd.isna(tid):
            continue
        tid = str(tid).strip()
        prefs = []
        for i in range(1, 6):
            vcol = van_cols.get(i)
            fcol = fq_cols.get(i)
            if not vcol or not fcol:
                continue
            van = row.get(vcol)
            freq = row.get(fcol)
            if pd.isna(van) or pd.isna(freq):
                continue
            van = str(van).strip()
            try:
                freq = int(freq)
            except Exception:
                continue
            if van:
                prefs.append((van, freq))
        if prefs:
            agg = defaultdict(int)
            for v, f in prefs:
                agg[v] += f
            memory[tid] = sorted(agg.items(), key=lambda x: -x[1])[:5]

    return memory, mem_df, tid_col


def update_van_memory(memory, routes_df, transporter_col):
    """Increase frequency for each Van in routes_df, keyed by transporter id."""
    if transporter_col is None or 'Van' not in routes_df.columns:
        return memory

    for _, row in routes_df.iterrows():
        tid = row.get(transporter_col)
        van = row.get('Van')
        if pd.isna(tid) or van is None:
            continue
        tid = str(tid).strip()
        van = str(van).strip()
        if not van or van.lower() == 'nan':
            continue

        if tid not in memory:
            memory[tid] = []
        found = False
        for i, (v, f) in enumerate(memory[tid]):
            if v == van:
                memory[tid][i] = (v, f + 1)
                found = True
                break
        if not found:
            memory[tid].append((van, 1))
        memory[tid] = sorted(memory[tid], key=lambda x: -x[1])[:5]

    return memory


def assign_vans_from_memory(routes_df, transporter_col, memory):
    """Auto-assign vans when the Van column is empty, using van frequency memory.

    If multiple drivers share the same top van, the driver processed first keeps it and
    later drivers fall back to their next most frequent van.
    """
    if transporter_col is None:
        return routes_df

    if 'Van' not in routes_df.columns:
        routes_df['Van'] = pd.NA

    assigned_vans = set()
    for idx, row in routes_df.iterrows():
        existing_van = row.get('Van')
        if isinstance(existing_van, str) and existing_van.strip() and existing_van.lower() != 'nan':
            assigned_vans.add(existing_van.strip())
            continue

        tid = row.get(transporter_col)
        if pd.isna(tid):
            continue
        tid = str(tid).strip()
        prefs = memory.get(tid)
        if not prefs:
            continue

        chosen = None
        for v, _f in prefs:
            if v not in assigned_vans:
                chosen = v
                break
        if chosen is None:
            chosen = prefs[0][0]

        routes_df.at[idx, 'Van'] = chosen
        assigned_vans.add(chosen)

    return routes_df


def van_memory_to_df(memory, base_df, tid_col):
    """Write memory back into the same Van/Fq columns we read from base_df."""
    df = base_df.copy()

    # ensure Van 1..5 / Fq 1..5 columns exist
    for i in range(1, 6):
        vcol = f"Van {i}"
        fcol = f"Fq {i}"
        if vcol not in df.columns:
            df[vcol] = ""
        if fcol not in df.columns:
            df[fcol] = 0

    for tid, prefs in memory.items():
        mask = df[tid_col].astype(str) == str(tid)
        if mask.any():
            idx = df[mask].index[0]
        else:
            idx = len(df)
            df.loc[idx, tid_col] = tid

        for i in range(1, 6):
            vcol = f"Van {i}"
            fcol = f"Fq {i}"
            if i <= len(prefs):
                v, f = prefs[i-1]
                df.loc[idx, vcol] = v
                df.loc[idx, fcol] = f
            else:
                df.loc[idx, vcol] = ""
                df.loc[idx, fcol] = 0

    return df

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

    total_rows = sum(len(g[2]) for g in groups)
    width = left_pad_w + idx_col_w + name_w + cx_w + van_w + 4*pic_w + stg_w + 40
    height = header_h + total_rows*(row_h+gap) + 40

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
    # Current date (PST) centered above the Van Pictures subcolumns
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
        col = pad_colors.get(int(p) if p==p else None, (220,220,220))
        block_h = len(sub)*(row_h+gap)
        d.rectangle([0,y,left_pad_w,y+block_h], fill=col, outline=(0,0,0))
        label = f"Pad {int(p)}\n{t}" if p==p else f"{t}"
        lines = label.split("\n")
        yy = y + block_h/2 - len(lines)*12
        for line in lines:
            w = d.textlength(line, font=font_bold)
            d.text((left_pad_w/2 - w/2, yy), line, fill=(0,0,0), font=font_bold)
            yy += 24

        for _, row in sub.iterrows():
            base_color = pad_colors.get(int(p), (220,220,220))
            row_color = tuple(int(c + (255 - c) * 0.6) for c in base_color)
            
            d.rectangle([left_pad_w, y, left_pad_w+idx_col_w, y+row_h], fill=row_color, outline=(0,0,0))
            w = d.textlength(str(idx), font=font_bold)
            d.text((left_pad_w + (idx_col_w - w)/2, y+8), str(idx), fill=(0,0,0), font=font_bold)            
            
            # name
            d.rectangle([left_pad_w+idx_col_w, y, left_pad_w+idx_col_w+name_w, y+row_h], fill=row_color, outline=(0,0,0))
            d.text((left_pad_w+idx_col_w+8, y+8), str(row['Driver name']), fill=(0,0,0), font=font)

            # CX
            x = left_pad_w+idx_col_w+name_w
            d.rectangle([x, y, x+cx_w, y+row_h], fill=row_color, outline=(0,0,0))
            d.text((x+8, y+8), str(row['CX']).replace('CX',''), fill=(0,0,0), font=font_bold)

            # Van
            x += cx_w
            d.rectangle([x, y, x+van_w, y+row_h], fill=row_color, outline=(0,0,0))
            d.text((x+8, y+8), "" if pd.isna(row['Van']) else str(row['Van']), fill=(0,0,0), font=font_bold)

            # Staging
            x += van_w
            d.rectangle([x, y, x+stg_w, y+row_h], fill=row_color, outline=(0,0,0))
            d.text((x+8, y+8), str(row['Staging Location']), fill=(0,0,0), font=font_bold)

            # Van Pictures
            x = left_pad_w + idx_col_w + name_w + cx_w + van_w + stg_w  # starting x for pictures
            for _ in range(4):
                d.rectangle([x, y, x+pic_w, y+row_h], fill=row_color, outline=(0,0,0))
                # leave empty
                x += pic_w

            y += row_h + gap
            idx += 1

    return img

st.title("SMSO Schedule Builder")
launcher = st.text_input("Launcher name", value="", placeholder="Enter launcher name")

col1, col2, col3 = st.columns(3)
with col1:
    routes_file = st.file_uploader("Upload Routes file (e.g., Routes_DWS4_... .xlsx)", type=["xlsx"], key="routes")
with col2:
    zonemap_file = st.file_uploader("Upload ZoneMap file (.xlsx)", type=["xlsx"], key="zonemap")
with col3:
    van_hist_file = st.file_uploader("Upload Van history (optional)", type=["xlsx"], key="van_hist")


if routes_file and zonemap_file:
    routes = parse_routes(routes_file)
    zonemap = parse_zonemap(zonemap_file)

    # --- Van memory handling ---
    transporter_col = find_transporter_col(routes)

    if 'van_hist_file' in locals() and van_hist_file is not None:
        van_memory, van_hist_df, van_tid_col = parse_van_memory(van_hist_file)
    else:
        van_memory, van_hist_df, van_tid_col = {}, None, None

    has_any_van = (
        'Van' in routes.columns
        and routes['Van'].apply(lambda v: isinstance(v, str) and v.strip() != '' and str(v).lower() != 'nan').any()
    )

    if transporter_col is not None:
        if has_any_van:
            # learn from explicit Vans in routes
            van_memory = update_van_memory(van_memory, routes, transporter_col)
            # also fill in any missing Vans using memory
            routes = assign_vans_from_memory(routes, transporter_col, van_memory)
        else:
            # no vans in routes => fully auto-assign
            routes = assign_vans_from_memory(routes, transporter_col, van_memory)
    # --- end van memory handling ---

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

    # Optional: export updated van history if a history file was uploaded
    if van_hist_df is not None and van_tid_col is not None:
        updated_hist_df = van_memory_to_df(van_memory, van_hist_df, van_tid_col)
        hist_buf = io.BytesIO()
        with pd.ExcelWriter(hist_buf, engine='openpyxl') as writer:
            updated_hist_df.to_excel(writer, index=False, sheet_name='VanHistory')
        hist_buf.seek(0)
        st.download_button(
            "Download updated Van history",
            data=hist_buf.getvalue(),
            file_name="van_history_updated.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    #buf = io.BytesIO()
    #img.save(buf, format="PNG")
    #st.download_button("Download PNG", data=buf.getvalue(), file_name="schedule.png", mime="image/png")
#else:
    #st.info("Upload both the **Routes** and **ZoneMap** files to generate the schedule.")
