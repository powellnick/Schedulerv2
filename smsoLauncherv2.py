import re, io, math
import pandas as pd
import streamlit as st
from PIL import Image, ImageDraw, ImageFont

st.set_page_config(page_title="SMSOLauncher", layout="wide")

def make_export_xlsx(df, launcher_name: str) -> bytes:
    export_cols = [
        'Order', 'Driver name', "CX #'s", 'Van', 'Staging Location', 'Pad', 'Time'
    ]
    out = df.copy().reset_index(drop=True)
    out.insert(0, 'Order', out.index + 1)
    out.rename(columns={'CX': "CX #'s"}, inplace=True)

    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        out.to_excel(writer, index=False, sheet_name='Schedule')
        ws = writer.sheets['Schedule']

        ws.freeze_panes = "A2"

        widths = {
            'A': 6,   # Order
            'B': 32,  # Driver name
            'C': 8,   # CX #'s
            'D': 10,  # Van
            'E': 18,  # Staging Location
            'F': 6,   # Pad
            'G': 8,   # Time
        }
        for col, w in widths.items():
            ws.column_dimensions[col].width = w

        # meta = writer.book.create_sheet("Meta")
        # meta['A1'] = "Launcher"; meta['B1'] = launcher_name

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

    return df[['CX','Driver name','Van']]

def parse_zonemap(file):
    z = pd.read_excel(file, sheet_name=0, header=None)
    out = []
    rows, cols = z.shape
    for r in range(rows):
        for c in range(cols):
            val = z.iat[r,c]
            if isinstance(val,str) and 'SMSO' in val and 'CX' in val:
                cx_m = re.search(r'(CX\d+)', val)
                if not cx_m: 
                    continue
                cx = cx_m.group(1)
                # staging near the cell
                staging = None
                for dc in [1,2,-1]:
                    cc = c+dc
                    if 0 <= cc < cols:
                        sval = z.iat[r,cc]
                        if isinstance(sval,str) and sval.startswith('STG'):
                            staging = sval.replace("STG.", "") if isinstance(sval, str) else sval; break
                pad=None; time=None
                for dr in range(1,10):
                    rr = r-dr
                    if rr < 0: break
                    text = z.iat[rr,c]
                    if isinstance(text,str):
                        if pad is None:
                            p = extract_pad(text)
                            if p is not None: pad = p
                        if time is None:
                            t = extract_time(text) or extract_time_range_start(text)
                            if t: time = t
                    if pad and time: break
                out.append({'CX':cx, 'Pad':pad, 'Time':time, 'Staging Location': staging})
    return pd.DataFrame(out).drop_duplicates(subset=['CX'])

def time_to_minutes(t):
    if not isinstance(t,str): return 1_000_000
    m = re.match(r'(\d{1,2}):(\d{2})', t)
    return int(m.group(1))*60 + int(m.group(2)) if m else 1_000_000

def render_schedule(df, launcher=""):
    # Layout
    left_pad_w = 200
    idx_col_w = 60   # more space for index numbers
    name_w = 600
    cx_w = 90
    van_w = 80
    stg_w = 140
    header_h = 110
    row_h = 38
    gap = 2

    groups = []
    for (t,p), sub in df.groupby(['Time','Pad']):
        groups.append((t, p, sub.sort_values('Driver name')))
    groups.sort(key=lambda x: (time_to_minutes(x[0]), (x[1] if x[1] is not None else 9)))

    total_rows = sum(len(g[2]) for g in groups)
    width = left_pad_w + idx_col_w + name_w + cx_w + van_w + stg_w + 40
    height = header_h + total_rows*(row_h+gap) + 40

    img = Image.new("RGB", (width, height), (245,245,245))
    d = ImageDraw.Draw(img)
    try:
        font_title = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf", 30)
        font = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf", 20)
        font_bold = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf", 20)
        font_small_bold = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf", 16)
    except:
        font_title=font=font_bold=font_small_bold=None

    # Header
    d.rectangle([0,0,width,header_h], fill=(255,255,255), outline=(0,0,0))
    d.text((16,16),"Launcher:", fill=(0,0,0), font=font_title)
    d.text((16,55), launcher or "", fill=(0,0,0), font=font_title)
    d.rectangle([left_pad_w,0,width, header_h], outline=(0,0,0))
    d.text((left_pad_w+idx_col_w+10, 16), "DRIVER NAME", fill=(0,0,0), font=font_title)

    # mini header row for the three added columns
    x0 = left_pad_w+idx_col_w+name_w
    def cell(x1,w,label):
        d.rectangle([x1, 62, x1+w, header_h-8], fill=(255,235,150), outline=(0,0,0))
        d.text((x1+10, 70), label, fill=(0,0,0), font=font_small_bold)
    cell(x0, cx_w, "CX #’s")
    cell(x0+cx_w, van_w, "Van")
    cell(x0+cx_w+van_w, stg_w, "Staging\nLocation")

    pad_colors = {1:(226,40,216), 2:(74,120,206), 3:(73,230,54)}

    y = header_h
    idx = 1
    for (t, p, sub) in groups:
        col = pad_colors.get(int(p) if p==p else None, (220,220,220))
        block_h = len(sub)*(row_h+gap)
        # left colored block
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
            row_color = tuple(min(255, int(c*1.3)) for c in base_color)
            
            # number cell (centered)
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

            y += row_h + gap
            idx += 1

    return img

st.title("SMSO Schedule Builder")
launcher = st.text_input("Launcher name", value="", placeholder="Enter launcher name")

col1, col2 = st.columns(2)
with col1:
    routes_file = st.file_uploader("Upload Routes file (e.g., Routes_DWS4_... .xlsx)", type=["xlsx"], key="routes")
with col2:
    zonemap_file = st.file_uploader("Upload ZoneMap file (.xlsx)", type=["xlsx"], key="zonemap")

if routes_file and zonemap_file:
    routes = parse_routes(routes_file)
    zonemap = parse_zonemap(zonemap_file)

    df = routes.merge(zonemap, on='CX', how='inner')
    df = df.dropna(subset=['Pad','Time']).copy()

    # sorting by time then pad
    df['__t'] = df['Time'].apply(lambda t: int(t.split(':')[0])*60 + int(t.split(':')[1]) if isinstance(t,str) and ':' in t else 999_999)
    df.sort_values(['__t','Pad','Driver name'], inplace=True)
    df.drop(columns='__t', inplace=True)

    img = render_schedule(df, launcher=launcher)
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

    xlsx_bytes = make_export_xlsx(df, launcher_name=launcher or "")
    st.download_button(
        "Download Excel",
        data=xlsx_bytes,
        file_name="schedule.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    #buf = io.BytesIO()
    #img.save(buf, format="PNG")
    #st.download_button("Download PNG", data=buf.getvalue(), file_name="schedule.png", mime="image/png")
#else:
    #st.info("Upload both the **Routes** and **ZoneMap** files to generate the schedule.")
