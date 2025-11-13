
import streamlit as st
import pandas as pd
import re
from io import BytesIO
import numpy as np
import os

# --- Configuration & Setup ---
st.set_page_config(
    page_title="Excel Sheet Data Validator",
    layout="wide",
    initial_sidebar_state="auto"
)

# ============================== 
# Core utility functions 
# ==============================
def _detect_engine(fname: str) -> str: 
    return "pyxlsb" if fname.lower().endswith(".xlsb") else "openpyxl"

@st.cache_data 
def get_excel_sheet_names(uploaded_file): 
    """Reads sheet names from the uploaded Excel file.""" 
    try: 
        engine = _detect_engine(uploaded_file.name) 
        uploaded_file.seek(0) 
        return pd.ExcelFile(uploaded_file, engine=engine).sheet_names 
    except Exception as e: 
        st.error(f"Error reading file structure: {e}. Ensure the file is a valid Excel format.") 
        return []


@st.cache_data
def clean_key_value(value):
    """ Standardizes a field value for comparison """
    if pd.isna(value) or value is None:
        return '_NULL_'
    if isinstance(value, (int, float)) and value == int(value):
        cleaned_value = str(int(value))
    else:
        cleaned_value = str(value)
    cleaned_value = cleaned_value.strip()
    if not cleaned_value or cleaned_value.lower() in ('nan', 'none', 'null', ''):
        return '_NULL_'
    return cleaned_value

@st.cache_data
def clean_unique_id(value):
    """ Cleans Unique ID for partial match indexing. Returns empty string for NaN. """
    if pd.isna(value) or value is None:
        return ''
    if isinstance(value, (int, float)) and value == int(value):
        cleaned_value = str(int(value))
    else:
        cleaned_value = str(value)
    cleaned_value = cleaned_value.strip()
    cleaned_value = re.sub(r'[\n\r\t]', '', cleaned_value)
    return cleaned_value

@st.cache_data
def generate_key_df(df, key_cols_list):
    """ Generates the 'Key' column by concatenating the specified list of columns. """
    if df is None or df.empty:
        return None, "DataFrame is empty."
    try:
        df_temp = df[key_cols_list].applymap(lambda x: clean_key_value(x))
    except KeyError as e:
        return None, f"Error: Key column '{e.args[0]}' not found in the selected sheet."
    df = df.copy()
    df['Key'] = df_temp.apply(lambda row: '|'.join(row), axis=1)
    return df, None

def sanitize_sheet_name(name: str) -> str:
    bad = '[]:*?/\\'
    for ch in bad:
        name = name.replace(ch, "-")
    name = name.strip().strip("'")
    return name[:31] if len(name) > 31 else name

def to_excel_multi(pairs_results, out_filename="Validated_MultiPairs.xlsx"):
    """
    pairs_results: list of tuples (sheet_title_1, df1, sheet_title_2, df2) in desired order.
    df1 already includes header+metadata rows; df2 has normal header.
    """
    bio = BytesIO()
    used = set()
    with pd.ExcelWriter(bio, engine="openpyxl") as writer:
        for (t1, df1, t2, df2) in pairs_results:
            n1, n2 = sanitize_sheet_name(t1), sanitize_sheet_name(t2)
            base1, base2 = n1, n2
            k = 1
            while n1 in used:
                suf = f"_{k}"
                n1 = sanitize_sheet_name((base1[: (31 - len(suf))]) + suf)
                k += 1
            used.add(n1)
            k = 1
            while n2 in used:
                suf = f"_{k}"
                n2 = sanitize_sheet_name((base2[: (31 - len(suf))]) + suf)
                k += 1
            used.add(n2)

            df1.to_excel(writer, sheet_name=n1, index=False, header=False)
            df2.to_excel(writer, sheet_name=n2, index=False)
    return bio.getvalue(), out_filename

def cols_from_range(all_cols, start_col, end_col):
    """Return contiguous columns in sheet order, inclusive, irrespective of selection order."""
    if not all_cols or start_col not in all_cols or end_col not in all_cols:
        return []
    si, ei = all_cols.index(start_col), all_cols.index(end_col)
    if si > ei:
        si, ei = ei, si
    return all_cols[si:ei+1]

# ==============================
# Main validation logic
# ==============================
@st.cache_data(show_spinner="Running detailed data validation and restructuring ...")
def run_validation(uploaded_file, sheet1_name, sheet2_name, key_cols_list, unique_id_col, key_select_mode, metadata_rows_after_header):
    """
    Loads data, runs the validation logic using the selected parameters, and restructures the output.
    - Supports per-pair/global metadata_rows_after_header.
    - Key selection mode string is accepted for transparency; logic uses resolved key_cols_list.
    """
    if not key_cols_list:
        return None, None, "Please select at least one column for the Key."

    engine = _detect_engine(uploaded_file.name)

    # --- 1) Load Data with custom header/metadata handling for Sheet1 ---
    try:
        # 1) metadata block (header row + N metadata rows) from Sheet1
        uploaded_file.seek(0)
        sheet1_metadata = pd.read_excel(
            uploaded_file, sheet_name=sheet1_name, engine=engine,
            header=None, nrows=1 + int(metadata_rows_after_header)
        )

        # 2) data rows of Sheet1 (header=0), skipping the metadata rows after header
        skiprows = list(range(1, 1 + int(metadata_rows_after_header))) if int(metadata_rows_after_header) > 0 else None
        uploaded_file.seek(0)
        df1_data = pd.read_excel(
            uploaded_file, sheet_name=sheet1_name, engine=engine,
            header=0, skiprows=skiprows
        )

        # 3) Sheet2 (normal single-row header)
        uploaded_file.seek(0)
        df2 = pd.read_excel(uploaded_file, sheet_name=sheet2_name, engine=engine)

    except Exception as e:
        return None, None, f"Data Loading Error : {e}"

    # Normalize column names to lower for robust matching; keep original copies for output ordering
    df1_data.columns = [str(c).strip() for c in df1_data.columns]
    df2.columns = [str(c).strip() for c in df2.columns]
    key_cols_list = [str(c).strip() for c in key_cols_list]
    unique_id_col = str(unique_id_col).strip()

    # --- 2) Key Generation and Setup ---
    df1_data, error1 = generate_key_df(df1_data, key_cols_list)
    df2, error2 = generate_key_df(df2.copy(), key_cols_list)
    if df1_data is None or df2 is None:
        return None, None, (error1 or error2)

    sheet2_keys_set = set(df2['Key'])
    sheet1_keys_set = set(df1_data['Key'])

    # Columns to compare per-field (exclude unique_id_col itself)
    key_cols_for_comparison = [col for col in key_cols_list if col != unique_id_col]

    # Build multi-map for duplicate unique IDs on both sides
    def _build_uid_map(df, uid_col, cmp_cols):
        tmp_uid = df[uid_col].apply(clean_unique_id)
        df_local = df.copy()
        df_local['__uid__'] = tmp_uid
        return (
            df_local.groupby('__uid__')[cmp_cols + ['Key']]
                    .apply(lambda g: g.to_dict('records'))
                    .to_dict()
        )

    df2_map_multi = _build_uid_map(df2, unique_id_col, key_cols_for_comparison)
    df1_map_multi = _build_uid_map(df1_data, unique_id_col, key_cols_for_comparison)

    def _validate_against(other_fullkeys, other_uid_map, row, cmp_cols, uid_col_label):
        # 1) Full key match first
        full_key = row['Key']
        if full_key in other_fullkeys:
            return 'Reconcile: Complete'

        # 2) UID fallback
        uid = clean_unique_id(row.get(unique_id_col, None))
        candidates = other_uid_map.get(uid, [])
        if candidates:
            # if any candidate matches all cmp fields -> key missing but material match
            for cand in candidates:
                if all(clean_key_value(row.get(c)) == clean_key_value(cand.get(c)) for c in cmp_cols):
                    return 'Error: Full Key Missing (Lookup Key Match Only)'
            # otherwise report a mismatch list based on best (fewest mismatches)
            best = None
            for cand in candidates:
                mismatched = [c for c in cmp_cols if clean_key_value(row.get(c)) != clean_key_value(cand.get(c))]
                if best is None or len(mismatched) < len(best):
                    best = mismatched
            if best:
                return f"Error: Key Mismatch in {', '.join(best)}"
            return "Error: Key Mismatch (undetermined)"
        # 3) No UID match in the other sheet
        return f"Error: No Match Found ({uid_col_label} Missing)"

    # Apply for each side
    df1_data['Status'] = df1_data.apply(
        lambda r: _validate_against(sheet2_keys_set, df2_map_multi, r, key_cols_for_comparison, unique_id_col),
        axis=1
    )
    df2['Status'] = df2.apply(
        lambda r: _validate_against(sheet1_keys_set, df1_map_multi, r, key_cols_for_comparison, unique_id_col),
        axis=1
    )

    # --- 3) Data Restructuring (Output Construction) ---
    # ใช้คอลัมน์จริงจาก df1_data เป็นฐาน แล้วเพิ่ม Key/Status ไว้หน้าสุด

    # 3.1 ensure Key/Status exist
    if 'Key' not in df1_data.columns:
        df1_data['Key'] = ''
    if 'Status' not in df1_data.columns:
        df1_data['Status'] = ''

    # 3.2 final headers: Key, Status + original data columns (ตามจริงจาก df1_data)
    original_cols = [c for c in df1_data.columns if c not in ['Key', 'Status']]
    final_headers = ['Key', 'Status'] + original_cols

    # 3.3 สร้าง data frame ฝั่งข้อมูล ให้คอลัมน์ตรงกับ final_headers เป๊ะ
    df1_data_output = df1_data[['Key', 'Status'] + original_cols].copy()
    df1_data_output.columns = final_headers  # บังคับให้ชื่อคอลัมน์ตรงกับ final_headers

    # 3.4 สร้าง header row = final_headers (หนึ่งแถว) ด้วย columns = final_headers
    header_row_df = pd.DataFrame([final_headers], columns=final_headers)

    # 3.5 เตรียม metadata rows (ถ้ามี): ข้ามแถว header เดิมของไฟล์ แล้วจัดความยาวให้เท่ากับ original_cols
    metadata_rows_list = sheet1_metadata.fillna('').values.tolist()
    target_len = 1 + int(metadata_rows_after_header)   # 1 = แถว header เดิมในไฟล์
    if len(metadata_rows_list) < target_len:
        # เติมให้ครบตามจำนวนที่ผู้ใช้ระบุ
        metadata_rows_list += [[''] * len(metadata_rows_list[0]) for _ in range(target_len - len(metadata_rows_list))]
    elif len(metadata_rows_list) > target_len:
        metadata_rows_list = metadata_rows_list[:target_len]

    # ตัดเอาเฉพาะแถว metadata (ไม่เอา header เดิมของไฟล์)
    metadata_only = metadata_rows_list[1:] if target_len > 1 else []

    # pad/truncate แถว metadata ให้ยาวเท่าจำนวน original_cols แล้วเติมช่องว่าง 2 ช่องหน้า (Key/Status)
    fixed_meta_rows = []
    for row in metadata_only:
        row_fixed = list(row[:len(original_cols)])
        if len(row_fixed) < len(original_cols):
            row_fixed += [''] * (len(original_cols) - len(row_fixed))
        fixed_meta_rows.append(['', ''] + row_fixed)

    # 3.6 สร้าง DataFrame metadata โดยใช้ columns = final_headers
    final_metadata_sheet1 = (
        pd.DataFrame(fixed_meta_rows, columns=final_headers)
        if fixed_meta_rows else pd.DataFrame(columns=final_headers)
    )

    # 3.7 รวม header + metadata + data ด้วย columns เดียวกัน -> ไม่เหลื่อม
    final_df1 = pd.concat([header_row_df, final_metadata_sheet1, df1_data_output], ignore_index=True)


    # --- Sheet 2 (Header normal + Key/Status first) ---
    all_cols_df2 = df2.columns.tolist()
    final_sheet2_cols_order = ['Key', 'Status'] + [c for c in all_cols_df2 if c not in ['Key', 'Status']]
    final_df2 = df2[final_sheet2_cols_order]

    return final_df1, final_df2, None

# ==============================
# Streamlit UI Layout
# ==============================
st.title("Multi-Sheet Data Validator")
st.markdown("1. Upload your Excel (.xlsx or .xlsb) file")

# Add your own warning (because 200MB upload limit ≠ safe processing)
st.markdown(
    """
    <div style='color:#777; font-size:15px; margin-bottom:8px;'>
        <strong>Note :</strong> The server can only process files up to about 
        <strong>30MB</strong> safely.  
        Large files may crash due to memory limits, even though the uploader shows a 200MB limit.
    </div>
    """,
    unsafe_allow_html=True
)

uploaded_file = st.file_uploader(
    "Upload Excel File for Validation :",
    type=['xlsx', 'xlsb', 'xls'],
    accept_multiple_files=False
)

sheet_names = []
if uploaded_file:
    sheet_names = get_excel_sheet_names(uploaded_file)
    if sheet_names:
        st.success(f"Uploaded **{uploaded_file.name}** — detected {len(sheet_names)} sheets.")
    else:
        st.warning("No sheets detected.")

# =========================
# Dynamic pairs + metadata mode
# =========================
if "pairs" not in st.session_state:
    st.session_state.pairs = []  # each item is a dict per pair

def add_pair():
    s1 = sheet_names[0] if sheet_names else None
    s2 = (sheet_names[1] if len(sheet_names) > 1 else (sheet_names[0] if sheet_names else None))
    # try fetch first-row headers for default kcols/uid
    engine = _detect_engine(uploaded_file.name) if uploaded_file else "openpyxl"
    s1_cols = []
    if uploaded_file and s1:
        try:
            uploaded_file.seek(0)
            df_cols = pd.read_excel(uploaded_file, sheet_name=s1, engine=engine, header=0, nrows=1)
            s1_cols = [str(c).strip() for c in df_cols.columns if not pd.isna(c)]
        except Exception:
            s1_cols = []
    default_keys = s1_cols[:1] if s1_cols else []
    default_uid = default_keys[0] if default_keys else (s1_cols[0] if s1_cols else None)

    st.session_state.pairs.append({
        "sheet1_name": s1,
        "sheet2_name": s2,
        "metadata_rows_after_header": 0,
        "key_select_mode": "Define Range (Contiguous)",
        "key_cols_list": default_keys,
        "kstart": (s1_cols[0] if s1_cols else None),
        "kend": (s1_cols[-1] if s1_cols else None),
        "unique_id_col": default_uid,
    })

def remove_pair(idx: int):
    st.session_state.pairs.pop(idx)

if uploaded_file and sheet_names:
    st.markdown("2. Configure Comparison Parameters")

    # ---- Metadata rows mode (Global vs Per pair) ----
    if "metadata_mode" not in st.session_state:
        st.session_state.metadata_mode = "Per pair (custom for each)"
    metadata_mode = st.radio(
        "Metadata rows mode :",
        ("Per pair (custom for each)", "Global (same for all pairs)"),
        index=(0 if st.session_state.metadata_mode.startswith("Per pair") else 1),
        horizontal=True
    )
    st.session_state.metadata_mode = metadata_mode

    global_mrows = 0
    if metadata_mode == "Global (same for all pairs)":
        global_mrows = st.number_input(
            "Global - metadata rows AFTER header (applies to all pairs) :",
            min_value=0, max_value=50, value=0, step=1
        )
        st.caption(f"Using global value : {global_mrows} for all pairs.")

    st.button("➕ Add Comparison Pair", on_click=add_pair)
    if not st.session_state.pairs:
        add_pair()

    pair_configs = []

    for i, cfg in enumerate(st.session_state.pairs):
        with st.expander(f"--- Pair {i+1} ---", expanded=True):
            c_top = st.columns([1, 1, 0.2])
            with c_top[0]:
                s1 = st.selectbox(
                    f"[Pair {i+1}] Sheet 1 (Source) :",
                    options=sheet_names,
                    index=sheet_names.index(cfg["sheet1_name"]) if cfg["sheet1_name"] in sheet_names else 0,
                    key=f"s1_{i}"
                )
            with c_top[1]:
                default_idx = sheet_names.index(cfg["sheet2_name"]) if cfg["sheet2_name"] in sheet_names else (1 if len(sheet_names) > 1 else 0)
                s2 = st.selectbox(
                    f"[Pair {i+1}] Sheet 2 (Target) :",
                    options=sheet_names,
                    index=default_idx,
                    key=f"s2_{i}"
                )
            with c_top[2]:
                st.button("🗑️", key=f"del_{i}", help="Remove this pair", on_click=remove_pair, args=(i,))

            # Per-pair metadata rows input (disabled if in Global mode)
            mrows = st.number_input(
                f"[Pair {i+1}] Metadata rows AFTER header :",
                min_value=0, max_value=50,
                value=int(cfg.get("metadata_rows_after_header", 0)),
                step=1, key=f"mrows_{i}",
                disabled=(metadata_mode == "Global (same for all pairs)")
            )
            if metadata_mode == "Global (same for all pairs)":
                st.caption(f"Using global value : {global_mrows} (per-pair input disabled)")

            # Load columns from Sheet1 to drive key selection
            engine = _detect_engine(uploaded_file.name)
            try:
                uploaded_file.seek(0)
                df_cols = pd.read_excel(uploaded_file, sheet_name=s1, engine=engine, header=0, nrows=1)
                all_cols_s1 = [str(c).strip() for c in df_cols.columns if not pd.isna(c)]
            except Exception:
                all_cols_s1 = []

            kmode = st.radio(
                f"[Pair {i+1}] Key Selection Mode :",
                ("Select Columns (Non-Contiguous)", "Define Range (Contiguous)"),
                horizontal=True,
                index=(0 if cfg.get("key_select_mode :","").startswith("Select") else 1),
                key=f"ksel_{i}"
            )

            if kmode == "Select Columns (Non-Contiguous)":
                kcols = st.multiselect(
                    f"[Pair {i+1}] Select Key Columns (order matters)",
                    options=all_cols_s1, default=[c for c in cfg.get("key_cols_list", []) if c in all_cols_s1], key=f"kcols_{i}"
                )
                uid_options = kcols if kcols else all_cols_s1
                kstart = cfg.get("kstart", all_cols_s1[0] if all_cols_s1 else None)
                kend = cfg.get("kend", all_cols_s1[-1] if all_cols_s1 else None)
            else:
                r1, r2 = st.columns(2)
                with r1:
                    kstart = st.selectbox(f"[Pair {i+1}] Key Start Column :", options=all_cols_s1,
                                          index=(all_cols_s1.index(cfg.get("kstart")) if cfg.get("kstart") in all_cols_s1 else (0 if all_cols_s1 else 0)),
                                          key=f"kstart_{i}")
                with r2:
                    kend = st.selectbox(f"[Pair {i+1}] Key End Column :", options=all_cols_s1,
                                        index=(all_cols_s1.index(cfg.get("kend")) if cfg.get("kend") in all_cols_s1 else (len(all_cols_s1)-1 if all_cols_s1 else 0)),
                                        key=f"kend_{i}")
                kcols = cols_from_range(all_cols_s1, kstart, kend) if all_cols_s1 else []
                uid_options = kcols if kcols else all_cols_s1
                st.caption(f"Contiguous key columns : {kcols}")

            uid = st.selectbox(
                f"[Pair {i+1}] Lookup Key column (used for partial match) :",
                options=uid_options,
                index=(uid_options.index(cfg.get("unique_id_col")) if cfg.get("unique_id_col") in uid_options else (0 if uid_options else 0)),
                key=f"uid_{i}",
                disabled=(not uid_options)
            )

            # Persist to session state
            cfg.update(dict(
                sheet1_name=s1,
                sheet2_name=s2,
                metadata_rows_after_header=int(mrows),
                key_select_mode=kmode,
                key_cols_list=kcols,
                kstart=kstart,
                kend=kend,
                unique_id_col=uid
            ))

            ok = all([s1, s2]) and bool(kcols) and (uid is not None and uid != "")
            pair_configs.append(cfg if ok else None)

    # =========================
    # Run all pairs & download
    # =========================
    if st.button("3. Run All Comparisons & Download Excel", type="primary"):
        errors, results = [], []
        with st.spinner("Running comparisons..."):
            for idx, cfg in enumerate(pair_configs, start=1):
                if not cfg:
                    errors.append(f"Pair {idx}: incomplete configuration.")
                    continue

                # choose effective metadata rows
                effective_mrows = global_mrows if metadata_mode == "Global (same for all pairs)" else cfg["metadata_rows_after_header"]

                try:
                    df1_out, df2_out, err = run_validation(
                        uploaded_file=uploaded_file,
                        sheet1_name=cfg["sheet1_name"],
                        sheet2_name=cfg["sheet2_name"],
                        key_cols_list=cfg["key_cols_list"],
                        unique_id_col=cfg["unique_id_col"],
                        key_select_mode=cfg["key_select_mode"],
                        metadata_rows_after_header=int(effective_mrows)
                    )
                    if err:
                        errors.append(f"Pair #{idx} ({cfg['sheet1_name']} vs {cfg['sheet2_name']}) : {err}")
                    else:
                        title1 = f"Validated_{cfg['sheet1_name']}"
                        title2 = f"Validated_{cfg['sheet2_name']}"
                        results.append((title1, df1_out, title2, df2_out))
                except Exception as e:
                    errors.append(f"Pair #{idx} failed : {e}")

        if results:
            base = os.path.splitext(uploaded_file.name)[0]
            xbytes, fname = to_excel_multi(results, out_filename=f"{base}_Validated_Output.xlsx")
            st.success(f"Completed {len(results)} pair(s).")
            st.download_button(
                label="Download Combined Validated Workbook",
                data=xbytes,
                file_name=fname,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="dl_all_pairs"
            )
        if errors:
            st.error("Some pairs had issues :\n" + "\n".join(f"- {msg}" for msg in errors))

else:
    st.info("Upload an Excel file to begin.")
