import streamlit as st
import os
password = os.getenv("APP_PASSWORD")
user_input = st.text_input("Enter password:", type="password")
if password and user_input != password:
    st.stop()
import pandas as pd
import numpy as np
import re
from typing import Dict, Any
import io

# -------------------------------
# Page & Style
# -------------------------------
st.set_page_config(page_title="Energy KPIs App", page_icon="📈", layout="wide")
st.markdown(
    """
    <style>
    .stApp {
        background: linear-gradient(135deg, #e6f2ff 0%, #f5fbff 60%, #ffffff 100%);
        font-family: 'Inter', system-ui, -apple-system, Segoe UI, Roboto, Helvetica, Arial, sans-serif;
    }
    .app-card {
        background: rgba(255,255,255,0.88);
        border-radius: 20px;
        padding: 24px;
        box-shadow: 0 8px 24px rgba(67, 133, 245, 0.15);
        border: 1px solid rgba(67,133,245,0.12);
    }
    .title {font-size:30px;font-weight:800;margin-bottom:4px;color:#0f3d78;}
    .subtitle {color:#326ba8;font-size:16px;margin-bottom:12px;}
    .divider {height:1px;background:linear-gradient(to right,rgba(67,133,245,0.15),rgba(67,133,245,0.05));margin:16px 0;}
    </style>
    """,
    unsafe_allow_html=True
)

st.markdown('<div class="app-card">', unsafe_allow_html=True)
st.markdown('<div class="title">Welcome 👋</div>', unsafe_allow_html=True)
st.markdown(
    '<div class="subtitle">Upload your Excel with sheets <b>Site data</b> and <b>Parameters</b>. '
    "I'll parse variables, values, and units for the next calculation step.</div>",
    unsafe_allow_html=True
)

# -------------------------------
# Helpers
# -------------------------------
def _clean_name(s: str) -> str:
    s = str(s or "").strip().lower()
    s = re.sub(r"\(.*?\)", "", s)           # remove anything in parentheses (units)
    s = s.replace("/", "_").replace("-", "_")
    s = re.sub(r"[^a-z0-9_]+", "_", s)
    s = re.sub(r"_+", "_", s).strip("_")
    return s

def _to_bool(x):
    if x is None or (isinstance(x, float) and pd.isna(x)): return None
    s = str(x).strip().lower()
    if s in ("true","y","yes","1"):  return True
    if s in ("false","n","no","0"):  return False
    return None

def _to_float(x):
    if x is None or (isinstance(x, float) and pd.isna(x)): return None
    if isinstance(x, (int, float)): return float(x)
    m = re.search(r"[-+]?\d*\.?\d+(?:[eE][-+]?\d+)?", str(x))
    return float(m.group(0)) if m else None

def read_parameters_sheet(xls: pd.ExcelFile, sheet_name: str = "Parameters") -> pd.DataFrame:
    """Return DF with columns: name, value, unit (name is normalized)."""
    dfp = pd.read_excel(xls, sheet_name=sheet_name, header=None)
    records = []
    for _, row in dfp.iterrows():
        key  = row.iloc[0] if len(row) > 0 else None
        val  = row.iloc[1] if len(row) > 1 else None
        unit = row.iloc[2] if len(row) > 2 else None
        if pd.isna(key): 
            continue
        key_clean = _clean_name(key)
        if not key_clean:
            continue
        v_num = _to_float(val)
        records.append({
            "name":  key_clean,
            "value": v_num if v_num is not None else val,
            "unit":  (str(unit).strip() if pd.notna(unit) else "")
        })
    return pd.DataFrame(records)

def params_to_dict(params_df: pd.DataFrame) -> Dict[str, Any]:
    return {} if params_df is None or params_df.empty else {r["name"]: r["value"] for _, r in params_df.iterrows()}

def read_site_data_sheet(xls: pd.ExcelFile, sheet_name: str = "Site data") -> pd.DataFrame:
    """
    Standardizes to:
      time, i_batt_a, v_batt_v, i_load_a, gen_signal_on, grid_on, i_rectifier_a, i_solar_a
    """
    df = pd.read_excel(xls, sheet_name=sheet_name, header=0)
    cols0 = [_clean_name(c) for c in df.columns]
    if "time" not in cols0:
        df = pd.read_excel(xls, sheet_name=sheet_name, header=1)
    df.columns = [_clean_name(c) for c in df.columns]

    # Drop a units row if present
    if df.shape[0] > 0:
        first_row_str = " ".join([str(x).lower() for x in df.iloc[0].tolist()])
        if "unit" in first_row_str:
            df = df.iloc[1:].reset_index(drop=True)

    rename_map = {
        "time": "time",
        "i_batt": "i_batt_(A)",
        "v_batt": "v_batt_(V)",
        "i_load": "i_load_(A)",
        "generator_signal_on": "gen_signal_on",
        "grid_on": "grid_on",
        "i_rectifier": "i_rectifier_(A)",
        "i_solar": "i_solar_(A)",
    }
    for k, target in list(rename_map.items()):
        if k in df.columns:
            df.rename(columns={k: target}, inplace=True)
        else:
            for col in df.columns:
                if k in col and target not in df.columns:
                    df.rename(columns={col: target}, inplace=True)
                    break

    if "time" in df.columns:
        df["time"] = pd.to_datetime(df["time"], errors="coerce")
    for bcol in ["gen_signal_on", "grid_on"]:
        if bcol in df.columns:
            df[bcol] = df[bcol].apply(_to_bool)
    for ncol in ["i_batt_a", "v_batt_v", "i_load_a", "i_rectifier_a", "i_solar_a"]:
        if ncol in df.columns:
            df[ncol] = pd.to_numeric(df[ncol], errors="coerce")

    return df

# -------------------------------
# File uploader & Definition step
# -------------------------------
uploaded = st.file_uploader("Upload your Excel (.xlsx) with 'Site data' and 'Parameters'", type=["xlsx"])

if uploaded is not None:
    try:
        xls = pd.ExcelFile(uploaded)

        # Validate sheets
        sheet_names_lower = [s.lower() for s in xls.sheet_names]
        if "parameters" not in sheet_names_lower:
            st.error("Missing sheet: 'Parameters'."); st.stop()
        if "site data" not in sheet_names_lower:
            st.error("Missing sheet: 'Site data'."); st.stop()

        # Read & parse
        params_df = read_parameters_sheet(xls, sheet_name="Parameters")
        site_df   = read_site_data_sheet(xls, sheet_name="Site data")

        st.success("Excel loaded successfully.")
        c1, c2 = st.columns([1, 1])

        with c1:
            st.markdown("<h2 style='color:#6a1b9a;'>⚙️ Parameters</h2>", unsafe_allow_html=True)
            st.dataframe(params_df, use_container_width=True)

        with c2:
            st.markdown("<h2 style='color:#000000;'>📊 Site Data (Flows)</h2>", unsafe_allow_html=True)
            st.dataframe(site_df.head(200), use_container_width=True)

        st.markdown('<div class="divider"></div>', unsafe_allow_html=True)

        params = params_to_dict(params_df)
        st.session_state["params"]  = params
        st.session_state["site_df"] = site_df

        # -------------------------------
        # Internal Flows Calculations
        # -------------------------------
        if site_df is not None and not site_df.empty:

            df = site_df.copy()

            # Validate required columns
            required_cols = ["i_rectifier_(A)","i_batt_(A)","i_load_(A)","i_solar_(A)","v_batt_(V)"]
            for col in required_cols:
                if col not in df.columns:
                    st.error(f"Missing column in 'Site data': {col}")
                    st.stop()
                df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0.0)

            # Signals & params
            gen_on_series = (df.get("gen_signal_on", False)).fillna(False)
            gen_on = gen_on_series.to_numpy(dtype=bool)

            I_rect = df["i_rectifier_(A)"].to_numpy(float)
            I_batt = df["i_batt_(A)"].to_numpy(float)
            I_load = df["i_load_(A)"].to_numpy(float)
            I_solar = df["i_solar_(A)"].to_numpy(float)
            V_batt = df["v_batt_(V)"].to_numpy(float)

            eta_rect   = float(params.get("n_rect", 1.0))
            p_gen_max  = float(params.get("p_gen_max", np.nan))
            sfc_a      = float(params.get("sfc_a", 0.0))
            sfc_b      = float(params.get("sfc_b", 0.0))
            STEP_HOURS = 1.0 / 6.0  # 10 minutes

            # -------- Currents (A)
            I_Gen        = np.where(gen_on, I_rect, 0.0)
            I_grid       = I_rect - I_Gen
            cond_Gen_load = (I_Gen > 0) & (I_load > I_solar) & (I_batt > 0)
            I_Gen_load   = np.where(cond_Gen_load, np.minimum(I_Gen, I_load - I_solar), 0.0)
            I_Gen_batt   = np.where(I_batt > 0, np.clip(I_Gen - I_Gen_load, 0.0, None), 0.0)
            I_grid_batt  = np.where((I_grid > 0) & (I_batt > 0), np.clip(I_grid - I_solar - I_load, 0.0, None), 0.0)
            I_grid_load  = I_grid - I_grid_batt
            I_batt_load  = np.where(I_batt < 0, I_batt, 0.0)
            I_solar_load = np.minimum(I_solar, I_load)
            I_solar_batt = np.where(I_batt > 0, np.clip(I_solar - I_solar_load, 0.0, None), 0.0)

            # -------- Powers (W)
            loss          = (1.0 - eta_rect)
            P_Gen         = (I_Gen  * V_batt) * (1.0 + loss)
            P_grid        = (I_grid * V_batt) * (1.0 + loss)
            P_solar       =  I_solar * V_batt
            P_disch_batt  =  I_batt_load * V_batt
            P_ch_batt     = np.where(I_batt > 0, I_batt * V_batt, 0.0)
            P_ch_batt_estim = (I_Gen_batt + I_solar_batt + I_grid_batt) * V_batt
            P_load        =  I_load * V_batt

            # -------- Fuel (L)
            P_Gen_kW        = P_Gen * 1e-3
            pct_load        = np.where(P_Gen > 0, np.where(np.isfinite(p_gen_max) & (p_gen_max > 0), P_Gen_kW / p_gen_max, 0.0), 0.0)
            SFC             = np.where(pct_load > 0, sfc_a * (pct_load ** sfc_b), 0.0)    # (L/kWh)
            fuel_consumption= SFC * P_Gen_kW * STEP_HOURS                                   # (L) per step
            TOTAL_FUEL      = float(fuel_consumption.sum())                                 # (L)

            # -------- Build DF (raw names)
            internal_raw = pd.DataFrame({
                "I_Gen": I_Gen, "I_grid": I_grid, "I_Gen_load": I_Gen_load, "I_Gen_batt": I_Gen_batt,
                "I_grid_batt": I_grid_batt, "I_grid_load": I_grid_load, "I_batt_load": I_batt_load,
                "I_solar_load": I_solar_load, "I_solar_batt": I_solar_batt,
                "P_Gen": P_Gen, "P_grid": P_grid, "P_solar": P_solar, "P_disch_batt": P_disch_batt,
                "P_ch_batt": P_ch_batt, "P_ch_batt_estim": P_ch_batt_estim, "P_load": P_load,
                "load_ratio": pct_load, "SFC": SFC, "fuel_consumption": fuel_consumption
            })

            # -------- Add units to headers for display/export
            units = {
                "I_Gen":"(A)","I_grid":"(A)","I_Gen_load":"(A)","I_Gen_batt":"(A)",
                "I_grid_batt":"(A)","I_grid_load":"(A)","I_batt_load":"(A)",
                "I_solar_load":"(A)","I_solar_batt":"(A)",
                "P_Gen":"(W)","P_grid":"(W)","P_solar":"(W)","P_disch_batt":"(W)",
                "P_ch_batt":"(W)","P_ch_batt_estim":"(W)","P_load":"(W)",
                "load_ratio":"(-)","SFC":"(l/kWh)","fuel_consumption":"(l)"
            }
            internal = internal_raw.copy()
            internal.columns = [f"{c} {units.get(c,'')}".strip() for c in internal.columns]

            # -------- UI
            st.markdown("<h2 style='color:#0a7e07;'>🔧 Internal Flows</h2>", unsafe_allow_html=True)
            st.dataframe(internal.head(200), use_container_width=True)
            st.metric(label="TOTAL_FUEL (l)", value=f"{TOTAL_FUEL:,.3f}")
            st.markdown('<div class="divider"></div>', unsafe_allow_html=True)

            # -------- KPIs Calculations --------
            dt_hours = STEP_HOURS  # 10 min = 1/6 h
            kpi_rows = []

            def add_kpi(category, name, value, unit):
                try:
                    val = float(value)
                except Exception:
                    val = np.nan
                kpi_rows.append({"Category": category, "KPI": name, "Value": round(val, 4), "Unit": unit})

            # --- 1) Generator KPIs ---
            Running_hours = (1/6) * np.sum(I_Gen > 0)
            Num_of_starts = np.sum((I_Gen[1:] > 0) & (I_Gen[:-1] == 0))
            Avg_power = np.mean(P_Gen[P_Gen > 0]) * 0.001 if np.any(P_Gen > 0) else 0
            Total_Energy_Consumption = np.sum(P_Gen) * 0.001 * (10/60)
            Total_Fuel_consumption = TOTAL_FUEL

            add_kpi("Generator", "Running_hours", Running_hours, "h")
            add_kpi("Generator", "Num_of_starts", Num_of_starts, "starts")
            add_kpi("Generator", "Avg_power", Avg_power, "kW")
            add_kpi("Generator", "Total_Energy_Consumption", Total_Energy_Consumption, "kWh")
            add_kpi("Generator", "Total_Fuel_consumption (model)", Total_Fuel_consumption, "l")

            # --- 2) Battery KPIs ---
            Energy_IN = 0.001 * np.sum(P_ch_batt) * (10/60)
            Energy_OUT = 0.001 * np.sum(P_disch_batt) * (10/60)
            Autonomy = np.sum(I_batt < 0) * (10/60)
            Cycle_transitions = int(np.sum((I_batt[1:] > 0) & (I_batt[:-1] < 0)))

            add_kpi("Battery", "Energy_IN", Energy_IN, "kWh")
            add_kpi("Battery", "Energy_OUT", Energy_OUT, "kWh")
            add_kpi("Battery", "Autonomy", Autonomy, "h")
            add_kpi("Battery", "Cycle_transitions", Cycle_transitions, "-")

            # --- 3) Load KPIs ---
            Load_Average = np.mean(P_load) * 0.001
            Load_Std = np.std(P_load, ddof=0) * 0.001

            add_kpi("Load", "Average", Load_Average, "kW")
            add_kpi("Load", "Std", Load_Std, "kW")

            # --- 4) Grid KPIs ---
            Uptime = np.sum(P_grid > 0) * (10/60)
            Energy_Outage = np.sum(P_grid) * 0.001 * (10/60)
            Avg_Consumption_power = 0.001 * np.mean(P_grid[P_grid > 0]) if np.any(P_grid > 0) else 0

            add_kpi("Grid", "Uptime", Uptime, "h")
            add_kpi("Grid", "Energy_Outage", Energy_Outage, "kWh")
            add_kpi("Grid", "Avg_Consumption_power", Avg_Consumption_power, "kW")

            # --- 5) PV KPIs ---
            Generation = np.sum(P_solar) * 0.001 * (10/60)
            Utilization_factor = 100 * Generation / (np.sum(P_load) * 0.001 * (10/60)) if np.sum(P_load) > 0 else 0

            add_kpi("PV", "Generation", Generation, "kWh")
            add_kpi("PV", "Utilization_factor", Utilization_factor, "%")

            kpis_df = pd.DataFrame(kpi_rows)

            # -------- Display KPIs --------
            st.markdown("<h2 style='color:#b91c1c;'>🏁 KPIs Summary</h2>", unsafe_allow_html=True)
            st.dataframe(kpis_df, use_container_width=True)
            st.markdown('<div class="divider"></div>', unsafe_allow_html=True)

            # -------- Export (openpyxl) --------
            buf = io.BytesIO()
            try:
                with pd.ExcelWriter(buf, engine="openpyxl") as w:
                    params_df.to_excel(w, index=False, sheet_name="Parameters")
                    site_df.to_excel(w, index=False, sheet_name="Site data")
                    internal.to_excel(w, index=False, sheet_name="Internal Flows")
                    kpis_df.to_excel(w, index=False, sheet_name="KPIs")
                st.download_button(
                    "⬇️ Download Excel (Parameters / Site data / Internal Flows / KPIs)",
                    data=buf.getvalue(),
                    file_name="results_with_KPIs.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            except Exception as ex:
                st.warning(f"Export skipped: {ex}")

            

    except Exception as e:
        st.error(f"Could not read file: {e}")
else:
    st.warning("No file uploaded yet. Please upload your Excel file.")

st.markdown('</div>', unsafe_allow_html=True)
