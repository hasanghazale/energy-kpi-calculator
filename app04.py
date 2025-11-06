import streamlit as st
import os
import base64
from pathlib import Path
import pandas as pd
import numpy as np
import re
from typing import Dict, Any
import io

# --------------------------------

# --- Custom style / branding ---
st.markdown("""
<style>
.app-card {
    background: transparent !important;
    border: none !important;
    box-shadow: none !important;
    padding: 0 !important;
}
</style>
""", unsafe_allow_html=True)

# -------------------------------
# Page & Style
# -------------------------------
st.set_page_config(page_title="Energy KPIs Calculator - by R&D team", page_icon="📈", layout="wide")
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
st.markdown('<div class="title">Welcome to the Energy KPIs Analysis Platform ⚙️</div>', unsafe_allow_html=True)
st.markdown(
    '<div class="subtitle">Please upload your Excel file with sheets <b>Site data</b> and <b>Parameters</b> <i>(use the Excel template provided by the team, filled by the data you aim to analyze)</i>.'
    " Variables, values, and units will be parsed for the next calculation steps.</div>",
    unsafe_allow_html=True
)
def _img_b64(path: str) -> str:
    p = Path(path)
    if not p.exists():
        return ""
    return base64.b64encode(p.read_bytes()).decode()

LOGO_B64 = _img_b64("logo2.png")

# Header with logo (base64, so no path issues)
st.markdown(f"""
<div style="display:flex; align-items:center; gap:14px; margin:8px 0 8px;">
  <img src="data:image/png;base64,{LOGO_B64}" style="width:450px; height:auto;" />
  <h1 style="color:#0f3d78; font-weight:600; margin:0;">Energy KPIs Calculator - by R&D team</h1>
</div>
""", unsafe_allow_html=True)


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
      time, i_batt_(A), v_batt_(V), i_load_(A), gen_signal_on, grid_on, i_rectifier_(A), i_solar_(A),
      i_gen_(A), fuel_level_tank_(l)
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
        "fuel_level_tank": "fuel_level_tank_(l)",
        # >>> ADDED: read generator current if present
        "i_gen": "i_gen_(A)",
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
    for ncol in ["i_batt_(A)", "v_batt_(V)", "i_load_(A)", "i_rectifier_(A)", "i_solar_(A)",
                 "i_gen_(A)", "fuel_level_tank_(l)"]:  # >>> ADDED numeric coercion
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
            p_gen_max  = float(params.get("pmax_gen", np.nan))
            sfc_a      = float(params.get("sfc_a", 0.0))
            sfc_b      = float(params.get("sfc_b", 0.0))
            a_rate     = float(params.get("a_rate", 0.0))   # to calculate fuel flow rate
            b_rate     = float(params.get("b_rate", 0.0))
            STEP_HOURS = 1.0 / 6.0  # 10 minutes in hrs

            # -------- Currents (A)
            # >>> CHANGED: prefer measured i_gen if available; else fall back to gen_signal_on
            if "i_gen_(A)" in df.columns and df["i_gen_(A)"].notna().any():
                I_Gen = np.clip(df["i_gen_(A)"].to_numpy(float), 0.0, None)
            else:
                I_Gen = np.where(gen_on, I_rect, 0.0)

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
            P_MAX = float(params.get("Pmax_Gen", np.nan))

            # -------- Fuel (L) — MODEL (by SFC)
            P_Gen_kW        = P_Gen * 1e-3
            pct_load        = np.where(P_Gen > 0,  P_Gen_kW / p_gen_max, 0.0)
            SFC             = np.where(pct_load > 0, sfc_a * (pct_load ** sfc_b), 0.0)    # (L/kWh)
            fuel_consumption= SFC * P_Gen_kW * STEP_HOURS                                   # (L) per step
            TOTAL_FUEL      = float(fuel_consumption.sum())                                 # (L)
            # -------- Fuel_01 (L) — MODEL (by rate consumption)
            fc_rate = np.where(P_Gen_kW > 0 , a_rate*P_Gen_kW + b_rate, 0.0)    # in l/hr 
            instant_fuel_consump = fc_rate*STEP_HOURS  # in l
            TOTAL_FUEL_01 = float(instant_fuel_consump.sum())  # total in l 
            

           
#----------

            # --- Measured fuel (simple, periodized; uses only real points, NaNs allowed in a period)
            measured_fuel = 0
            if "fuel_level_tank_(l)" in df.columns:
                fuel_series = pd.to_numeric(df["fuel_level_tank_(l)"], errors="coerce")

                # Boolean mask for ON samples
                on_mask = (I_Gen > 0)
                if np.any(on_mask):
                    # Build group ids that increment ONLY at the start of each ON run
                    on_mask_s = pd.Series(on_mask, index=df.index)
                    # starts_of_runs is True when current sample is ON and previous was OFF
                    starts_of_runs = on_mask_s & ~on_mask_s.shift(fill_value=False)
                    group_id = starts_of_runs.cumsum()              # grows across the whole series
                    group_id = group_id.where(on_mask_s, np.nan)    # keep id only during ON; OFF -> NaN (dropped by groupby)

                    total = 0.0
                    # Group fuel levels by each contiguous ON period
                    for gid, s in fuel_series.groupby(group_id):
                        if pd.isna(gid):
                            continue
                        s_valid = s.dropna()
                        if s_valid.size >= 2:
                            # per-period consumption = (max - min)
                            total += float(s_valid.max() - s_valid.min())

                    measured_fuel = float(total)
                else:
                    measured_fuel = 0

            # --- Measured fuel (robust, periodized): median-smooth + sum of negative drops within each ON period
            measured_fuel_robust = 0
            if "fuel_level_tank_(l)" in df.columns:
                fuel_series = pd.to_numeric(df["fuel_level_tank_(l)"], errors="coerce")

                on_mask = (I_Gen > 0)
                if np.any(on_mask):
                    on_mask_s = pd.Series(on_mask, index=df.index)
                    # contiguous-ON run IDs: increment only at the start of each ON run
                    starts = on_mask_s & ~on_mask_s.shift(fill_value=False)
                    gid = starts.cumsum().where(on_mask_s, np.nan)  # OFF samples -> NaN => ignored by groupby

                    total = 0.0
                    any_period = False

                    for run_id, s in fuel_series.groupby(gid):
                        if pd.isna(run_id):
                            continue

                        # Option A (safe with NaNs): smooth directly with min_periods=1
                        s_smooth = s.rolling(window=3, center=True, min_periods=1).median()

                        # Deltas; only count consumption (negative steps). NaNs are ignored automatically.
                        deltas = s_smooth.diff()
                        neg = deltas[deltas < 0]

                        if not neg.empty:
                            drop_l = -float(neg.sum())  # liters consumed in this ON period
                            if drop_l > 0:
                                total += drop_l
                                any_period = True

                    measured_fuel_robust = float(total) if any_period else measured_fuel_robust==0
                else:
                    measured_fuel_robust = 0


 
            # -------- Build DF (raw names)
            internal_raw = pd.DataFrame({
                "I_Gen": I_Gen, "I_grid": I_grid, "I_Gen_load": I_Gen_load, "I_Gen_batt": I_Gen_batt,
                "I_grid_batt": I_grid_batt, "I_grid_load": I_grid_load, "I_batt_load": I_batt_load,
                "I_solar_load": I_solar_load, "I_solar_batt": I_solar_batt,
                "P_Gen": P_Gen, "P_grid": P_grid, "P_solar": P_solar, "P_disch_batt": P_disch_batt,
                "P_ch_batt": P_ch_batt, "P_ch_batt_estim": P_ch_batt_estim, "P_load": P_load,
                "load_ratio": pct_load, "SFC": SFC, "fuel_consumption_sfc": fuel_consumption, "fuel_consumption_p": instant_fuel_consump
            })

            # -------- Add units to headers for display/export
            units = {
                "I_Gen":"(A)","I_grid":"(A)","I_Gen_load":"(A)","I_Gen_batt":"(A)",
                "I_grid_batt":"(A)","I_grid_load":"(A)","I_batt_load":"(A)",
                "I_solar_load":"(A)","I_solar_batt":"(A)",
                "P_Gen":"(W)","P_grid":"(W)","P_solar":"(W)","P_disch_batt":"(W)",
                "P_ch_batt":"(W)","P_ch_batt_estim":"(W)","P_load":"(W)",
                "load_ratio":"(-)","SFC":"(l/kWh)","fuel_consumption":"(l)", "fuel_consumption_p": "(l)"
            }
            internal = internal_raw.copy()
            internal.columns = [f"{c} {units.get(c,'')}".strip() for c in internal.columns]

            # -------- UI
            st.markdown("<h2 style='color:#0a7e07;'>🔧 Internal Flows</h2>", unsafe_allow_html=True)
            st.dataframe(internal.head(200), use_container_width=True)
            st.markdown(f"<h4 style='color:red;'>Total_Estimated_FUEL_consump_SFC_model (l): {TOTAL_FUEL:,.3f}</h4>", unsafe_allow_html=True)
            st.markdown(f"<h4 style='color:red;'>Total_Estimated_FUEL_consump_POWER_model (l): {TOTAL_FUEL_01:,.3f}</h4>", unsafe_allow_html=True)
            #st.markdown(f"<h4 style='color:red;'>pmax_gen (kW): {p_gen_max:,.3f}</h4>", unsafe_allow_html=True)
            #st.markdown(f"<h4 style='color:red;'>P_max_GEN (kW): {P_MAX:,.3f}</h4>", unsafe_allow_html=True)
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
            add_kpi("Generator", "Total_Fuel_consumption (model estimation_SFC)", Total_Fuel_consumption, "l")
            add_kpi("Generator", "Total_Fuel_consumption (model estimation_POWER)", TOTAL_FUEL_01, "l")
            add_kpi("Generator", "Measured_Fuel_consumption (TRION_fuel_sensor)", measured_fuel, "l")
            #add_kpi("Generator", "Measured_Fuel_consumption (robust)", measured_fuel_robust, "l")


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
