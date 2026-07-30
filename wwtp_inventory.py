import os
import re
from pathlib import Path
import numpy as np
import pandas as pd
import statsmodels.formula.api as smf
import tkinter as tk
from tkinter import filedialog

# ==========================================================
# OVERVIEW
# This script implements the bottom-up hierarchical framework described in the
# Methods section of the paper, using China's above-ground municipal wastewater
# treatment plant (WWTP) dataset as the main case study.
#
# The workflow constructs component-specific operational intensity targets from
# local WWTP datasets, fits log-linear mixed-effects models with log(design
# capacity) as a fixed effect and technology-level partial pooling, and applies
# the fitted models to a national WWTP dataset to generate province-resolved and
# national inventory estimates for Scope 1, Scope 2, and an operational Scope 3
# proxy. Effluent standard class and climate zone are included when available
# and retained in the fitted specification.
#
# Uncertainty is propagated by Monte Carlo simulation (MC_N iterations), and
# covariate importance is assessed using a Sobol-like 2^4 full-factorial design
# (16 runs).
#
# INPUTS
# 1) LOCAL DATASETS (Excel; multiple files allowed)
#    Required fields include design capacity, treatment technology, influent
#    COD and TN, electricity-use intensity, operating expenditures, effluent
#    standard, and climate zone; location information is recommended.
#
# 2) TECHNOLOGY CROSSWALK (2 columns)
#    tech_nat -> tech_local, used to align national reporting categories with
#    the local technology taxonomy used for model fitting.
#
# 3) NATIONAL DATASET (Excel or CSV)
#    Required fields include plant type, design capacity, treatment technology,
#    effluent standard, climate zone, and location information.
#
# OUTPUTS
# One Excel workbook containing:
#   - National_Summary_MC
#   - Provincial_Summary_MC
#   - Tech_Prov_Weighted_MC
#   - Sobol_Y_16runs_kgCO2e_m3
#   - Sobol_FirstOrder_%
#   - Sobol_TotalOrder_%
# ==========================================================

# ==========================================================
# Stage 0. Setup
# ==========================================================
GWP_CH4, GWP_N2O = 28.0, 265.0

EF_GRID_MEAN = 0.54
EF_CH4_MEAN = 3.8
EF_N2O_MEAN = 13.9

EF_IO_CHEM_MEAN = 33.08
EF_IO_REPAIR_MEAN = 3.02
EF_IO_ADMIN_MEAN = 4.41

TECH6 = ["AAO", "AO", "Biofilm", "MBR", "OD", "SBR"]

MC_N = 1000
RANDOM_SEED = 42

CV_TECH_INTENSITY = 0.50
CV_EF_CH4_N2O = 0.50
CV_GRID = 0.20
CV_IO = 0.20

OUT_XLSX_NAME = "Inventory_outputs_PP_MC_Sobol.xlsx"

# ==========================================================
# Utility functions
# ==========================================================

def normalize(s: str) -> str:
    """Normalise header strings for robust column matching across files."""
    s = str(s).replace("\n", " ").replace("\r", " ")
    return re.sub(r"\s+", " ", s).strip()


def _pick_col(df, keys):
    """Return the first column whose normalised header matches any keyword in keys."""
    cols = list(df.columns)
    cols_norm = [normalize(c).lower() for c in cols]
    keys_norm = [k.lower() for k in keys]
    for c, cn in zip(cols, cols_norm):
        if any(k in cn for k in keys_norm):
            return c
    return None


def _norm_cat(x):
    """Standardise categorical labels; empty or missing values are mapped to 'Unknown'."""
    s = re.sub(r"[^0-9A-Za-z]+", "_", str(x)).strip("_")
    return s if s else "Unknown"


def _norm_effluent_5class(val: str) -> str:
    """
    Map heterogeneous effluent-standard strings to the 5-class taxonomy used in the paper:
    Above_1A, Class_1A, Class_1B, Class_2_or_below, Others, and Unknown.
    """
    if val is None or str(val).strip() == "":
        return "Unknown"

    s = str(val).lower().strip()

    if (
        ("above" in s and "1a" in s)
        or ("better" in s and "1a" in s)
        or ("优于" in s and "a" in s)
        or ("以上" in s and "1a" in s)
        or ("以上" in s and "一级a" in s)
        or ("1a以上" in s)
        or ("一级a以上" in s)
        or ("高于" in s and "1a" in s)
    ):
        return "Above_1A"

    if ("1a" in s) or ("ia" in s) or ("一级a" in s):
        return "Class_1A"

    if ("1b" in s) or ("ib" in s) or ("一级b" in s) or ("一級b" in s) or ("1 b" in s) or ("class 1b" in s):
        return "Class_1B"

    if ("class 2" in s) or re.search(r"\b2\b", s) or ("二级" in s) or ("二類" in s) or ("2类" in s) or ("2級" in s):
        return "Class_2_or_below"
    if ("class 3" in s) or re.search(r"\b3\b", s) or ("三级" in s) or ("3类" in s) or ("3級" in s):
        return "Class_2_or_below"
    if ("below" in s) or ("及以下" in s) or ("以下" in s):
        return "Class_2_or_below"

    return "Others"


def lognormal_params_from_mean_cv(mean, cv):
    """Convert mean and coefficient of variation to the underlying normal parameters of a lognormal."""
    sigma = np.sqrt(np.log(1.0 + cv**2))
    mu = np.log(mean) - 0.5 * sigma**2
    return mu, sigma


def sample_lognormal(mean, cv, size):
    """Draw lognormal samples with specified mean and coefficient of variation."""
    mu, sigma = lognormal_params_from_mean_cv(mean, cv)
    return np.random.lognormal(mean=mu, sigma=sigma, size=size)


def sample_lognormal_multiplier(cv, size):
    """Draw mean-one lognormal multipliers for plant-level heterogeneity."""
    sigma = np.sqrt(np.log(1.0 + cv**2))
    mu = -0.5 * sigma**2
    return np.random.lognormal(mean=mu, sigma=sigma, size=size)


def sample_trunc_normal(mean, cv, size):
    """Draw normal samples with SD = mean × CV and truncate at zero."""
    sd = mean * cv
    x = np.random.normal(loc=mean, scale=sd, size=size)
    x[x < 0] = 0.0
    return x


def total_intensity_tCO2e_m3(ch4_kg_m3, n2o_kg_m3, s2_kgco2_m3, s3_kgco2_m3):
    """Return total intensity in tCO2e/m3."""
    return (ch4_kg_m3 * GWP_CH4 + n2o_kg_m3 * GWP_N2O + s2_kgco2_m3 + s3_kgco2_m3) / 1000.0


def total_intensity_kgCO2e_m3(ch4_kg_m3, n2o_kg_m3, s2_kgco2_m3, s3_kgco2_m3):
    """Return total intensity in kgCO2e/m3."""
    return ch4_kg_m3 * GWP_CH4 + n2o_kg_m3 * GWP_N2O + s2_kgco2_m3 + s3_kgco2_m3


def prov_flow_weighted_intensity_kg(df, I_kgCO2e_m3, w="Q_m3_y"):
    """Return province-level flow-weighted mean total intensity in kgCO2e/m3."""
    tmp = df[["province", w]].copy()
    tmp["_I"] = I_kgCO2e_m3.values if hasattr(I_kgCO2e_m3, "values") else I_kgCO2e_m3
    return tmp.groupby("province").apply(lambda g: (g[w] * g["_I"]).sum() / g[w].sum(), include_groups=False)


def read_table(path: str) -> pd.DataFrame:
    path = str(path)
    if path.lower().endswith((".xlsx", ".xls")):
        return pd.read_excel(path)
    if path.lower().endswith(".csv"):
        return pd.read_csv(path)
    raise ValueError(f"Unsupported file type: {path}")


def pick_files():
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    root.update()

    print("Stage 1a: Select LOCAL WWTP datasets (multiple files).")
    local_paths = filedialog.askopenfilenames(
        title="Step 1 of 4: Select LOCAL WWTP datasets (multiple files)",
        filetypes=[("Excel/CSV files", "*.xlsx *.xls *.csv")],
    )
    if not local_paths:
        raise RuntimeError("No LOCAL datasets selected.")

    print("Stage 1b: Select TECHNOLOGY CROSSWALK (single file).")
    crosswalk_path = filedialog.askopenfilename(
        title="Step 2 of 4: Select TECHNOLOGY CROSSWALK (single file)",
        filetypes=[("Excel/CSV files", "*.xlsx *.xls *.csv")],
    )
    if not crosswalk_path:
        raise RuntimeError("No TECHNOLOGY CROSSWALK selected.")

    print("Stage 1c: Select NATIONAL dataset (single file).")
    national_path = filedialog.askopenfilename(
        title="Step 3 of 4: Select NATIONAL dataset (single file)",
        filetypes=[("Excel/CSV files", "*.xlsx *.xls *.csv")],
    )
    if not national_path:
        raise RuntimeError("No NATIONAL dataset selected.")

    print("Stage 1d: Select output folder.")
    output_dir = filedialog.askdirectory(title="Step 4 of 4: Select output folder")
    if not output_dir:
        output_dir = str(Path.cwd() / "outputs")

    root.destroy()
    return list(local_paths), str(crosswalk_path), str(national_path), str(output_dir)


# ==========================================================
# Stage 1. Data inputs and harmonisation
# ==========================================================
local_paths, crosswalk_path, national_path, OUTPUT_DIR = pick_files()
os.makedirs(OUTPUT_DIR, exist_ok=True)

local_raw_frames = [read_table(p) for p in local_paths]
print(f"Loaded {len(local_raw_frames)} LOCAL datasets.")

print("Stage 1b: Read TECHNOLOGY CROSSWALK.")
cw = read_table(crosswalk_path)
cw = cw.iloc[:, :2].copy()
cw.columns = ["tech_nat_raw", "tech_local_raw"]
cw["tech_nat"] = cw["tech_nat_raw"].astype(str).str.strip().map(_norm_cat)
cw["tech_local"] = cw["tech_local_raw"].astype(str).str.strip().map(_norm_cat)
cw_map_nat_to_local = dict(zip(cw["tech_nat"], cw["tech_local"]))

print("Stage 1c: Read NATIONAL dataset.")
nat = read_table(national_path)

col_type = _pick_col(nat, ["Type", "Class", "厂站类型", "Type_WWTP"])
col_scale = _pick_col(nat, ["Scale", "Capacity", "规模", "Scale_WWTP"])
col_proc = _pick_col(nat, ["Process", "工艺", "Process_category"])
col_prov = _pick_col(nat, ["Local_province", "Province", "省"])
col_eff = _pick_col(nat, ["Discharge_standard_WWTP", "Effluent", "出水标准", "排放标准", "class"])
col_clim = _pick_col(nat, ["Climate zone", "climate zone", "ClimateZone", "Climate", "Koppen", "气候", "climate_zone"])

if col_type is None or col_scale is None or col_proc is None:
    raise ValueError("NATIONAL dataset is missing required Type, Scale, or Process columns.")
if col_clim is None:
    raise ValueError("NATIONAL dataset is missing the Climate zone column.")

nat["type_norm"] = nat[col_type].astype(str).str.lower().str.replace(r"[^a-z]", "", regex=True)
nat_A = nat[nat["type_norm"].str.contains("a")].copy()

# National capacity unit assumed to be 10^4 m3/d
nat_A["cap_thousand_m3_d"] = pd.to_numeric(nat_A[col_scale], errors="coerce") * 10.0
nat_A = nat_A[nat_A["cap_thousand_m3_d"] > 0].copy()

nat_A["province"] = nat_A[col_prov].astype(str).str.strip() if col_prov else "Unknown"
nat_A["province"] = nat_A["province"].map(_norm_cat)

TECH6_norm = [_norm_cat(x) for x in TECH6]
nat_A["tech_nat"] = nat_A[col_proc].astype(str).str.strip().map(_norm_cat)
nat_A["tech_nat"] = nat_A["tech_nat"].where(nat_A["tech_nat"].isin(TECH6_norm), other=_norm_cat("Others"))
nat_A["tech_local"] = nat_A["tech_nat"].map(cw_map_nat_to_local).fillna(_norm_cat("Unknown"))

nat_A["effluent_class"] = nat_A[col_eff].map(_norm_effluent_5class) if col_eff else "Unknown"
nat_A["effluent_class"] = nat_A["effluent_class"].map(_norm_cat)
nat_A["climate_zone"] = nat_A[col_clim].astype(str).map(_norm_cat)

nat_A["log_cap"] = np.log(nat_A["cap_thousand_m3_d"] + 1e-3)
nat_A["Q_m3_y"] = nat_A["cap_thousand_m3_d"] * 1000.0 * 365.0

# ==========================================================
# Stage 2. Target construction and model fitting
# ==========================================================

def clean_local_df(raw_df):
    """Harmonise one local WWTP dataset and construct component-specific operational intensity targets."""
    df = raw_df.loc[:, ~raw_df.columns.duplicated()].copy()

    col_tech = _pick_col(df, ["Technology", "Process", "工艺", "tech", "local tech"])
    col_cap = _pick_col(df, ["Design Capacity", "Capacity", "规模", "Scale"])
    col_cod = _pick_col(df, ["COD (mg/L)", "COD", "CODcr", "Influent COD"])
    col_tn = _pick_col(df, ["TN (mg/L)", "TN", "Influent TN", "Ammonia"])
    col_e = _pick_col(df, [
        "Electricity Requirement",
        "Electricity (kWh/m3)",
        "Electricity use intensity",
        "electricity use intensity",
        "electricity use",
        "kwh/m3",
        "kwh m-3",
        "电耗",
    ])
    col_eff = _pick_col(df, ["Effluent standard", "Effluent", "出水标准", "排放标准", "discharge", "Discharge standard"])
    col_clim = _pick_col(df, ["Climate zone", "climate zone", "Climate", "Koppen", "气候", "climate_zone"])
    col_prov = _pick_col(df, ["Province", "省", "Local_province", "行政省份"])

    data = pd.DataFrame()
    data["tech_local"] = df[col_tech].astype(str).str.strip() if col_tech else "Unknown"

    data["cap_thousand_m3_d"] = pd.to_numeric(df[col_cap], errors="coerce").fillna(0.0) if col_cap else 0.0
    data["CODin_mg_L"] = pd.to_numeric(df[col_cod], errors="coerce").fillna(0.0) if col_cod else 0.0
    data["TNin_mg_L"] = pd.to_numeric(df[col_tn], errors="coerce").fillna(0.0) if col_tn else 0.0
    data["e_kwh_per_m3"] = pd.to_numeric(df[col_e], errors="coerce").fillna(0.0) if col_e else 0.0

    col_chem = _pick_col(df, ["Chemical Cost", "Chem cost", "化学品费", "Chemicals"])
    col_rep = _pick_col(df, ["Repair and maintenance", "Repair Cost", "维修", "Maintenance"])
    col_admin = _pick_col(df, ["Administrative Cost", "Admin cost", "管理", "Administration"])

    data["chem_cost_rmb_per_km3"] = pd.to_numeric(df[col_chem], errors="coerce").fillna(0.0) if col_chem else 0.0
    data["repair_cost_rmb_per_km3"] = pd.to_numeric(df[col_rep], errors="coerce").fillna(0.0) if col_rep else 0.0
    data["admin_cost_rmb_per_km3"] = pd.to_numeric(df[col_admin], errors="coerce").fillna(0.0) if col_admin else 0.0

    data["province"] = df[col_prov].astype(str).str.strip() if col_prov else "Unknown"
    data["effluent_class"] = df[col_eff].map(_norm_effluent_5class) if col_eff else "Unknown"
    data["climate_zone"] = df[col_clim].astype(str).map(_norm_cat) if col_clim else "Unknown"

    data = data[data["cap_thousand_m3_d"] > 0].copy()

    data["int_ch4_kg_m3"] = (data["CODin_mg_L"] * 1e-3) * (EF_CH4_MEAN / 1000.0)
    data["int_n2o_kg_m3"] = (data["TNin_mg_L"] * 1e-3) * (EF_N2O_MEAN / 1000.0)
    data["int_s2_kgCO2e_m3"] = data["e_kwh_per_m3"] * EF_GRID_MEAN

    term_t_per_1000m3 = (
        data["chem_cost_rmb_per_km3"] * EF_IO_CHEM_MEAN
        + data["repair_cost_rmb_per_km3"] * EF_IO_REPAIR_MEAN
        + data["admin_cost_rmb_per_km3"] * EF_IO_ADMIN_MEAN
    ) * 1e-6

    data["int_s3_kgCO2e_m3"] = term_t_per_1000m3
    data["log_cap"] = np.log(data["cap_thousand_m3_d"] + 1e-3)

    for c in ["tech_local", "effluent_class", "climate_zone", "province"]:
        data[c] = data[c].map(_norm_cat)

    return data


local_frames = [clean_local_df(df) for df in local_raw_frames]
local_all = pd.concat(local_frames, ignore_index=True)
for c in ["tech_local", "effluent_class", "climate_zone", "province"]:
    local_all[c] = local_all[c].map(_norm_cat)


def fit_pp_keep(train_df, target_col, group_col="tech_local", eff_col="effluent_class", clim_col="climate_zone"):
    """
    Fit one log-linear component model for a local intensity target.

    Model structure:
      - response: log(target)
      - fixed effects: log(capacity) + optional effluent + optional climate
      - random intercept: technology group (partial pooling)

    Fallback rules:
      - if the training subset is too small, return a mean-only model
      - if MixedLM fails, fall back to OLS
    """
    td = train_df[train_df[target_col] > 1e-12].copy()

    if len(td) < 6:
        gm = td[target_col].mean() if len(td) > 0 else 0.0
        return {
            "kind": "mean_only",
            "mean": gm,
            "has_eff": False,
            "has_clim": False,
            "eff_col": eff_col,
            "clim_col": clim_col,
            "group_col": group_col,
        }

    has_eff = td[eff_col].nunique(dropna=True) > 1
    has_clim = td[clim_col].nunique(dropna=True) > 1

    rhs = ["log_cap"]
    if has_eff:
        rhs.append(f"C({eff_col})")
    if has_clim:
        rhs.append(f"C({clim_col})")

    formula = "log_target ~ " + " + ".join(rhs)
    td["log_target"] = np.log(td[target_col].clip(lower=1e-12))

    try:
        res = smf.mixedlm(formula, td, groups=td[group_col]).fit()
        params = res.params.to_dict()
        re_dict = {k: float(v.iloc[0] if hasattr(v, 'iloc') else v[0]) for k, v in res.random_effects.items()}
        return {
            "kind": "mixedlm",
            "formula": formula,
            "params": params,
            "re_dict": re_dict,
            "has_eff": has_eff,
            "has_clim": has_clim,
            "eff_col": eff_col,
            "clim_col": clim_col,
            "group_col": group_col,
        }
    except Exception:
        from statsmodels.formula.api import ols
        ols_res = ols(formula, data=td).fit()
        params = ols_res.params.to_dict()
        return {
            "kind": "ols",
            "formula": formula + " (OLS fallback)",
            "params": params,
            "re_dict": {},
            "has_eff": has_eff,
            "has_clim": has_clim,
            "eff_col": eff_col,
            "clim_col": clim_col,
            "group_col": group_col,
        }


# Fit one mixed-effects model for each component-specific local intensity target
targets = ["int_ch4_kg_m3", "int_n2o_kg_m3", "int_s2_kgCO2e_m3", "int_s3_kgCO2e_m3"]
models = {t: fit_pp_keep(local_all, t, group_col="tech_local") for t in targets}

# ==========================================================
# Stage 3. National extrapolation
# ==========================================================

def predict_pp(modelinfo, df, *, use_cap=True, use_climate=True, use_effluent=True, use_tech=True,
               mean_log_cap=None, base_eff=None, base_clim=None):
    """
    Predict plant-level component intensity from the fitted model.

    Prediction rules:
      - if a covariate term was not retained in fitting, prediction uses the
        reduced specification with no contribution from that term
      - if a retained categorical term encounters an unseen or omitted level,
        its contribution is zero under reference-category coding
      - optional baseline arguments are defined here for reusable prediction
        and are reused later in the Sobol-like factor-switching design
      - if technology pooling is disabled or unavailable, prediction uses fixed
        effects only
    """
    if modelinfo["kind"] == "mean_only":
        return pd.Series([modelinfo["mean"]] * len(df), index=df.index)

    params = modelinfo["params"]
    re_dict = modelinfo.get("re_dict", {})
    eff_col = modelinfo["eff_col"]
    clim_col = modelinfo["clim_col"]
    group_col = modelinfo["group_col"]
    has_eff = modelinfo.get("has_eff", False)
    has_clim = modelinfo.get("has_clim", False)

    if mean_log_cap is None:
        mean_log_cap = float(df["log_cap"].mean())
    if base_eff is None:
        base_eff = df[eff_col].value_counts().idxmax() if eff_col in df.columns else "Unknown"
    if base_clim is None:
        base_clim = df[clim_col].value_counts().idxmax() if clim_col in df.columns else "Unknown"

    log_cap_series = df["log_cap"].astype(float) if use_cap else pd.Series([float(mean_log_cap)] * len(df), index=df.index)
    eff_series = df[eff_col].astype(str) if (use_effluent and has_eff) else pd.Series([str(base_eff)] * len(df), index=df.index)
    clim_series = df[clim_col].astype(str) if (use_climate and has_clim) else pd.Series([str(base_clim)] * len(df), index=df.index)
    tech_series = df[group_col].astype(str)

    out = np.empty(len(df), dtype=float)
    for i in range(len(df)):
        v = float(params.get("Intercept", 0.0)) + float(params.get("log_cap", 0.0)) * float(log_cap_series.iloc[i])
        if has_eff:
            v += float(params.get(f"C({eff_col})[T.{eff_series.iloc[i]}]", 0.0))
        if has_clim:
            v += float(params.get(f"C({clim_col})[T.{clim_series.iloc[i]}]", 0.0))
        if use_tech and modelinfo["kind"] == "mixedlm":
            v += float(re_dict.get(tech_series.iloc[i], 0.0))
        out[i] = np.exp(v)

    return pd.Series(out, index=df.index)


pred_ch4 = predict_pp(models["int_ch4_kg_m3"], nat_A).values
pred_n2o = predict_pp(models["int_n2o_kg_m3"], nat_A).values
pred_s2 = predict_pp(models["int_s2_kgCO2e_m3"], nat_A).values
pred_s3 = predict_pp(models["int_s3_kgCO2e_m3"], nat_A).values

# ==========================================================
# Stage 4. Inventory aggregation and outputs
# ==========================================================
Q = nat_A["Q_m3_y"].values
prov = nat_A["province"].values
tech_nat = nat_A["tech_nat"].values
cap = nat_A["cap_thousand_m3_d"].values

prov_codes, prov_levels = pd.factorize(prov, sort=True)
techprov_codes, techprov_levels = pd.factorize(list(zip(prov, tech_nat)), sort=False)

nP = len(prov_levels)
nTP = len(techprov_levels)

prov_n_plants = np.bincount(prov_codes, minlength=nP)
prov_cap_sum = np.bincount(prov_codes, weights=cap, minlength=nP)
prov_Q_sum = np.bincount(prov_codes, weights=Q, minlength=nP)

tp_Q_sum = np.bincount(techprov_codes, weights=Q, minlength=nTP)
tp_n_plants = np.bincount(techprov_codes, minlength=nTP)

prov_CH4 = np.zeros((MC_N, nP))
prov_N2O = np.zeros((MC_N, nP))
prov_S1 = np.zeros((MC_N, nP))
prov_S2 = np.zeros((MC_N, nP))
prov_S3 = np.zeros((MC_N, nP))
prov_TOT = np.zeros((MC_N, nP))
prov_I_TOT = np.zeros((MC_N, nP))

tp_S1 = np.zeros((MC_N, nTP))
tp_S2 = np.zeros((MC_N, nTP))
tp_S3 = np.zeros((MC_N, nTP))
tp_TOT = np.zeros((MC_N, nTP))

# ==========================================================
# Stage 5. Uncertainty and sensitivity analysis
# ==========================================================
# Stage 5.1 Monte Carlo uncertainty analysis
np.random.seed(RANDOM_SEED)

samp_EF_CH4 = sample_lognormal(EF_CH4_MEAN, CV_EF_CH4_N2O, MC_N)
samp_EF_N2O = sample_lognormal(EF_N2O_MEAN, CV_EF_CH4_N2O, MC_N)
samp_GRID = sample_trunc_normal(EF_GRID_MEAN, CV_GRID, MC_N)

samp_IO_CHEM = sample_lognormal(EF_IO_CHEM_MEAN, CV_IO, MC_N)
samp_IO_REPAIR = sample_lognormal(EF_IO_REPAIR_MEAN, CV_IO, MC_N)
samp_IO_ADMIN = sample_lognormal(EF_IO_ADMIN_MEAN, CV_IO, MC_N)

mult_CH4 = samp_EF_CH4 / EF_CH4_MEAN
mult_N2O = samp_EF_N2O / EF_N2O_MEAN
mult_S2 = samp_GRID / EF_GRID_MEAN
IO_sum_mean = EF_IO_CHEM_MEAN + EF_IO_REPAIR_MEAN + EF_IO_ADMIN_MEAN
mult_S3 = (samp_IO_CHEM + samp_IO_REPAIR + samp_IO_ADMIN) / IO_sum_mean

for k in range(MC_N):
    m_ch4 = sample_lognormal_multiplier(CV_TECH_INTENSITY, len(Q))
    m_n2o = sample_lognormal_multiplier(CV_TECH_INTENSITY, len(Q))
    m_s2 = sample_lognormal_multiplier(CV_TECH_INTENSITY, len(Q))
    m_s3 = sample_lognormal_multiplier(CV_TECH_INTENSITY, len(Q))

    ch4_k = pred_ch4 * mult_CH4[k] * m_ch4
    n2o_k = pred_n2o * mult_N2O[k] * m_n2o
    s2_k = pred_s2 * mult_S2[k] * m_s2
    s3_k = pred_s3 * mult_S3[k] * m_s3

    E_CH4 = (ch4_k * Q * GWP_CH4) / 1000.0
    E_N2O = (n2o_k * Q * GWP_N2O) / 1000.0
    E_S1 = E_CH4 + E_N2O
    E_S2 = (s2_k * Q) / 1000.0
    E_S3 = (s3_k * Q) / 1000.0
    E_TOT = E_S1 + E_S2 + E_S3

    prov_CH4[k, :] = np.bincount(prov_codes, weights=E_CH4, minlength=nP)
    prov_N2O[k, :] = np.bincount(prov_codes, weights=E_N2O, minlength=nP)
    prov_S1[k, :] = np.bincount(prov_codes, weights=E_S1, minlength=nP)
    prov_S2[k, :] = np.bincount(prov_codes, weights=E_S2, minlength=nP)
    prov_S3[k, :] = np.bincount(prov_codes, weights=E_S3, minlength=nP)
    prov_TOT[k, :] = np.bincount(prov_codes, weights=E_TOT, minlength=nP)
    prov_I_TOT[k, :] = prov_TOT[k, :] / np.maximum(prov_Q_sum, 1e-30)

    tp_S1[k, :] = np.bincount(techprov_codes, weights=E_S1, minlength=nTP)
    tp_S2[k, :] = np.bincount(techprov_codes, weights=E_S2, minlength=nTP)
    tp_S3[k, :] = np.bincount(techprov_codes, weights=E_S3, minlength=nTP)
    tp_TOT[k, :] = np.bincount(techprov_codes, weights=E_TOT, minlength=nTP)


def mean_ci(arr_2d):
    """Return the mean and 2.5th/97.5th percentile interval along axis 0."""
    mean = arr_2d.mean(axis=0)
    lo = np.percentile(arr_2d, 2.5, axis=0)
    hi = np.percentile(arr_2d, 97.5, axis=0)
    return mean, lo, hi


def mean_ci_1d(arr):
    """Return the mean and 2.5th/97.5th percentile interval for a 1D array."""
    return float(arr.mean()), float(np.percentile(arr, 2.5)), float(np.percentile(arr, 97.5))


mCH4, loCH4, hiCH4 = mean_ci(prov_CH4)
mN2O, loN2O, hiN2O = mean_ci(prov_N2O)
mS1, loS1, hiS1 = mean_ci(prov_S1)
mS2, loS2, hiS2 = mean_ci(prov_S2)
mS3, loS3, hiS3 = mean_ci(prov_S3)
mTOT, loTOT, hiTOT = mean_ci(prov_TOT)
mI, loI, hiI = mean_ci(prov_I_TOT)

prov_summary = pd.DataFrame({
    "Province": prov_levels,
    "N_plants": prov_n_plants.astype(int),
    "Capacity_thousand_m3_d": prov_cap_sum,
    "CH4_MtCO2e_mean": mCH4 / 1e6, "CH4_MtCO2e_lo": loCH4 / 1e6, "CH4_MtCO2e_hi": hiCH4 / 1e6,
    "N2O_MtCO2e_mean": mN2O / 1e6, "N2O_MtCO2e_lo": loN2O / 1e6, "N2O_MtCO2e_hi": hiN2O / 1e6,
    "Scope1_MtCO2e_mean": mS1 / 1e6, "Scope1_MtCO2e_lo": loS1 / 1e6, "Scope1_MtCO2e_hi": hiS1 / 1e6,
    "Scope2_MtCO2e_mean": mS2 / 1e6, "Scope2_MtCO2e_lo": loS2 / 1e6, "Scope2_MtCO2e_hi": hiS2 / 1e6,
    "Scope3_MtCO2e_mean": mS3 / 1e6, "Scope3_MtCO2e_lo": loS3 / 1e6, "Scope3_MtCO2e_hi": hiS3 / 1e6,
    "Total_MtCO2e_mean": mTOT / 1e6, "Total_MtCO2e_lo": loTOT / 1e6, "Total_MtCO2e_hi": hiTOT / 1e6,
    "I_total_tCO2e_m3_mean": mI, "I_total_tCO2e_m3_lo": loI, "I_total_tCO2e_m3_hi": hiI,
    "I_total_kgCO2e_m3_mean": mI * 1000.0, "I_total_kgCO2e_m3_lo": loI * 1000.0, "I_total_kgCO2e_m3_hi": hiI * 1000.0,
}).sort_values("Total_MtCO2e_mean", ascending=False).reset_index(drop=True)

nat_CH4 = prov_CH4.sum(axis=1)
nat_N2O = prov_N2O.sum(axis=1)
nat_S1 = prov_S1.sum(axis=1)
nat_S2 = prov_S2.sum(axis=1)
nat_S3 = prov_S3.sum(axis=1)
nat_TOT = prov_TOT.sum(axis=1)

nCH4m, nCH4lo, nCH4hi = mean_ci_1d(nat_CH4)
nN2Om, nN2Olo, nN2Ohi = mean_ci_1d(nat_N2O)
nS1m, nS1lo, nS1hi = mean_ci_1d(nat_S1)
nS2m, nS2lo, nS2hi = mean_ci_1d(nat_S2)
nS3m, nS3lo, nS3hi = mean_ci_1d(nat_S3)
nTOTm, nTOTlo, nTOThi = mean_ci_1d(nat_TOT)

national_summary = pd.DataFrame({
    "Metric": [
        "Number of plants",
        "CH4 (MtCO2e/y)",
        "N2O (MtCO2e/y)",
        "Scope 1 (MtCO2e/y)",
        "Scope 2 (MtCO2e/y)",
        "Scope 3 (MtCO2e/y)",
        "Total (MtCO2e/y)",
    ],
    "Mean": [int(len(nat_A)), nCH4m / 1e6, nN2Om / 1e6, nS1m / 1e6, nS2m / 1e6, nS3m / 1e6, nTOTm / 1e6],
    "CI2.5%": [np.nan, nCH4lo / 1e6, nN2Olo / 1e6, nS1lo / 1e6, nS2lo / 1e6, nS3lo / 1e6, nTOTlo / 1e6],
    "CI97.5%": [np.nan, nCH4hi / 1e6, nN2Ohi / 1e6, nS1hi / 1e6, nS2hi / 1e6, nS3hi / 1e6, nTOThi / 1e6],
})

tp_mS1, tp_loS1, tp_hiS1 = mean_ci(tp_S1)
tp_mS2, tp_loS2, tp_hiS2 = mean_ci(tp_S2)
tp_mS3, tp_loS3, tp_hiS3 = mean_ci(tp_S3)
tp_mT, tp_loT, tp_hiT = mean_ci(tp_TOT)

tp_I1 = tp_S1 / np.maximum(tp_Q_sum[None, :], 1e-30)
tp_I2 = tp_S2 / np.maximum(tp_Q_sum[None, :], 1e-30)
tp_I3 = tp_S3 / np.maximum(tp_Q_sum[None, :], 1e-30)
tp_IT = tp_TOT / np.maximum(tp_Q_sum[None, :], 1e-30)

tp_mI1, tp_loI1, tp_hiI1 = mean_ci(tp_I1)
tp_mI2, tp_loI2, tp_hiI2 = mean_ci(tp_I2)
tp_mI3, tp_loI3, tp_hiI3 = mean_ci(tp_I3)
tp_mIT, tp_loIT, tp_hiIT = mean_ci(tp_IT)

tp_prov = [p for (p, t) in techprov_levels]
tp_tech = [t for (p, t) in techprov_levels]

tech_prov_weighted = pd.DataFrame({
    "Province": tp_prov,
    "Tech": tp_tech,
    "N_plants": tp_n_plants.astype(int),
    "Q_m3_y": tp_Q_sum,
    "Scope1_MtCO2e_mean": tp_mS1 / 1e6, "Scope1_MtCO2e_lo": tp_loS1 / 1e6, "Scope1_MtCO2e_hi": tp_hiS1 / 1e6,
    "Scope2_MtCO2e_mean": tp_mS2 / 1e6, "Scope2_MtCO2e_lo": tp_loS2 / 1e6, "Scope2_MtCO2e_hi": tp_hiS2 / 1e6,
    "Scope3_MtCO2e_mean": tp_mS3 / 1e6, "Scope3_MtCO2e_lo": tp_loS3 / 1e6, "Scope3_MtCO2e_hi": tp_hiS3 / 1e6,
    "Total_MtCO2e_mean": tp_mT / 1e6, "Total_MtCO2e_lo": tp_loT / 1e6, "Total_MtCO2e_hi": tp_hiT / 1e6,
    "I_scope1_tCO2e_m3_mean": tp_mI1, "I_scope1_tCO2e_m3_lo": tp_loI1, "I_scope1_tCO2e_m3_hi": tp_hiI1,
    "I_scope2_tCO2e_m3_mean": tp_mI2, "I_scope2_tCO2e_m3_lo": tp_loI2, "I_scope2_tCO2e_m3_hi": tp_hiI2,
    "I_scope3_tCO2e_m3_mean": tp_mI3, "I_scope3_tCO2e_m3_lo": tp_loI3, "I_scope3_tCO2e_m3_hi": tp_hiI3,
    "I_total_tCO2e_m3_mean": tp_mIT, "I_total_tCO2e_m3_lo": tp_loIT, "I_total_tCO2e_m3_hi": tp_hiIT,
})

tech_order = [_norm_cat(x) for x in (TECH6 + ["Others"])]
prov_order = sorted(list(set(prov_levels)))
full_grid = pd.MultiIndex.from_product([prov_order, tech_order], names=["Province", "Tech"]).to_frame(index=False)

tech_prov_weighted = full_grid.merge(tech_prov_weighted, on=["Province", "Tech"], how="left")
for c in [c for c in tech_prov_weighted.columns if c not in ["Province", "Tech"]]:
    tech_prov_weighted[c] = tech_prov_weighted[c].fillna(0.0)

tech_prov_weighted["Tech"] = pd.Categorical(tech_prov_weighted["Tech"], categories=tech_order, ordered=True)
tech_prov_weighted = tech_prov_weighted.sort_values(["Province", "Tech"]).reset_index(drop=True)

# Stage 5.2 Sobol-like sensitivity analysis
# Actual baseline values used when factors are switched off in the Sobol-like design
mean_logcap_nat = float(nat_A["log_cap"].mean())
BASE_EFF = local_all["effluent_class"].value_counts().idxmax()
BASE_CLIM = local_all["climate_zone"].value_counts().idxmax()


def flags_from_mask(mask: int):
    """Map a 0-15 mask to factor switches (1=capacity, 2=climate, 4=effluent, 8=technology)."""
    return dict(
        use_cap=bool(mask & 1),
        use_climate=bool(mask & 2),
        use_effluent=bool(mask & 4),
        use_tech=bool(mask & 8),
    )


def get_prov_intensity_kg(mask: int) -> pd.Series:
    flags = flags_from_mask(mask)
    ch4 = predict_pp(models["int_ch4_kg_m3"], nat_A, mean_log_cap=mean_logcap_nat, base_eff=BASE_EFF, base_clim=BASE_CLIM, **flags)
    n2o = predict_pp(models["int_n2o_kg_m3"], nat_A, mean_log_cap=mean_logcap_nat, base_eff=BASE_EFF, base_clim=BASE_CLIM, **flags)
    s2 = predict_pp(models["int_s2_kgCO2e_m3"], nat_A, mean_log_cap=mean_logcap_nat, base_eff=BASE_EFF, base_clim=BASE_CLIM, **flags)
    s3 = predict_pp(models["int_s3_kgCO2e_m3"], nat_A, mean_log_cap=mean_logcap_nat, base_eff=BASE_EFF, base_clim=BASE_CLIM, **flags)
    I_kg = total_intensity_kgCO2e_m3(ch4, n2o, s2, s3)
    return prov_flow_weighted_intensity_kg(nat_A, I_kg, w="Q_m3_y")


provinces = sorted(nat_A["province"].unique())
Y16 = pd.DataFrame(index=provinces)
for mask in range(16):
    Y16[f"m{mask:02d}"] = get_prov_intensity_kg(mask).reindex(provinces).values

masks = np.arange(16)
A = np.where((masks & 1) > 0, 1, -1)
B = np.where((masks & 2) > 0, 1, -1)
C = np.where((masks & 4) > 0, 1, -1)
D = np.where((masks & 8) > 0, 1, -1)

terms = {
    "Capacity": A,
    "Climate": B,
    "Effluent": C,
    "Tech": D,
    "Cap:Clim": A * B,
    "Cap:Eff": A * C,
    "Cap:Tech": A * D,
    "Clim:Eff": B * C,
    "Clim:Tech": B * D,
    "Eff:Tech": C * D,
    "Cap:Clim:Eff": A * B * C,
    "Cap:Clim:Tech": A * B * D,
    "Cap:Eff:Tech": A * C * D,
    "Clim:Eff:Tech": B * C * D,
    "Cap:Clim:Eff:Tech": A * B * C * D,
}

factors = ["Capacity", "Climate", "Effluent", "Tech"]
first_df = pd.DataFrame(index=provinces, columns=factors, dtype=float)
total_df = pd.DataFrame(index=provinces, columns=factors, dtype=float)

for prov_name in provinces:
    y = Y16.loc[prov_name].values.astype(float)
    yc = y - y.mean()

    coef = {k: float(np.mean(yc * v)) for k, v in terms.items()}
    var_comp = {k: coef[k] ** 2 for k in coef}
    V = sum(var_comp.values())

    if V <= 1e-18:
        first_df.loc[prov_name, :] = 0.0
        total_df.loc[prov_name, :] = 0.0
        continue

    for f in factors:
        first_df.loc[prov_name, f] = 100.0 * var_comp[f] / V

    total_df.loc[prov_name, "Capacity"] = 100.0 * (
        var_comp["Capacity"] + var_comp["Cap:Clim"] + var_comp["Cap:Eff"] + var_comp["Cap:Tech"]
        + var_comp["Cap:Clim:Eff"] + var_comp["Cap:Clim:Tech"] + var_comp["Cap:Eff:Tech"] + var_comp["Cap:Clim:Eff:Tech"]
    ) / V
    total_df.loc[prov_name, "Climate"] = 100.0 * (
        var_comp["Climate"] + var_comp["Cap:Clim"] + var_comp["Clim:Eff"] + var_comp["Clim:Tech"]
        + var_comp["Cap:Clim:Eff"] + var_comp["Cap:Clim:Tech"] + var_comp["Clim:Eff:Tech"] + var_comp["Cap:Clim:Eff:Tech"]
    ) / V
    total_df.loc[prov_name, "Effluent"] = 100.0 * (
        var_comp["Effluent"] + var_comp["Cap:Eff"] + var_comp["Clim:Eff"] + var_comp["Eff:Tech"]
        + var_comp["Cap:Clim:Eff"] + var_comp["Cap:Eff:Tech"] + var_comp["Clim:Eff:Tech"] + var_comp["Cap:Clim:Eff:Tech"]
    ) / V
    total_df.loc[prov_name, "Tech"] = 100.0 * (
        var_comp["Tech"] + var_comp["Cap:Tech"] + var_comp["Clim:Tech"] + var_comp["Eff:Tech"]
        + var_comp["Cap:Clim:Tech"] + var_comp["Cap:Eff:Tech"] + var_comp["Clim:Eff:Tech"] + var_comp["Cap:Clim:Eff:Tech"]
    ) / V

# ==========================================================
# Final export
# ==========================================================
out_xlsx = os.path.join(OUTPUT_DIR, OUT_XLSX_NAME)

with pd.ExcelWriter(out_xlsx, engine="openpyxl") as w:
    national_summary.to_excel(w, sheet_name="National_Summary_MC", index=False)
    prov_summary.to_excel(w, sheet_name="Provincial_Summary_MC", index=False)
    tech_prov_weighted.to_excel(w, sheet_name="Tech_Prov_Weighted_MC", index=False)
    Y16.reset_index().rename(columns={"index": "Province"}).to_excel(w, sheet_name="Sobol_Y_16runs_kgCO2e_m3", index=False)
    first_df.reset_index().rename(columns={"index": "Province"}).to_excel(w, sheet_name="Sobol_FirstOrder_%", index=False)
    total_df.reset_index().rename(columns={"index": "Province"}).to_excel(w, sheet_name="Sobol_TotalOrder_%", index=False)

print(f"Saved workbook: {out_xlsx}")
