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
# local WWTP datasets, fits log-linear mixed-effects models with log(plant capacity)
# as a fixed effect and technology-level partial pooling, and applies the fitted
# models to a national WWTP dataset to generate province-resolved and national
# inventory estimates for Scope 1, Scope 2, and a partial upstream operational-
# input proxy. Effluent standard class and climate zone are included when
# available and retained in the fitted specification.
#
# Uncertainty is propagated by Monte Carlo simulation (MC_N iterations).
# Dependence of provincial intensity predictions on predictor availability is
# evaluated using a 2^4 full-factorial predictor-ablation analysis (16 refitted
# inclusion/exclusion model specifications).
#
# INPUTS
# 1) LOCAL DATASETS (Excel; multiple files allowed)
#    Required for each plant: plant capacity (10^3 m^3/d), treatment technology.
#    Component-specific data:
#      CH4/N2O: COD/TN removal or influent-effluent concentrations (mg/L).
#      Influent-only records require a configured removal fraction.
#      Electricity: electricity-use intensity (kWh/m^3).
#      Upstream proxy: selected operating expenditures (RMB/10^3 m^3).
#    Recommended: effluent standard, climate zone and location.
#
# 2) TECHNOLOGY CROSSWALK (2 columns)
#    Align local technology labels with the national reporting categories.
#
# 3) NATIONAL DATASET (Excel or CSV)
#    Required: plant type, design capacity (10^4 m^3/d), treatment technology
#    and province. Recommended: effluent standard and climate zone.
#
# OUTPUTS
# One Excel workbook containing:
#   - National_Summary_MC
#   - Provincial_Summary_MC
#   - Tech_Prov_Weighted_MC
#   - Factorial_Y_16runs
#   - Factorial_MainEffect_pct
#   - Factorial_TotalContrib_pct
#   - Factorial_Model_Specs
# ==========================================================

# ==========================================================
# Stage 0. Setup
# ==========================================================
GWP_CH4, GWP_N2O = 28.0, 265.0

EF_CH4_MEAN = 3.8
EF_N2O_MEAN = 13.9

EF_IO_CHEM_MEAN = 33.08
EF_IO_REPAIR_MEAN = 3.02
EF_IO_ADMIN_MEAN = 4.41

# 2015 regional electricity factors (kg CO2/kWh)
GRID_EF_2015 = {
    "North": 0.844,
    "Northeast": 0.909,
    "East": 0.648,
    "Central": 0.443,
    "Northwest": 0.663,
    "South": 0.414,
    "Northwest_proxy": 0.663,
}

# Province-to-grid assignment.
PROVINCE_GRID_REGION_2015 = {
    "Beijing": "North",
    "Tianjin": "North",
    "Hebei": "North",
    "Shanxi": "North",
    "Inner_Mongolia": "North",
    "Liaoning": "Northeast",
    "Jilin": "Northeast",
    "Heilongjiang": "Northeast",
    "Shanghai": "East",
    "Jiangsu": "East",
    "Zhejiang": "East",
    "Anhui": "East",
    "Fujian": "East",
    "Jiangxi": "Central",
    "Shandong": "North",
    "Henan": "Central",
    "Hubei": "Central",
    "Hunan": "Central",
    "Guangdong": "South",
    "Guangxi": "South",
    "Hainan": "South",
    "Chongqing": "Central",
    "Sichuan": "Central",
    "Guizhou": "South",
    "Yunnan": "South",
    "Tibet": "Northwest_proxy",
    "Shaanxi": "Northwest",
    "Gansu": "Northwest",
    "Qinghai": "Northwest",
    "Ningxia": "Northwest",
    "Xinjiang": "Northwest",
}

# 2015 province-level capacity-utilisation factors.
PROVINCE_UTILISATION_2015 = {
    "Beijing": 0.8736475504991874,
    "Tianjin": 0.8242871441765107,
    "Hebei": 0.779058678613647,
    "Shanxi": 0.7330078815598873,
    "Inner_Mongolia": 0.7108167485138278,
    "Liaoning": 0.7671128453156463,
    "Jilin": 0.6433623910336239,
    "Heilongjiang": 0.6651130511589026,
    "Shanghai": 0.7415548381467586,
    "Jiangsu": 0.7790644774877185,
    "Zhejiang": 0.7938619759488177,
    "Anhui": 0.8603200079628731,
    "Fujian": 0.8089552470092356,
    "Jiangxi": 0.8281947795804963,
    "Shandong": 0.7966178535551535,
    "Henan": 0.7721146741225148,
    "Hubei": 0.8577017938793753,
    "Hunan": 0.8339400120749585,
    "Guangdong": 0.8744454027704546,
    "Guangxi": 0.7980242186723978,
    "Hainan": 0.6980728051391863,
    "Chongqing": 0.9236004625643797,
    "Sichuan": 0.844265312738682,
    "Guizhou": 0.8886225266362253,
    "Yunnan": 0.8955237866551096,
    "Tibet": 0.9174379859311366,
    "Shaanxi": 0.724790233346263,
    "Gansu": 0.6645040774514924,
    "Qinghai": 0.7286101938732141,
    "Ningxia": 0.8657534246575342,
    "Xinjiang": 0.6781316768450651,
}

# Removal fractions for datasets without paired effluent measurements.
REMOVAL_FRACTION_PROXY = {
    "Shenzhen": {"COD": 0.850581395349, "TN": 0.606694560669},
    "Xinjiang": {"COD": 0.691064419303, "TN": 0.612132352941},
}

TECH6 = ["AAO", "AO", "Biofilm", "MBR", "OD", "SBR"]

MC_N = 1000
RANDOM_SEED = 42

# Monte Carlo coefficients of variation.
CV_EF_CH4_N2O = 0.50
CV_GRID = 0.03
CV_IO = 0.20
CV_REMOVAL_COD = 0.05
CV_REMOVAL_TN = 0.08
CV_UTILISATION = 0.10

OUT_XLSX_NAME = "Inventory_outputs.xlsx"

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
    if pd.isna(x):
        return "Unknown"
    s = re.sub(r"[^0-9A-Za-z]+", "_", str(x)).strip("_")
    return s if s else "Unknown"


# Canonical province names.
_PROVINCE_CANONICAL = {
    re.sub(r"[^A-Za-z]+", "", k).lower(): k
    for k in PROVINCE_UTILISATION_2015.keys()
}

# Pinyin-style province names used in the national source file.
_PROVINCE_ALIASES = {
    "neimenggu": "Inner_Mongolia",
    "xizang": "Tibet",
}

def _norm_province(x):
    """Align province names with the parameter lookups."""
    compact = re.sub(r"[^A-Za-z]+", "", str(x)).lower()
    if compact in _PROVINCE_ALIASES:
        return _PROVINCE_ALIASES[compact]
    return _PROVINCE_CANONICAL.get(compact, _norm_cat(x))


def _norm_effluent_5class(val: str) -> str:
    """Align effluent classes; use Unknown for missing values."""
    if pd.isna(val) or str(val).strip() == "":
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


def _removal_proxy_key(source_name):
    """Match a dataset identifier to a configured removal-fraction entry."""
    matches = [key for key in REMOVAL_FRACTION_PROXY
               if key.casefold() in str(source_name).casefold()]
    if len(matches) > 1:
        raise ValueError(f"Ambiguous removal-fraction configuration for {source_name}: {matches}")
    return matches[0] if matches else None


def lognormal_params_from_mean_cv(mean, cv):
    """Convert mean and coefficient of variation to the underlying normal parameters of a lognormal."""
    sigma = np.sqrt(np.log(1.0 + cv**2))
    mu = np.log(mean) - 0.5 * sigma**2
    return mu, sigma


def sample_lognormal(mean, cv, size):
    """Draw lognormal samples with specified mean and coefficient of variation."""
    mu, sigma = lognormal_params_from_mean_cv(mean, cv)
    return np.random.lognormal(mean=mu, sigma=sigma, size=size)


def sample_trunc_normal(mean, cv, size):
    """Draw normal samples with SD = mean * CV and clip negative values to zero."""
    sd = mean * cv
    x = np.random.normal(loc=mean, scale=sd, size=size)
    x[x < 0] = 0.0
    return x


def sample_bounded_normal(mean, cv, size, lower=0.0, upper=1.0):
    """Draw normal samples with SD = mean * CV and clip to [lower, upper]."""
    sd = mean * cv
    x = np.random.normal(loc=mean, scale=sd, size=size)
    return np.clip(x, lower, upper)


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
local_source_names = [Path(p).stem for p in local_paths]
print(f"Loaded {len(local_raw_frames)} LOCAL datasets.")

print("Stage 1b: Read TECHNOLOGY CROSSWALK.")
cw = read_table(crosswalk_path)
cw = cw.iloc[:, :2].copy()
cw.columns = ["tech_nat_raw", "tech_local_raw"]

# Map local technology labels to national categories.
local_alias_to_target = {}

for _, row in cw.iterrows():
    if pd.isna(row["tech_nat_raw"]):
        continue

    target = _norm_cat(row["tech_nat_raw"])
    if target in {"Unknown", ""}:
        continue

    aliases = [str(row["tech_nat_raw"]).strip()]

    if not pd.isna(row["tech_local_raw"]):
        aliases.extend(
            a.strip()
            for a in re.split(r"[,;|]+", str(row["tech_local_raw"]))
            if a.strip()
        )

    for alias in aliases:
        local_alias_to_target[_norm_cat(alias)] = target


print("Stage 1c: Read NATIONAL dataset.")
nat = read_table(national_path)

col_type = _pick_col(nat, ["Type", "Class", "厂站类型", "Type_WWTP"])
col_scale = _pick_col(nat, ["Scale", "Capacity", "规模", "Scale_WWTP"])
col_proc = _pick_col(nat, ["Process", "工艺", "Process_category"])
col_prov = _pick_col(nat, ["Local_province", "Province", "省"])
col_eff = _pick_col(nat, ["Discharge_standard_WWTP", "Effluent", "出水标准", "排放标准", "class"])
col_clim = _pick_col(nat, ["Climate zone", "climate zone", "ClimateZone", "Climate", "Koppen", "气候", "climate_zone"])

if any(c is None for c in (col_type, col_scale, col_proc, col_prov)):
    raise ValueError("NATIONAL dataset is missing required Type, Design capacity, Process, or Province columns.")

nat["type_norm"] = nat[col_type].astype(str).str.lower().str.replace(r"[^a-z]", "", regex=True)
nat_A = nat[nat["type_norm"].str.contains("a")].copy()

# National design-capacity unit is 10^4 m^3/d in the source dataset.
nat_A["cap_thousand_m3_d"] = pd.to_numeric(nat_A[col_scale], errors="coerce") * 10.0
nat_A = nat_A[nat_A["cap_thousand_m3_d"] > 0].copy()

nat_A["province"] = nat_A[col_prov].astype(str).str.strip() if col_prov else "Unknown"
nat_A["province"] = nat_A["province"].map(_norm_province)

# Require grid and utilisation parameters for all retained provinces.
missing_util = sorted(set(nat_A.loc[~nat_A["province"].isin(PROVINCE_UTILISATION_2015), "province"]))
missing_grid = sorted(set(nat_A.loc[~nat_A["province"].isin(PROVINCE_GRID_REGION_2015), "province"]))
if missing_util:
    raise ValueError(f"Missing 2015 capacity-utilisation factor for province(s): {missing_util}")
if missing_grid:
    raise ValueError(f"Missing 2015 grid assignment for province(s): {missing_grid}")

nat_A["capacity_utilisation_2015"] = nat_A["province"].map(PROVINCE_UTILISATION_2015)
nat_A["grid_region_2015"] = nat_A["province"].map(PROVINCE_GRID_REGION_2015)
nat_A["grid_ef_2015_kgCO2_kWh"] = nat_A["grid_region_2015"].map(GRID_EF_2015)

TECH6_norm = [_norm_cat(x) for x in TECH6]
nat_A["tech_nat"] = nat_A[col_proc].astype(str).str.strip().map(_norm_cat)
nat_A["tech_nat"] = nat_A["tech_nat"].where(
    nat_A["tech_nat"].isin(TECH6_norm),
    other=_norm_cat("Others"),
)


nat_A["tech_local"] = nat_A["tech_nat"]

nat_A["effluent_class"] = nat_A[col_eff].map(_norm_effluent_5class) if col_eff else "Unknown"
nat_A["effluent_class"] = nat_A["effluent_class"].map(_norm_cat)
nat_A["climate_zone"] = nat_A[col_clim].map(_norm_cat) if col_clim else "Unknown"

# Use optional covariates only when national prediction data are available.
USE_EFFLUENT = nat_A["effluent_class"].ne("Unknown").any()
USE_CLIMATE = nat_A["climate_zone"].ne("Unknown").any()

nat_A["log_cap"] = np.log(nat_A["cap_thousand_m3_d"] + 1e-3)

# Annual throughput = design capacity * 365 * provincial utilisation.
nat_A["Q_m3_y"] = (
    nat_A["cap_thousand_m3_d"]
    * 1000.0
    * 365.0
    * nat_A["capacity_utilisation_2015"]
)

print(
    "National throughput after province-level utilisation calibration: "
    f"{nat_A['Q_m3_y'].sum() / 1e9:.3f} billion m3/y"
)

# ==========================================================
# Stage 2. Target construction and model fitting
# ==========================================================

def clean_local_df(raw_df, source_name):
    """Harmonise local inputs and construct component-specific training targets."""
    df = raw_df.loc[:, ~raw_df.columns.duplicated()].copy()

    col_tech = _pick_col(df, ["Technology", "Process", "工艺", "tech", "local tech"])

    # Plant capacity field (10^3 m^3/d).
    col_capacity = _pick_col(df, [
        "Design capacity",
        "Average treated flow",
        "Influent wastewater flow",
        "Plant capacity",
        "Capacity",
        "规模",
        "Scale",
    ])

    # Reported removal (mg/L).
    col_cod_removed = _pick_col(df, ["COD removed"])
    col_tn_removed = _pick_col(df, ["TN removed"])

    # Influent and effluent concentrations (mg/L).
    col_cod_in = _pick_col(df, ["Influent COD"])
    col_cod_out = _pick_col(df, ["Effluent COD"])
    col_tn_in = _pick_col(df, ["Influent TN"])
    col_tn_out = _pick_col(df, ["Effluent TN"])

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
    col_eff = _pick_col(df, ["Effluent standard", "出水标准", "排放标准", "discharge", "Discharge standard"])
    col_clim = _pick_col(df, ["Climate zone", "climate zone", "Climate", "Koppen", "气候", "climate_zone"])
    col_prov = _pick_col(df, ["Province", "省", "Local_province", "行政省份"])

    if col_tech is None:
        raise ValueError(f"{source_name}: Technology column not found.")
    if col_capacity is None:
        raise ValueError(f"{source_name}: plant capacity column not found.")

    data = pd.DataFrame(index=df.index)
    data["source_name"] = source_name

    tech_raw_norm = df[col_tech].astype(str).str.strip().map(_norm_cat)
    tech_harmonised = tech_raw_norm.map(local_alias_to_target)

    unknown_local_tech = sorted(
        set(tech_raw_norm.loc[tech_harmonised.isna()].dropna().tolist())
    )
    if unknown_local_tech:
        raise ValueError(
            f"{source_name}: local input label(s) not found in the crosswalk: "
            f"{unknown_local_tech}. Add or correct the required mapping."
        )

    data["tech_local"] = tech_harmonised
    data["COD_in_mg_L"] = pd.to_numeric(df[col_cod_in], errors="coerce") if col_cod_in else np.nan
    data["TN_in_mg_L"] = pd.to_numeric(df[col_tn_in], errors="coerce") if col_tn_in else np.nan

    data["cap_thousand_m3_d"] = pd.to_numeric(df[col_capacity], errors="coerce")

    # Electricity-use intensity (kWh/m^3).
    data["e_kwh_per_m3"] = (
        pd.to_numeric(df[col_e], errors="coerce") if col_e else np.nan
    )

    # Use reported removal, then paired concentrations, then a configured proxy.
    proxy_key = _removal_proxy_key(source_name)
    pollutant_columns = [
        ("COD", col_cod_removed, col_cod_in, col_cod_out),
        ("TN", col_tn_removed, col_tn_in, col_tn_out),
    ]
    for pollutant, removed_col, in_col, out_col in pollutant_columns:
        removed = (pd.to_numeric(df[removed_col], errors="coerce")
                   if removed_col else pd.Series(np.nan, index=df.index))
        if in_col and out_col:
            paired = (pd.to_numeric(df[in_col], errors="coerce")
                      - pd.to_numeric(df[out_col], errors="coerce"))
            removed = removed.fillna(paired)
        proxy_mask = pd.Series(False, index=df.index)
        if in_col and proxy_key is not None:
            influent = pd.to_numeric(df[in_col], errors="coerce")
            proxy_mask = removed.isna() & influent.notna()
            removed.loc[proxy_mask] = influent.loc[proxy_mask] * REMOVAL_FRACTION_PROXY[proxy_key][pollutant]
        data[f"{pollutant}_removed_mg_L"] = removed
        data[f"{pollutant}_proxy_key"] = ""
        data.loc[proxy_mask, f"{pollutant}_proxy_key"] = proxy_key
        if in_col and removed.isna().any() and proxy_key is None:
            raise ValueError(
                f"{source_name}: {pollutant} removal is unavailable for "
                f"{int(removed.isna().sum())} row(s). Provide removal data or configure a removal fraction."
            )

    # Validate removal concentrations.
    if (data["COD_removed_mg_L"].dropna() < 0).any():
        raise ValueError(f"{source_name}: negative COD removed detected.")
    if (data["TN_removed_mg_L"].dropna() < 0).any():
        raise ValueError(f"{source_name}: negative TN removed detected.")

    # Partial upstream operational-input proxy.
    col_chem = _pick_col(df, ["Chemical Cost", "Chem cost", "化学品费", "Chemicals"])
    col_rep = _pick_col(df, ["Repair and maintenance", "Repair Cost", "维修", "Maintenance"])
    col_admin = _pick_col(df, ["Administrative Cost", "Admin cost", "管理", "Administration"])

    data["chem_cost_rmb_per_km3"] = pd.to_numeric(df[col_chem], errors="coerce").fillna(0.0) if col_chem else 0.0
    data["repair_cost_rmb_per_km3"] = pd.to_numeric(df[col_rep], errors="coerce").fillna(0.0) if col_rep else 0.0
    data["admin_cost_rmb_per_km3"] = pd.to_numeric(df[col_admin], errors="coerce").fillna(0.0) if col_admin else 0.0

    data["province"] = df[col_prov].map(_norm_province) if col_prov else "Unknown"

    data["effluent_class"] = df[col_eff].map(_norm_effluent_5class) if col_eff else "Unknown"
    data["climate_zone"] = df[col_clim].map(_norm_cat) if col_clim else "Unknown"

    data = data[data["cap_thousand_m3_d"] > 0].copy()

    # Scope 1 targets (kg gas/m^3).
    data["int_ch4_kg_m3"] = (data["COD_removed_mg_L"] * 1e-3) * (EF_CH4_MEAN / 1000.0)
    data["int_n2o_kg_m3"] = (data["TN_removed_mg_L"] * 1e-3) * (EF_N2O_MEAN / 1000.0)

    term_t_per_1000m3 = (
        data["chem_cost_rmb_per_km3"] * EF_IO_CHEM_MEAN
        + data["repair_cost_rmb_per_km3"] * EF_IO_REPAIR_MEAN
        + data["admin_cost_rmb_per_km3"] * EF_IO_ADMIN_MEAN
    ) * 1e-6

    data["int_s3_kgCO2e_m3"] = term_t_per_1000m3
    data["log_cap"] = np.log(data["cap_thousand_m3_d"] + 1e-3)

    for c in ["tech_local", "effluent_class", "climate_zone", "province"]:
        data[c] = data[c].map(_norm_province if c == "province" else _norm_cat)

    return data


local_frames = [
    clean_local_df(df, source_name)
    for df, source_name in zip(local_raw_frames, local_source_names)
]
local_all = pd.concat(local_frames, ignore_index=True)

for c in ["tech_local", "effluent_class", "climate_zone", "province"]:
    local_all[c] = local_all[c].map(_norm_cat)


def fit_pp_keep(train_df, target_col, group_col="tech_local", eff_col="effluent_class", clim_col="climate_zone"):
    """Fit log-intensity using plant capacity and optional categorical covariates.

    Technology has a random intercept. Small samples use a mean-only model;
    unsuccessful mixed-model fits use a pooled-OLS fallback.
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

    has_eff = bool(USE_EFFLUENT and td[eff_col].nunique(dropna=True) > 1)
    has_clim = bool(USE_CLIMATE and td[clim_col].nunique(dropna=True) > 1)

    rhs = ["log_cap"]
    if has_eff:
        rhs.append(f"C({eff_col})")
    if has_clim:
        rhs.append(f"C({clim_col})")

    formula = "log_target ~ " + " + ".join(rhs)
    td["log_target"] = np.log(td[target_col].clip(lower=1e-12))

    try:
        mixed_model = smf.mixedlm(formula, td, groups=td[group_col], eval_env=-1)
        res = mixed_model.fit()
        if not bool(getattr(res, "converged", False)):
            res = mixed_model.fit(method="powell", maxiter=2000, disp=False)
        if not bool(getattr(res, "converged", False)):
            raise RuntimeError("MixedLM did not converge after the Powell retry.")
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
    except Exception as exc:
        # Report fallback because technology pooling is unavailable.
        print(f"Warning: {target_col}: {exc}. Using pooled-OLS fallback.")
        from statsmodels.formula.api import ols
        ols_res = ols(formula, data=td, eval_env=-1).fit()
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


# Fit component-specific models for Scope 1 and the partial upstream proxy.
targets = ["int_ch4_kg_m3", "int_n2o_kg_m3", "int_s3_kgCO2e_m3"]
models = {t: fit_pp_keep(local_all, t, group_col="tech_local") for t in targets}

# Fit electricity-use intensity (kWh/m^3).
electricity_model = fit_pp_keep(local_all, "e_kwh_per_m3", group_col="tech_local")

# ==========================================================
# Stage 3. National extrapolation
# ==========================================================

def predict_pp(modelinfo, df):
    """Predict intensities; unestimated categorical effects contribute zero."""
    if modelinfo["kind"] == "mean_only":
        return pd.Series([modelinfo["mean"]] * len(df), index=df.index)

    params = modelinfo["params"]
    re_dict = modelinfo.get("re_dict", {})
    eff_col = modelinfo["eff_col"]
    clim_col = modelinfo["clim_col"]
    group_col = modelinfo["group_col"]
    has_eff = modelinfo.get("has_eff", False)
    has_clim = modelinfo.get("has_clim", False)

    log_cap_series = df["log_cap"].astype(float)
    eff_series = df[eff_col].astype(str)
    clim_series = df[clim_col].astype(str)
    tech_series = df[group_col].astype(str)

    out = np.empty(len(df), dtype=float)
    for i in range(len(df)):
        v = float(params.get("Intercept", 0.0)) + float(params.get("log_cap", 0.0)) * float(log_cap_series.iloc[i])
        if has_eff:
            v += float(params.get(f"C({eff_col})[T.{eff_series.iloc[i]}]", 0.0))
        if has_clim:
            v += float(params.get(f"C({clim_col})[T.{clim_series.iloc[i]}]", 0.0))
        if modelinfo["kind"] == "mixedlm":
            v += float(re_dict.get(tech_series.iloc[i], 0.0))
        out[i] = np.exp(v)

    return pd.Series(out, index=df.index)


pred_ch4 = predict_pp(models["int_ch4_kg_m3"], nat_A).values
pred_n2o = predict_pp(models["int_n2o_kg_m3"], nat_A).values

# Convert predicted electricity use to Scope 2 emissions.
pred_electricity_kwh_m3 = predict_pp(electricity_model, nat_A).values
pred_s2 = pred_electricity_kwh_m3 * nat_A["grid_ef_2015_kgCO2_kWh"].values

pred_s3 = predict_pp(models["int_s3_kgCO2e_m3"], nat_A).values

# ==========================================================
# Stage 4. Inventory aggregation and outputs
# ==========================================================
Q = nat_A["Q_m3_y"].values
prov = nat_A["province"].values
tech_nat = nat_A["tech_nat"].values
cap = nat_A["cap_thousand_m3_d"].values

prov_codes, prov_levels = pd.factorize(prov, sort=True)
techprov_index = pd.MultiIndex.from_arrays([prov, tech_nat], names=["Province", "Tech"])
techprov_codes, techprov_levels = pd.factorize(techprov_index, sort=False)

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
prov_Q_MC = np.zeros((MC_N, nP))

tp_S1 = np.zeros((MC_N, nTP))
tp_S2 = np.zeros((MC_N, nTP))
tp_S3 = np.zeros((MC_N, nTP))
tp_TOT = np.zeros((MC_N, nTP))
tp_Q_MC = np.zeros((MC_N, nTP))

# ==========================================================
# Stage 5. Uncertainty and predictor-ablation analyses
# ==========================================================
# Stage 5.1 Monte Carlo uncertainty analysis
#
# Uncertain inputs are sampled where they enter the calculation:
# - Scope 1: emission factors and removal fractions in target construction.
# - Scope 2: regional grid factors after electricity-use prediction.
# - Upstream proxy: EEIO coefficients in target construction.
# - All components: provincial utilisation in annual treated volume.
#
np.random.seed(RANDOM_SEED)

# Scope 1 emission factors
samp_EF_CH4 = sample_lognormal(EF_CH4_MEAN, CV_EF_CH4_N2O, MC_N)
samp_EF_N2O = sample_lognormal(EF_N2O_MEAN, CV_EF_CH4_N2O, MC_N)

# Upstream-proxy EEIO coefficients
samp_IO_CHEM = sample_lognormal(EF_IO_CHEM_MEAN, CV_IO, MC_N)
samp_IO_REPAIR = sample_lognormal(EF_IO_REPAIR_MEAN, CV_IO, MC_N)
samp_IO_ADMIN = sample_lognormal(EF_IO_ADMIN_MEAN, CV_IO, MC_N)

# Draw one removal fraction per configured dataset and iteration.
removal_draws = {}
for source_name, vals in REMOVAL_FRACTION_PROXY.items():
    removal_draws[source_name] = {
        "COD": sample_bounded_normal(vals["COD"], CV_REMOVAL_COD, MC_N, 0.0, 1.0),
        "TN": sample_bounded_normal(vals["TN"], CV_REMOVAL_TN, MC_N, 0.0, 1.0),
    }

# Draw one emission factor per grid per iteration.
canonical_grid_regions = ["North", "Northeast", "East", "Central", "Northwest", "South"]
grid_draws = {
    region: sample_trunc_normal(GRID_EF_2015[region], CV_GRID, MC_N)
    for region in canonical_grid_regions
}

# One utilisation-factor draw per province per MC iteration.
util_draws = {
    p: sample_trunc_normal(u, CV_UTILISATION, MC_N)
    for p, u in PROVINCE_UTILISATION_2015.items()
}

# Grid uncertainty affects conversion, not the fitted electricity-use model.
pred_electricity_central = pred_electricity_kwh_m3.copy()

for k in range(MC_N):
    if (k + 1) % 50 == 0 or k == 0:
        print(f"Monte Carlo iteration {k + 1}/{MC_N}")

    # ----- Reconstruct uncertainty-sensitive local training targets -----
    local_k = local_all.copy()

    # Update only removals constructed from a proxy; retain reported/paired values.
    for proxy_key in REMOVAL_FRACTION_PROXY:
        for pollutant in ("COD", "TN"):
            mask = local_k[f"{pollutant}_proxy_key"].eq(proxy_key)
            local_k.loc[mask, f"{pollutant}_removed_mg_L"] = (
                local_k.loc[mask, f"{pollutant}_in_mg_L"]
                * removal_draws[proxy_key][pollutant][k]
            )

    # Scope 1 target construction with sampled emission factors.
    local_k["int_ch4_mc"] = (
        local_k["COD_removed_mg_L"] * 1e-3 * (samp_EF_CH4[k] / 1000.0)
    )
    local_k["int_n2o_mc"] = (
        local_k["TN_removed_mg_L"] * 1e-3 * (samp_EF_N2O[k] / 1000.0)
    )

    # Upstream-proxy target construction with sampled EEIO coefficients.
    local_k["int_s3_mc"] = (
        local_k["chem_cost_rmb_per_km3"] * samp_IO_CHEM[k]
        + local_k["repair_cost_rmb_per_km3"] * samp_IO_REPAIR[k]
        + local_k["admin_cost_rmb_per_km3"] * samp_IO_ADMIN[k]
    ) * 1e-6

    # Refit only models whose local training targets change under uncertainty.
    model_ch4_k = fit_pp_keep(local_k, "int_ch4_mc", group_col="tech_local")
    model_n2o_k = fit_pp_keep(local_k, "int_n2o_mc", group_col="tech_local")
    model_s3_k = fit_pp_keep(local_k, "int_s3_mc", group_col="tech_local")

    ch4_k = predict_pp(model_ch4_k, nat_A).values
    n2o_k = predict_pp(model_n2o_k, nat_A).values
    s3_k = predict_pp(model_s3_k, nat_A).values

    # Scope 2: sample the regional grid factor after predicting electricity use.
    sampled_grid_by_plant = np.array([
        grid_draws["Northwest"][k]
        if r == "Northwest_proxy"
        else grid_draws[r][k]
        for r in nat_A["grid_region_2015"].values
    ])
    s2_k = pred_electricity_central * sampled_grid_by_plant

    # Sample province-level utilisation before calculating annual treated volume.
    sampled_util_by_plant = np.array([
        util_draws[p][k] for p in nat_A["province"].values
    ])
    Q_k = cap * 1000.0 * 365.0 * sampled_util_by_plant

    E_CH4 = (ch4_k * Q_k * GWP_CH4) / 1000.0
    E_N2O = (n2o_k * Q_k * GWP_N2O) / 1000.0
    E_S1 = E_CH4 + E_N2O
    E_S2 = (s2_k * Q_k) / 1000.0
    E_S3 = (s3_k * Q_k) / 1000.0
    E_TOT = E_S1 + E_S2 + E_S3

    prov_Q_k = np.bincount(prov_codes, weights=Q_k, minlength=nP)
    tp_Q_k = np.bincount(techprov_codes, weights=Q_k, minlength=nTP)

    prov_Q_MC[k, :] = prov_Q_k
    prov_CH4[k, :] = np.bincount(prov_codes, weights=E_CH4, minlength=nP)
    prov_N2O[k, :] = np.bincount(prov_codes, weights=E_N2O, minlength=nP)
    prov_S1[k, :] = np.bincount(prov_codes, weights=E_S1, minlength=nP)
    prov_S2[k, :] = np.bincount(prov_codes, weights=E_S2, minlength=nP)
    prov_S3[k, :] = np.bincount(prov_codes, weights=E_S3, minlength=nP)
    prov_TOT[k, :] = np.bincount(prov_codes, weights=E_TOT, minlength=nP)
    prov_I_TOT[k, :] = prov_TOT[k, :] / np.maximum(prov_Q_k, 1e-30)

    tp_Q_MC[k, :] = tp_Q_k
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


mQ, loQ, hiQ = mean_ci(prov_Q_MC)
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
    "Q_billion_m3_y_mean": mQ / 1e9,
    "Q_billion_m3_y_lo": loQ / 1e9,
    "Q_billion_m3_y_hi": hiQ / 1e9,
    "CH4_MtCO2e_mean": mCH4 / 1e6, "CH4_MtCO2e_lo": loCH4 / 1e6, "CH4_MtCO2e_hi": hiCH4 / 1e6,
    "N2O_MtCO2e_mean": mN2O / 1e6, "N2O_MtCO2e_lo": loN2O / 1e6, "N2O_MtCO2e_hi": hiN2O / 1e6,
    "Scope1_MtCO2e_mean": mS1 / 1e6, "Scope1_MtCO2e_lo": loS1 / 1e6, "Scope1_MtCO2e_hi": hiS1 / 1e6,
    "Scope2_MtCO2e_mean": mS2 / 1e6, "Scope2_MtCO2e_lo": loS2 / 1e6, "Scope2_MtCO2e_hi": hiS2 / 1e6,
    "Scope3_MtCO2e_mean": mS3 / 1e6, "Scope3_MtCO2e_lo": loS3 / 1e6, "Scope3_MtCO2e_hi": hiS3 / 1e6,
    "Total_MtCO2e_mean": mTOT / 1e6, "Total_MtCO2e_lo": loTOT / 1e6, "Total_MtCO2e_hi": hiTOT / 1e6,
    "I_total_tCO2e_m3_mean": mI, "I_total_tCO2e_m3_lo": loI, "I_total_tCO2e_m3_hi": hiI,
    "I_total_kgCO2e_m3_mean": mI * 1000.0, "I_total_kgCO2e_m3_lo": loI * 1000.0, "I_total_kgCO2e_m3_hi": hiI * 1000.0,
}).sort_values("Total_MtCO2e_mean", ascending=False).reset_index(drop=True)

nat_Q = prov_Q_MC.sum(axis=1)
nat_CH4 = prov_CH4.sum(axis=1)
nat_N2O = prov_N2O.sum(axis=1)
nat_S1 = prov_S1.sum(axis=1)
nat_S2 = prov_S2.sum(axis=1)
nat_S3 = prov_S3.sum(axis=1)
nat_TOT = prov_TOT.sum(axis=1)

nQm, nQlo, nQhi = mean_ci_1d(nat_Q)
nCH4m, nCH4lo, nCH4hi = mean_ci_1d(nat_CH4)
nN2Om, nN2Olo, nN2Ohi = mean_ci_1d(nat_N2O)
nS1m, nS1lo, nS1hi = mean_ci_1d(nat_S1)
nS2m, nS2lo, nS2hi = mean_ci_1d(nat_S2)
nS3m, nS3lo, nS3hi = mean_ci_1d(nat_S3)
nTOTm, nTOTlo, nTOThi = mean_ci_1d(nat_TOT)

national_summary = pd.DataFrame({
    "Metric": [
        "Number of plants",
        "Annual treated volume (billion m3/y)",
        "CH4 (MtCO2e/y)",
        "N2O (MtCO2e/y)",
        "Scope 1 (MtCO2e/y)",
        "Scope 2 (MtCO2e/y)",
        "Partial upstream proxy (MtCO2e/y)",
        "Total (MtCO2e/y)",
    ],
    "Mean": [int(len(nat_A)), nQm / 1e9, nCH4m / 1e6, nN2Om / 1e6, nS1m / 1e6, nS2m / 1e6, nS3m / 1e6, nTOTm / 1e6],
    "UI2.5%": [np.nan, nQlo / 1e9, nCH4lo / 1e6, nN2Olo / 1e6, nS1lo / 1e6, nS2lo / 1e6, nS3lo / 1e6, nTOTlo / 1e6],
    "UI97.5%": [np.nan, nQhi / 1e9, nCH4hi / 1e6, nN2Ohi / 1e6, nS1hi / 1e6, nS2hi / 1e6, nS3hi / 1e6, nTOThi / 1e6],
})

tp_mQ, tp_loQ, tp_hiQ = mean_ci(tp_Q_MC)
tp_mS1, tp_loS1, tp_hiS1 = mean_ci(tp_S1)
tp_mS2, tp_loS2, tp_hiS2 = mean_ci(tp_S2)
tp_mS3, tp_loS3, tp_hiS3 = mean_ci(tp_S3)
tp_mT, tp_loT, tp_hiT = mean_ci(tp_TOT)

tp_I1 = tp_S1 / np.maximum(tp_Q_MC, 1e-30)
tp_I2 = tp_S2 / np.maximum(tp_Q_MC, 1e-30)
tp_I3 = tp_S3 / np.maximum(tp_Q_MC, 1e-30)
tp_IT = tp_TOT / np.maximum(tp_Q_MC, 1e-30)

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
    "Q_m3_y_mean": tp_mQ,
    "Q_m3_y_lo": tp_loQ,
    "Q_m3_y_hi": tp_hiQ,
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

# ==========================================================
# Stage 5.2. 2^4 full-factorial predictor-ablation analysis
# ==========================================================
# Refit all 16 predictor inclusion/exclusion specifications.


def flags_from_mask(mask: int):
    """Map a 0-15 mask to predictor inclusion flags (1=capacity, 2=climate, 4=effluent, 8=technology)."""
    return dict(
        use_capacity=bool(mask & 1),
        use_climate=bool(mask & 2),
        use_effluent=bool(mask & 4),
        use_technology=bool(mask & 8),
    )


def fit_pp_ablation(
    train_df,
    target_col,
    *,
    use_capacity=True,
    use_climate=True,
    use_effluent=True,
    use_technology=True,
    group_col="tech_local",
    eff_col="effluent_class",
    clim_col="climate_zone",
):
    """Refit with the selected fixed effects and technology random intercept.

    If mixed-model fitting fails, use a fixed technology effect in OLS.
    """
    td = train_df[train_df[target_col] > 1e-12].copy()

    if len(td) < 6:
        gm = td[target_col].mean() if len(td) > 0 else 0.0
        return {
            "kind": "mean_only",
            "mean": gm,
            "formula": "mean-only",
            "re_dict": {},
            "has_eff": False,
            "has_clim": False,
            "eff_col": eff_col,
            "clim_col": clim_col,
            "group_col": group_col,
        }

    has_eff = bool(use_effluent and USE_EFFLUENT and td[eff_col].nunique(dropna=True) > 1)
    has_clim = bool(use_climate and USE_CLIMATE and td[clim_col].nunique(dropna=True) > 1)

    rhs = []
    if use_capacity:
        rhs.append("log_cap")
    if has_eff:
        rhs.append(f"C({eff_col})")
    if has_clim:
        rhs.append(f"C({clim_col})")

    formula = "log_target ~ " + (" + ".join(rhs) if rhs else "1")
    td["log_target"] = np.log(td[target_col].clip(lower=1e-12))

    if use_technology and td[group_col].nunique(dropna=True) > 1:
        try:
            mixed_model = smf.mixedlm(
                formula, td, groups=td[group_col], eval_env=-1
            )
            res = mixed_model.fit()
            if not bool(getattr(res, "converged", False)):
                res = mixed_model.fit(method="powell", maxiter=2000, disp=False)
            if not bool(getattr(res, "converged", False)):
                raise RuntimeError("MixedLM did not converge after the Powell retry.")
            params = res.params.to_dict()
            re_dict = {
                k: float(v.iloc[0] if hasattr(v, "iloc") else v[0])
                for k, v in res.random_effects.items()
            }
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
        except Exception as exc:
            # Preserve technology in the OLS fallback.
            print(
                f"Warning: MixedLM failed for {target_col} under {formula}; "
                f"using OLS with fixed technology effect. Reason: {exc}"
            )
            from statsmodels.formula.api import ols
            rhs_ols = list(rhs) + [f"C({group_col})"]
            formula_ols = "log_target ~ " + (" + ".join(rhs_ols) if rhs_ols else "1")
            ols_res = ols(formula_ols, data=td, eval_env=-1).fit()
            return {
                "kind": "ols_tech_fixed",
                "formula": formula_ols + " (technology fixed-effect fallback)",
                "params": ols_res.params.to_dict(),
                "re_dict": {},
                "has_eff": has_eff,
                "has_clim": has_clim,
                "has_tech_fixed": True,
                "eff_col": eff_col,
                "clim_col": clim_col,
                "group_col": group_col,
            }

    from statsmodels.formula.api import ols
    ols_res = ols(formula, data=td, eval_env=-1).fit()
    return {
        "kind": "ols",
        "formula": formula,
        "params": ols_res.params.to_dict(),
        "re_dict": {},
        "has_eff": has_eff,
        "has_clim": has_clim,
        "eff_col": eff_col,
        "clim_col": clim_col,
        "group_col": group_col,
    }


def predict_ablation(modelinfo, df):
    """Predict from a refitted ablation model specification."""
    if modelinfo["kind"] == "mean_only":
        return pd.Series([modelinfo["mean"]] * len(df), index=df.index)

    params = modelinfo["params"]
    re_dict = modelinfo.get("re_dict", {})
    eff_col = modelinfo["eff_col"]
    clim_col = modelinfo["clim_col"]
    group_col = modelinfo["group_col"]
    has_eff = modelinfo.get("has_eff", False)
    has_clim = modelinfo.get("has_clim", False)
    has_tech_fixed = modelinfo.get("has_tech_fixed", False)

    out = np.empty(len(df), dtype=float)
    for i in range(len(df)):
        v = float(params.get("Intercept", 0.0))
        v += float(params.get("log_cap", 0.0)) * float(df["log_cap"].iloc[i])

        if has_eff:
            eff = str(df[eff_col].iloc[i])
            v += float(params.get(f"C({eff_col})[T.{eff}]", 0.0))

        if has_clim:
            clim = str(df[clim_col].iloc[i])
            v += float(params.get(f"C({clim_col})[T.{clim}]", 0.0))

        tech = str(df[group_col].iloc[i])
        if modelinfo["kind"] == "mixedlm":
            v += float(re_dict.get(tech, 0.0))
        elif has_tech_fixed:
            v += float(params.get(f"C({group_col})[T.{tech}]", 0.0))

        out[i] = np.exp(v)

    return pd.Series(out, index=df.index)


masks = np.arange(16)
FAC_CAPACITY = np.where((masks & 1) > 0, 1, -1)
FAC_CLIM = np.where((masks & 2) > 0, 1, -1)
FAC_EFF = np.where((masks & 4) > 0, 1, -1)
FAC_TECH = np.where((masks & 8) > 0, 1, -1)

terms = {
    "Plant_capacity": FAC_CAPACITY,
    "Technology": FAC_TECH,
    "Effluent": FAC_EFF,
    "Climate": FAC_CLIM,
    "Plant_capacity:Technology": FAC_CAPACITY * FAC_TECH,
    "Plant_capacity:Effluent": FAC_CAPACITY * FAC_EFF,
    "Plant_capacity:Climate": FAC_CAPACITY * FAC_CLIM,
    "Technology:Effluent": FAC_TECH * FAC_EFF,
    "Technology:Climate": FAC_TECH * FAC_CLIM,
    "Effluent:Climate": FAC_EFF * FAC_CLIM,
    "Plant_capacity:Technology:Effluent": FAC_CAPACITY * FAC_TECH * FAC_EFF,
    "Plant_capacity:Technology:Climate": FAC_CAPACITY * FAC_TECH * FAC_CLIM,
    "Plant_capacity:Effluent:Climate": FAC_CAPACITY * FAC_EFF * FAC_CLIM,
    "Technology:Effluent:Climate": FAC_TECH * FAC_EFF * FAC_CLIM,
    "Plant_capacity:Technology:Effluent:Climate": FAC_CAPACITY * FAC_TECH * FAC_EFF * FAC_CLIM,
}

factors = ["Plant_capacity", "Technology", "Effluent", "Climate"]
provinces = sorted(nat_A["province"].unique())

Y16_refit = pd.DataFrame(index=provinces)
model_spec_rows = []

for mask in range(16):
    flags = flags_from_mask(mask)
    print(f"Factorial model {mask + 1}/16")

    model_ch4 = fit_pp_ablation(
        local_all, "int_ch4_kg_m3", group_col="tech_local", **flags
    )
    model_n2o = fit_pp_ablation(
        local_all, "int_n2o_kg_m3", group_col="tech_local", **flags
    )
    model_elec = fit_pp_ablation(
        local_all, "e_kwh_per_m3", group_col="tech_local", **flags
    )
    model_s3 = fit_pp_ablation(
        local_all, "int_s3_kgCO2e_m3", group_col="tech_local", **flags
    )

    ch4 = predict_ablation(model_ch4, nat_A)
    n2o = predict_ablation(model_n2o, nat_A)
    electricity = predict_ablation(model_elec, nat_A)
    s2 = electricity * nat_A["grid_ef_2015_kgCO2_kWh"]
    s3 = predict_ablation(model_s3, nat_A)

    I_kg = total_intensity_kgCO2e_m3(ch4, n2o, s2, s3)
    prov_I = prov_flow_weighted_intensity_kg(nat_A, I_kg, w="Q_m3_y")
    Y16_refit[f"m{mask:02d}"] = prov_I.reindex(provinces).values

    model_spec_rows.append({
        "Run": f"m{mask:02d}",
        "Plant_capacity_included": int(flags["use_capacity"]),
        "Technology_included": int(flags["use_technology"]),
        "Effluent_included": int(flags["use_effluent"] and USE_EFFLUENT),
        "Climate_included": int(flags["use_climate"] and USE_CLIMATE),
        "CH4_model": model_ch4.get("formula", model_ch4["kind"]),
        "N2O_model": model_n2o.get("formula", model_n2o["kind"]),
        "Electricity_model": model_elec.get("formula", model_elec["kind"]),
        "Scope3_model": model_s3.get("formula", model_s3["kind"]),
    })

main_refit = pd.DataFrame(index=provinces, columns=factors, dtype=float)
total_refit = pd.DataFrame(index=provinces, columns=factors, dtype=float)

for prov_name in provinces:
    y = Y16_refit.loc[prov_name].values.astype(float)
    yc = y - y.mean()

    coef = {k: float(np.mean(yc * v)) for k, v in terms.items()}
    var_comp = {k: coef[k] ** 2 for k in coef}
    V = sum(var_comp.values())

    if V <= 1e-18:
        main_refit.loc[prov_name, :] = 0.0
        total_refit.loc[prov_name, :] = 0.0
        continue

    for f in factors:
        main_refit.loc[prov_name, f] = 100.0 * var_comp[f] / V

    # Total contribution includes the main effect and all related interactions.
    for f in factors:
        total_refit.loc[prov_name, f] = 100.0 * sum(
            v for k, v in var_comp.items() if f in k.split(":")
        ) / V

model_specs = pd.DataFrame(model_spec_rows)

# ==========================================================
# Final export
# ==========================================================
out_xlsx = os.path.join(OUTPUT_DIR, OUT_XLSX_NAME)

with pd.ExcelWriter(out_xlsx, engine="openpyxl") as w:
    national_summary.to_excel(w, sheet_name="National_Summary_MC", index=False)
    prov_summary.to_excel(w, sheet_name="Provincial_Summary_MC", index=False)
    tech_prov_weighted.to_excel(w, sheet_name="Tech_Prov_Weighted_MC", index=False)

    Y16_refit.reset_index().rename(columns={"index": "Province"}).to_excel(
        w, sheet_name="Factorial_Y_16runs", index=False
    )
    main_refit.reset_index().rename(columns={"index": "Province"}).to_excel(
        w, sheet_name="Factorial_MainEffect_pct", index=False
    )
    total_refit.reset_index().rename(columns={"index": "Province"}).to_excel(
        w, sheet_name="Factorial_TotalContrib_pct", index=False
    )
    model_specs.to_excel(w, sheet_name="Factorial_Model_Specs", index=False)

print(f"Saved workbook: {out_xlsx}")
