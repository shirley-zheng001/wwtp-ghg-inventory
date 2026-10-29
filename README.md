# Data-Limited National WWTP Greenhouse-Gas Inventory Framework

## Overview

This repository provides the Python implementation of a hierarchical bottom-up framework for constructing national greenhouse-gas (GHG) inventories of municipal wastewater treatment plants (WWTPs) under data-limited conditions.

The framework derives component-specific emission intensities from local plant-level datasets and extrapolates them to a national WWTP dataset. It estimates direct methane (CH₄) and nitrous oxide (N₂O) emissions, electricity-related emissions, and a partial upstream operational-input proxy for Scope 3. Uncertainty is propagated through Monte Carlo simulation, and predictor sensitivity is assessed using a 2⁴ full-factorial sensitivity (predictor-ablation) analysis.

China is used as the application case in the associated manuscript.

## Method summary

The workflow consists of four main stages:

1. Harmonise treatment-technology labels, effluent classes, climate zones, and capacity units across the local and national datasets.
2. Construct component-specific intensity targets from local plant-level data.
3. Fit hierarchical log-linear models and extrapolate the fitted intensities to the national WWTP dataset.
4. Aggregate national and provincial emissions, perform Monte Carlo uncertainty analysis, and conduct a 2⁴ full-factorial sensitivity (predictor-ablation) analysis.

Detailed methodological definitions, equations, assumptions, and limitations are provided in the associated manuscript and Supporting Information.

## Repository contents

```text
wwtp-ghg-inventory/
├── README.md
├── wwtp_inventory.py
├── requirements.txt
├── LICENSE
├── CITATION.cff
├── technology_crosswalk.xlsx
├── example_inputs/
│   ├── local_wwtp_example.xlsx
│   └── national_wwtp_example.xlsx
└── results/
    └── Inventory_outputs.xlsx
```

The example workbooks each contain 10 synthetic WWTP records and are provided solely to demonstrate the required input structure and workflow. They do not contain empirical plant-level observations and do not reproduce the national results reported in the manuscript.

## Software requirements

The script uses Python 3 and the packages specified in `requirements.txt`.

Install the required packages using:

```bash
pip install -r requirements.txt
```

The script uses a desktop file-selection window provided by Python's `tkinter` standard-library module.

## Running the analysis

Run the main script:

```bash
python wwtp_inventory.py
```

If numerical model fitting stalls on Windows, the following PowerShell commands can limit the numerical libraries to one CPU thread before running the script:

```powershell
$env:OPENBLAS_NUM_THREADS="1"
$env:OMP_NUM_THREADS="1"
$env:MKL_NUM_THREADS="1"
python wwtp_inventory.py
```

These commands set session-level environment variables; they are not additional Python dependencies.

The script will ask the user to select, in order:

1. one or more local WWTP datasets;
2. `technology_crosswalk.xlsx`;
3. the national WWTP dataset; and
4. an output folder.

To test the workflow, select:

- `example_inputs/local_wwtp_example.xlsx` as the local dataset;
- `technology_crosswalk.xlsx` as the technology crosswalk; and
- `example_inputs/national_wwtp_example.xlsx` as the national dataset.

## Required input structure

### Local plant-level datasets

| Field | Unit or format | Role in the framework |
|---|---|---|
| City | Text | Descriptive location information (not a model predictor) |
| Name of WWTPs | Text | Plant identifier (not a model predictor) |
| Plant capacity | Thousand m³ d⁻¹ | Capacity predictor |
| Technology | Categorical text | Treatment-technology grouping |
| Electricity Use Intensity | kWh m⁻³ | Electricity-use intensity target |
| COD removed | mg L⁻¹ | CH₄ intensity target |
| TN removed | mg L⁻¹ | N₂O intensity target |
| Chemical Cost | RMB per thousand m³ | Partial upstream operational-input proxy |
| Repair and Maintenance Cost | RMB per thousand m³ | Partial upstream operational-input proxy |
| Administrative Cost | RMB per thousand m³ | Partial upstream operational-input proxy |
| Effluent Standard | Categorical text | Optional fixed effect |
| Climate Zone | Categorical text | Optional fixed effect |

### National plant-level dataset

| Field | Unit or format | Role in the framework |
|---|---|---|
| Type_WWTP | Categorical text | Identification of above-ground municipal WWTPs |
| Name_WWTP | Text | Plant identifier |
| Local_province | Categorical text | Provincial aggregation |
| Local_city | Text | Location information |
| Local_lon | Decimal degrees | Location information |
| Local_lat | Decimal degrees | Location information |
| Design Capacity | 10⁴ m³ d⁻¹ | National capacity predictor |
| Process_category | Categorical text | National treatment-technology category |
| Effluent Standard | Categorical text | Fixed-effect category |
| Climate Zone | Categorical text | Fixed-effect category |

The script recognises several equivalent column labels. For Scope 1, removal can also be calculated from paired influent–effluent concentrations or estimated using configured source-specific removal fractions where neither is available.

## Technology crosswalk

`technology_crosswalk.xlsx` maps the treatment-technology categories reported in the national dataset to the harmonised technology classes used for model fitting. The first column contains the national categories, and the second column contains the corresponding local modelling categories.

## Empirical data sources

The empirical datasets used in the associated manuscript were compiled from the following published sources.

| Dataset | Role | Records used | Source |
|---|---|---:|---|
| Shenzhen | Local model fitting | 26 WWTPs | Liao et al. [1] |
| Shanghai | Local model fitting | 45 WWTPs | Xi et al. [2] |
| Beijing | Local model fitting | 38 WWTPs | Zhou et al. [3] |
| Wuhan | Local model fitting | 30 WWTPs | Twagirayezu et al. [4] |
| Binzhou | Local model fitting | 4 WWTPs | Wang et al. [5] |
| Xinjiang | Local model fitting | 5 WWTPs | Zhao et al. [6] |
| China national dataset | National model application | 2,464 A-WWTP records | Zhou et al. [7], originally compiled from China Water Association data |

The original national source dataset contained 2,665 WWTP records across 29 provincial-level regions, including both above-ground and underground WWTPs. Application of the operational boundary retained 2,464 above-ground municipal WWTP records. Seven records with non-positive design capacity were excluded, leaving 2,457 WWTPs for national extrapolation. The six local datasets contain 148 plant records.

## Outputs

The script generates an Excel workbook named `Inventory_outputs.xlsx` containing seven worksheets:

- `National_Summary_MC`
- `Provincial_Summary_MC`
- `Tech_Prov_Weighted_MC`
- `Factorial_Y_16runs`
- `Factorial_MainEffect_pct`
- `Factorial_TotalContrib_pct`
- `Factorial_Model_Specs`

The workbook provided in `results/` contains the aggregated outputs used in the associated manuscript. It is an unmodified copy of the analytical output workbook and does not contain the restricted plant-level input data.

## Data availability

The empirical local and national plant-level input datasets are not redistributed through this repository because they were obtained from third-party sources. Users should consult the cited publications and original data providers for access conditions.

This repository provides:

- synthetic local and national input examples;
- the input-field structure and units;
- the author-developed technology crosswalk;
- the complete analysis code; and
- aggregated national, provincial, province–technology, and sensitivity-analysis outputs.

## Reproducibility settings

The public script uses:

- 1,000 Monte Carlo iterations;
- random seed 42; and
- a 2⁴ full-factorial sensitivity (predictor-ablation) analysis (16 model specifications).

Retaining these settings supports computational reproducibility when the same empirical input datasets are available. Results may differ when alternative datasets, emission factors, model specifications, or software versions are used.

## Citation

When using this code or framework, please cite the associated article:

> Full article citation and DOI to be added following publication.

The machine-readable citation information is provided in `CITATION.cff` and should be updated with the final article citation and DOI after publication.

## License

The code is released under the MIT License. Copyright remains with the author. The license permits use, modification, and redistribution of the code, including commercial use, provided that the copyright and license notices are retained.

## References

[1] Liao, X.; Tian, Y.; Gan, Y.; Ji, J. Quantifying Urban Wastewater Treatment Sector’s Greenhouse Gas Emissions Using a Hybrid Life Cycle Analysis Method: An Application on Shenzhen City in China. *Science of the Total Environment* **2020**, *745*, 141176. https://doi.org/10.1016/j.scitotenv.2020.141176.

[2] Xi, J.; Gong, H.; Zhang, Y.; Dai, X.; Chen, L. The Evaluation of GHG Emissions from Shanghai Municipal Wastewater Treatment Plants Based on IPCC and Operational Data Integrated Methods (ODIM). *Science of the Total Environment* **2021**, *797*, 148967. https://doi.org/10.1016/j.scitotenv.2021.148967.

[3] Zhou, X.; Yang, F.; Yang, F.; Feng, D.; Pan, T.; Liao, H. Analyzing Greenhouse Gas Emissions from Municipal Wastewater Treatment Plants Using Pollutants Parameter Normalizing Method: A Case Study of Beijing. *Journal of Cleaner Production* **2022**, *376*, 134093. https://doi.org/10.1016/j.jclepro.2022.134093.

[4] Twagirayezu, E.; Luo, X.; Fan, L.; Xie, X.; Zan, F.; Lu, X.; Wu, X.; Liu, X. Operational Data-Based Greenhouse Gas Analysis for Wastewater and Sludge Treatment at a Megacity Level. *Journal of Cleaner Production* **2025**, *532*, 146894. https://doi.org/10.1016/j.jclepro.2025.146894.

[5] Wang, H.; Li, L.; Lin, Z.; Abulimiti, A.; Guan, M.; Zhao, T.; Tian, Y. Optimization of Greenhouse Gas Accounting Methods for Wastewater Treatment Plants in East Chinese Regions: A Comparative Analysis of IPCC and Group Standards Based on 49 Plants in Shandong Province. *Applied Sciences* **2025**, *15* (11), 6175. https://doi.org/10.3390/app15116175.

[6] Zhao, J.; Lu, J.; Zhao, H.; Yan, Y.; Dong, H. In Five Wastewater Treatment Plants in Xinjiang, China: Removal Processes for Illicit Drugs, Their Occurrence in Receiving River Waters, and Ecological Risk Assessment. *Chemosphere* **2023**, *339*, 139668. https://doi.org/10.1016/j.chemosphere.2023.139668.

[7] Zhou, S.; Gong, H.; Chen, X.; Wang, X.; Wang, H.; Zhang, Y.; Zhu, D.; Cao, X.; Li, S.; Dai, X. A Dataset of Distribution and Characterization of Underground Wastewater Treatment Plants in China. *Scientific Data* **2024**, *11*, 941. https://doi.org/10.1038/s41597-024-03815-x.
