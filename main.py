
from pathlib import Path

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt

from scipy.optimize import curve_fit

from openpyxl import Workbook
from openpyxl.styles import (
    Font,
    PatternFill,
    Alignment,
    Border,
    Side
)
from openpyxl.chart import ScatterChart, Series, Reference
from openpyxl.utils import get_column_letter


# =====================================================
# 1. PROJECT PATHS
# =====================================================

PROJECT_DIR = Path(__file__).resolve().parent

DATA_FILE = PROJECT_DIR / "data" / "enzyme_kinetics_data.xlsx"
RESULTS_DIR = PROJECT_DIR / "results"

RESULTS_DIR.mkdir(exist_ok=True)

OUTPUT_FILE = RESULTS_DIR / "enzyme_kinetics_analysis.xlsx"
GRAPH_FILE = RESULTS_DIR / "michaelis_menten_curve.png"
CSV_FILE = RESULTS_DIR / "enzyme_statistics.csv"


# =====================================================
# 2. LOAD AND INSPECT DATA
# =====================================================

data = pd.read_excel(DATA_FILE)

required_columns = [
    "Experiment_ID",
    "Substrate_Concentration_mM",
    "Replicate",
    "Reaction_Rate"
]

missing_columns = [
    column for column in required_columns
    if column not in data.columns
]

if missing_columns:
    raise ValueError(
        f"Dataset is missing required columns: {missing_columns}"
    )

if data.empty:
    raise ValueError("The input dataset is empty.")

if data[required_columns].isnull().any().any():
    raise ValueError(
        "The dataset contains missing values in required columns."
    )

if (data["Substrate_Concentration_mM"] <= 0).any():
    raise ValueError(
        "Substrate concentrations must be greater than zero."
    )

if (data["Reaction_Rate"] < 0).any():
    raise ValueError(
        "Reaction rates cannot be negative."
    )

print("\n--- ENZYME KINETICS ANALYSIS ---\n")

print(f"Observations: {len(data)}")
print(f"Missing values: {data.isnull().sum().sum()}")
print(f"Duplicate rows: {data.duplicated().sum()}")


# =====================================================
# 3. SUMMARY STATISTICS
# =====================================================

summary = data.groupby(
    "Substrate_Concentration_mM"
)["Reaction_Rate"].agg(
    Mean="mean",
    SD="std",
    N="count"
).reset_index()

summary["CV_Percent"] = (
    summary["SD"] / summary["Mean"]
) * 100

summary["SD"] = summary["SD"].fillna(0)

print("\n--- SUMMARY STATISTICS ---\n")
print(summary.round(3).to_string(index=False))


# =====================================================
# 4. MICHAELIS-MENTEN MODEL
# =====================================================

def michaelis_menten(substrate, vmax, km):
    return (vmax * substrate) / (km + substrate)


substrate = summary[
    "Substrate_Concentration_mM"
].values

velocity = summary["Mean"].values

parameters, covariance = curve_fit(
    michaelis_menten,
    substrate,
    velocity,
    p0=[90, 3],
    bounds=(0, np.inf),
    maxfev=10000
)

vmax, km = parameters

half_vmax = vmax / 2

print("\n--- KINETIC PARAMETERS ---\n")
print(f"Estimated Vmax: {vmax:.3f}")
print(f"Estimated Km: {km:.3f} mM")
print(f"Half Vmax: {half_vmax:.3f}")


# =====================================================
# 5. GENERATE FITTED CURVE DATA
# =====================================================

smooth_substrate = np.linspace(
    0,
    substrate.max() * 1.15,
    300
)

fitted_velocity = michaelis_menten(
    smooth_substrate,
    vmax,
    km
)

fitted_data = pd.DataFrame({
    "Substrate_Concentration_mM": smooth_substrate,
    "Predicted_Reaction_Rate": fitted_velocity
})


# =====================================================
# 6. GENERATE MICHAELIS-MENTEN GRAPH
# =====================================================

plt.figure(figsize=(10, 6))

plt.errorbar(
    substrate,
    velocity,
    yerr=summary["SD"].values,
    fmt="o",
    capsize=5,
    label="Simulated mean ± SD"
)

plt.plot(
    smooth_substrate,
    fitted_velocity,
    label="Michaelis-Menten fitted curve"
)

plt.axhline(
    vmax,
    linestyle="--",
    label=f"Vmax = {vmax:.2f}"
)

plt.axvline(
    km,
    linestyle=":",
    label=f"Km = {km:.2f} mM"
)

plt.title("Enzyme Kinetics: Michaelis-Menten Analysis")
plt.xlabel("Substrate Concentration (mM)")
plt.ylabel("Mean Reaction Rate (arbitrary units)")
plt.legend()
plt.grid(alpha=0.3)
plt.tight_layout()

plt.savefig(GRAPH_FILE, dpi=300)
plt.close()


# =====================================================
# 7. CREATE EXCEL WORKBOOK
# =====================================================

workbook = Workbook()

ws_summary = workbook.active
ws_summary.title = "Summary Statistics"

ws_parameters = workbook.create_sheet(
    "Kinetic Parameters"
)

ws_curve = workbook.create_sheet(
    "Fitted Curve Data"
)

ws_dashboard = workbook.create_sheet(
    "Dashboard"
)


# =====================================================
# 8. DEFINE FORMATTING
# =====================================================

header_fill = PatternFill(
    fill_type="solid",
    fgColor="1F4E78"
)

header_font = Font(
    bold=True,
    color="FFFFFF",
    size=11
)

title_font = Font(
    bold=True,
    size=16,
    color="1F4E78"
)

subtitle_font = Font(
    size=10,
    color="666666"
)

section_font = Font(
    bold=True,
    size=12,
    color="FFFFFF"
)

kpi_label_font = Font(
    bold=True,
    size=11,
    color="FFFFFF"
)

kpi_value_font = Font(
    bold=True,
    size=19,
    color="1F4E78"
)

normal_font = Font(
    size=10,
    color="333333"
)

thin_border = Border(
    bottom=Side(
        style="thin",
        color="D9E2F3"
    )
)

blue_fill = PatternFill(
    fill_type="solid",
    fgColor="1F4E78"
)

light_blue_fill = PatternFill(
    fill_type="solid",
    fgColor="D9EAF7"
)

light_grey_fill = PatternFill(
    fill_type="solid",
    fgColor="F2F2F2"
)


def format_worksheet(worksheet):

    worksheet.freeze_panes = "A2"

    for cell in worksheet[1]:

        cell.fill = header_fill
        cell.font = header_font

        cell.alignment = Alignment(
            horizontal="center",
            vertical="center"
        )

    worksheet.row_dimensions[1].height = 25

    for column_cells in worksheet.columns:

        max_length = 0

        column_letter = get_column_letter(
            column_cells[0].column
        )

        for cell in column_cells:

            cell.alignment = Alignment(
                vertical="center"
            )

            cell.border = thin_border

            if cell.value is not None:

                max_length = max(
                    max_length,
                    len(str(cell.value))
                )

        worksheet.column_dimensions[
            column_letter
        ].width = min(max_length + 4, 40)


# =====================================================
# 9. SUMMARY STATISTICS WORKSHEET
# =====================================================

ws_summary.append([
    "Substrate Concentration (mM)",
    "Mean Reaction Rate",
    "Standard Deviation",
    "Replicate Count",
    "CV (%)"
])

for _, row in summary.iterrows():

    ws_summary.append([
        float(row["Substrate_Concentration_mM"]),
        float(row["Mean"]),
        float(row["SD"]),
        int(row["N"]),
        float(row["CV_Percent"])
    ])

format_worksheet(ws_summary)

for row in ws_summary.iter_rows(
    min_row=2,
    max_row=ws_summary.max_row,
    min_col=1,
    max_col=5
):

    for cell in row:
        cell.number_format = "0.000"


# =====================================================
# 10. KINETIC PARAMETERS WORKSHEET
# =====================================================

ws_parameters.append([
    "Parameter",
    "Estimated Value",
    "Unit",
    "Scientific Interpretation"
])

parameter_rows = [
    [
        "Vmax",
        float(vmax),
        "Arbitrary rate units",
        "Predicted maximum reaction rate"
    ],
    [
        "Km",
        float(km),
        "mM",
        "Substrate concentration at half Vmax"
    ],
    [
        "Half Vmax",
        float(half_vmax),
        "Arbitrary rate units",
        "Half of the estimated maximum reaction rate"
    ]
]

for row in parameter_rows:
    ws_parameters.append(row)

format_worksheet(ws_parameters)

ws_parameters["A7"] = "DATASET NOTE"
ws_parameters["A7"].font = title_font

ws_parameters["A8"] = (
    "All observations in this project are simulated "
    "educational data, not real experimental measurements."
)

ws_parameters.merge_cells("A8:D8")

ws_parameters["A8"].alignment = Alignment(
    wrap_text=True,
    vertical="center"
)

ws_parameters.row_dimensions[8].height = 35

ws_parameters["A10"] = "MODEL"
ws_parameters["A10"].font = title_font

ws_parameters["A11"] = (
    "Michaelis-Menten equation: "
    "v = (Vmax × [S]) / (Km + [S])"
)

ws_parameters.merge_cells("A11:D11")

ws_parameters.column_dimensions["D"].width = 55


# =====================================================
# 11. FITTED CURVE DATA WORKSHEET
# =====================================================

ws_curve.append([
    "Substrate Concentration (mM)",
    "Predicted Reaction Rate"
])

for _, row in fitted_data.iterrows():

    ws_curve.append([
        float(row["Substrate_Concentration_mM"]),
        float(row["Predicted_Reaction_Rate"])
    ])

format_worksheet(ws_curve)


# =====================================================
# 12. FITTED CURVE CHART
# =====================================================

curve_chart = ScatterChart()

curve_chart.title = "Michaelis-Menten Fitted Curve"

curve_chart.x_axis.title = "Substrate Concentration (mM)"
curve_chart.y_axis.title = "Predicted Reaction Rate"

curve_chart.style = 13
curve_chart.height = 10
curve_chart.width = 18

curve_x = Reference(
    ws_curve,
    min_col=1,
    min_row=2,
    max_row=ws_curve.max_row
)

curve_y = Reference(
    ws_curve,
    min_col=2,
    min_row=1,
    max_row=ws_curve.max_row
)

curve_series = Series(
    curve_y,
    curve_x,
    title_from_data=True
)

curve_series.graphicalProperties.line.width = 25000
curve_series.graphicalProperties.line.solidFill = "ED7D31"

curve_series.marker.symbol = "none"

curve_chart.series.append(curve_series)

ws_curve.add_chart(
    curve_chart,
    "D2"
)


# =====================================================
# 13. DASHBOARD: GENERAL LAYOUT
# =====================================================

ws_dashboard.sheet_view.showGridLines = False

for column in range(1, 15):

    ws_dashboard.column_dimensions[
        get_column_letter(column)
    ].width = 14


# Main title

ws_dashboard.merge_cells("A1:N1")

ws_dashboard["A1"] = "ENZYME KINETICS ANALYSIS DASHBOARD"

ws_dashboard["A1"].font = Font(
    bold=True,
    size=20,
    color="FFFFFF"
)

ws_dashboard["A1"].fill = blue_fill

ws_dashboard["A1"].alignment = Alignment(
    horizontal="center",
    vertical="center"
)

ws_dashboard.row_dimensions[1].height = 38


# Subtitle

ws_dashboard.merge_cells("A2:N2")

ws_dashboard["A2"] = (
    "Statistical Analysis and Michaelis-Menten Modelling"
)

ws_dashboard["A2"].font = subtitle_font

ws_dashboard["A2"].alignment = Alignment(
    horizontal="center"
)

ws_dashboard.row_dimensions[2].height = 25


# Data source note

ws_dashboard.merge_cells("A3:N3")

ws_dashboard["A3"] = (
    "Educational analysis using simulated enzyme kinetics data"
)

ws_dashboard["A3"].font = Font(
    italic=True,
    size=9,
    color="777777"
)

ws_dashboard["A3"].alignment = Alignment(
    horizontal="center"
)


# =====================================================
# 14. DASHBOARD: KPI CARDS
# =====================================================

def create_kpi_card(
    label_range,
    value_range,
    label,
    formula,
    number_format="0.000"
):

    ws_dashboard.merge_cells(label_range)
    ws_dashboard.merge_cells(value_range)

    label_cell = ws_dashboard[label_range.split(":")[0]]
    value_cell = ws_dashboard[value_range.split(":")[0]]

    label_cell.value = label
    label_cell.fill = blue_fill
    label_cell.font = kpi_label_font

    label_cell.alignment = Alignment(
        horizontal="center",
        vertical="center"
    )

    value_cell.value = formula
    value_cell.fill = light_blue_fill
    value_cell.font = kpi_value_font

    value_cell.number_format = number_format

    value_cell.alignment = Alignment(
        horizontal="center",
        vertical="center"
    )


create_kpi_card(
    "A5:C5",
    "A6:C7",
    "ESTIMATED VMAX",
    "='Kinetic Parameters'!B2"
)

create_kpi_card(
    "E5:G5",
    "E6:G7",
    "ESTIMATED KM",
    "='Kinetic Parameters'!B3",
    "0.000"
)

create_kpi_card(
    "I5:K5",
    "I6:K7",
    "HALF VMAX",
    "='Kinetic Parameters'!B4"
)

create_kpi_card(
    "A9:C9",
    "A10:C11",
    "TOTAL OBSERVATIONS",
    "=SUM('Summary Statistics'!D2:D9)",
    "0"
)

create_kpi_card(
    "E9:G9",
    "E10:G11",
    "SUBSTRATE CONCENTRATIONS",
    "=COUNTA('Summary Statistics'!A2:A9)",
    "0"
)

create_kpi_card(
    "I9:K9",
    "I10:K11",
    "REPLICATES PER CONCENTRATION",
    "=AVERAGE('Summary Statistics'!D2:D9)",
    "0.0"
)


# =====================================================
# 15. DASHBOARD: MICHAELIS-MENTEN CHART
# =====================================================

dashboard_curve_chart = ScatterChart()

dashboard_curve_chart.title = "Michaelis-Menten Model"

dashboard_curve_chart.x_axis.title = (
    "Substrate Concentration (mM)"
)

dashboard_curve_chart.y_axis.title = (
    "Predicted Reaction Rate"
)

dashboard_curve_chart.style = 13

dashboard_curve_chart.height = 9
dashboard_curve_chart.width = 17

dashboard_curve_series = Series(
    curve_y,
    curve_x,
    title_from_data=True
)

dashboard_curve_series.graphicalProperties.line.width = 25000

dashboard_curve_series.graphicalProperties.line.solidFill = "ED7D31"

dashboard_curve_series.marker.symbol = "none"

dashboard_curve_chart.series.append(
    dashboard_curve_series
)

ws_dashboard.add_chart(
    dashboard_curve_chart,
    "A14"
)


# =====================================================
# 16. DASHBOARD: OBSERVED REACTION RATE CHART
# =====================================================

observed_chart = ScatterChart()

observed_chart.title = "Observed Mean Reaction Rates"

observed_chart.x_axis.title = (
    "Substrate Concentration (mM)"
)

observed_chart.y_axis.title = (
    "Mean Reaction Rate"
)

observed_chart.style = 13

observed_chart.height = 9
observed_chart.width = 17

observed_x = Reference(
    ws_summary,
    min_col=1,
    min_row=2,
    max_row=ws_summary.max_row
)

observed_y = Reference(
    ws_summary,
    min_col=2,
    min_row=1,
    max_row=ws_summary.max_row
)

observed_series = Series(
    observed_y,
    observed_x,
    title_from_data=True
)

observed_series.graphicalProperties.line.solidFill = "4472C4"

observed_series.marker.symbol = "circle"
observed_series.marker.size = 7

observed_chart.series.append(
    observed_series
)

ws_dashboard.add_chart(
    observed_chart,
    "I14"
)


# =====================================================
# 17. DASHBOARD: SCIENTIFIC INTERPRETATION
# =====================================================

ws_dashboard.merge_cells("A33:N33")

ws_dashboard["A33"] = "SCIENTIFIC INTERPRETATION"

ws_dashboard["A33"].fill = blue_fill
ws_dashboard["A33"].font = section_font

ws_dashboard["A33"].alignment = Alignment(
    horizontal="center"
)

interpretations = [
    (
        "1. Estimated Vmax",
        f"The fitted maximum reaction rate is approximately "
        f"{vmax:.2f} arbitrary units."
    ),
    (
        "2. Estimated Km",
        f"The estimated Km is {km:.3f} mM, corresponding "
        f"to half Vmax under the fitted model."
    ),
    (
        "3. Substrate concentration effect",
        "Reaction rate increases with substrate concentration "
        "and gradually approaches saturation."
    ),
    (
        "4. Data quality",
        "Three simulated replicates were used at each concentration. "
        "The calculated CV values were relatively low."
    )
]

current_row = 35

for heading, explanation in interpretations:

    ws_dashboard.merge_cells(
        start_row=current_row,
        start_column=1,
        end_row=current_row,
        end_column=14
    )

    heading_cell = ws_dashboard.cell(
        row=current_row,
        column=1
    )

    heading_cell.value = heading
    heading_cell.font = Font(
        bold=True,
        size=11,
        color="1F4E78"
    )

    current_row += 1

    ws_dashboard.merge_cells(
        start_row=current_row,
        start_column=1,
        end_row=current_row,
        end_column=14
    )

    explanation_cell = ws_dashboard.cell(
        row=current_row,
        column=1
    )

    explanation_cell.value = explanation

    explanation_cell.font = normal_font

    explanation_cell.alignment = Alignment(
        wrap_text=True,
        vertical="center"
    )

    ws_dashboard.row_dimensions[current_row].height = 28

    current_row += 2


# =====================================================
# 18. DASHBOARD: DATA SOURCE AND LIMITATIONS
# =====================================================

ws_dashboard.merge_cells("A44:N44")

ws_dashboard["A44"] = "DATA SOURCE AND LIMITATIONS"

ws_dashboard["A44"].fill = blue_fill

ws_dashboard["A44"].font = section_font

ws_dashboard["A44"].alignment = Alignment(
    horizontal="center"
)

ws_dashboard.merge_cells("A46:N47")

ws_dashboard["A46"] = (
    "This project uses simulated educational enzyme kinetics data. "
    "The estimated kinetic parameters are derived through mathematical "
    "curve fitting and should not be interpreted as experimentally "
    "measured enzyme kinetic constants."
)

ws_dashboard["A46"].font = Font(
    size=10,
    italic=True,
    color="555555"
)

ws_dashboard["A46"].fill = light_grey_fill

ws_dashboard["A46"].alignment = Alignment(
    wrap_text=True,
    vertical="center",
    horizontal="left"
)

ws_dashboard.row_dimensions[46].height = 35
ws_dashboard.row_dimensions[47].height = 25


# =====================================================
# 19. DASHBOARD: FOOTER
# =====================================================

ws_dashboard.merge_cells("A50:N50")

ws_dashboard["A50"] = (
    "Computational Biochemistry | Enzyme Kinetics Data Analysis"
)

ws_dashboard["A50"].font = Font(
    italic=True,
    size=9,
    color="777777"
)

ws_dashboard["A50"].alignment = Alignment(
    horizontal="center"
)


# =====================================================
# 20. FINAL DASHBOARD FORMATTING
# =====================================================

for row in ws_dashboard.iter_rows():

    for cell in row:

        if cell.value is not None:

            cell.alignment = Alignment(
                horizontal=cell.alignment.horizontal or "left",
                vertical="center",
                wrap_text=True
            )


# =====================================================
# 21. EXPORT RESULTS
# =====================================================

workbook.save(OUTPUT_FILE)

summary.to_csv(
    CSV_FILE,
    index=False
)


# =====================================================
# 22. FINAL PROJECT REPORT
# =====================================================

print("\n--- EXCEL EXPORT COMPLETED ---\n")

print("Workbook saved to:")
print(OUTPUT_FILE)

print("\nGraph saved to:")
print(GRAPH_FILE)

print("\nCSV saved to:")
print(CSV_FILE)

print("\nWorksheets created:")

for index, worksheet in enumerate(
    workbook.worksheets,
    start=1
):

    print(f"{index}. {worksheet.title}")

print("\n--- PROJECT ANALYSIS COMPLETED SUCCESSFULLY ---")
