# Enzyme Kinetics Data Analysis

## Project Overview

This project demonstrates the application of Python programming, statistical analysis and mathematical modelling to enzyme kinetics.

A simulated enzyme kinetics dataset was analysed to investigate the relationship between substrate concentration and enzyme reaction rate using the Michaelis–Menten model.

The project combines computational biochemistry with scientific data visualisation and spreadsheet-based reporting.

## Project Objectives

- Analyse and organise enzyme kinetics data using Python.
- Calculate mean reaction rates and standard deviations.
- Determine coefficients of variation.
- Fit the Michaelis–Menten equation to reaction-rate data.
- Estimate maximum reaction rate (Vmax) and Michaelis constant (Km).
- Visualise enzyme kinetics using Python.
- Present results in an Excel workbook and scientific report.

## Dataset Description

The dataset contains:

- 24 observations.
- Eight substrate concentrations.
- Three replicates per concentration.
- Substrate concentrations ranging from 0.5 mM to 20 mM.

### Variables

| Variable | Description |
|---|---|
| Experiment_ID | Unique experiment identification |
| Substrate_Concentration_mM | Substrate concentration in millimolar |
| Replicate | Replicate number |
| Reaction_Rate | Recorded simulated reaction rate |

**Important:** The dataset is simulated for educational and computational analysis. It does not represent actual laboratory measurements.

## Methodology

The project follows these analytical stages:

1. Dataset importation and inspection using Pandas.
2. Data organisation and statistical analysis.
3. Calculation of mean, standard deviation and coefficient of variation.
4. Nonlinear curve fitting using SciPy.
5. Michaelis–Menten modelling.
6. Graphical visualisation using Matplotlib.
7. Export of results into Excel and CSV formats.
8. Development of a scientific dashboard using WPS Spreadsheets.

## Michaelis–Menten Model

The Michaelis–Menten equation is:

\[
v = \frac{V_{\max}[S]}{K_m+[S]}
\]

Where:

- \(v\) = reaction rate
- \(V_{\max}\) = maximum reaction rate predicted by the model
- \([S]\) = substrate concentration
- \(K_m\) = Michaelis constant

## Key Results

| Parameter | Estimated value |
|---|---:|
| Vmax | Approximately 93.67 |
| Km | Approximately 3.772 mM |
| Half Vmax | Approximately 46.835 |
| Number of observations | 24 |

The reaction rate increased with substrate concentration and gradually approached saturation, consistent with the general behaviour described by Michaelis–Menten kinetics.

These parameter estimates are derived from the simulated dataset and are not experimentally measured kinetic constants.

## Technologies Used

- Python 3.13.3
- Pandas
- NumPy
- SciPy
- Matplotlib
- OpenPyXL
- WPS Spreadsheets
- WPS Writer

## Project Structure

```text
enzyme-kinetics-data-analysis/
│
├── data/
│   └── enzyme_kinetics_data.xlsx
│
├── scripts/
│
├── results/
│   ├── enzyme_statistics.csv
│   ├── michaelis_menten_curve.png
│   └── enzyme_kinetics_analysis.xlsx
│
├── main.py
├── Enzyme_Kinetics_Scientific_Report.docx
└── README.md
```

## Limitations

- The dataset is simulated rather than experimentally obtained.
- Only eight substrate concentrations were analysed.
- Three replicates were used at each concentration.
- Additional experimental factors such as temperature, pH and enzyme concentration were not investigated.

## Future Improvements

- Analyse real laboratory enzyme kinetics data.
- Investigate the effects of temperature and pH.
- Compare different enzyme kinetic models.
- Develop interactive scientific visualisations.
- Extend the analysis to enzyme inhibition studies.

## Author

BSc Biochemistry Student

Computational Biochemistry | Scientific Data Analysis | Bioinformatics

## Disclaimer

This project is an educational computational biochemistry exercise. All dataset observations and derived kinetic parameters are based on simulated data.