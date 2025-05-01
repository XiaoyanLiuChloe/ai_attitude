# AI Attitude Analysis Project

## Project Overview

This project aims to analyze people's attitudes towards AI using machine learning techniques. It combines Principal Component Analysis (PCA) and Random Forest Regression to understand the relationship between comfortableness and capability perceptions of AI.

## Features

- Data preprocessing and standardization
- Principal Component Analysis (PCA) for dimensionality reduction
- Random Forest Regression with hyperparameter tuning
- Comprehensive visualization of results
- Automated report generation in Word format

## Requirements

```bash
pip install numpy pandas seaborn matplotlib scikit-learn python-docx tabulate colorama
```

## Project Structure

```
.
├── data/
│   └── survey_data.csv      # Raw survey data
├── outputs/
│   ├── tables/             # Generated tables
│   ├── visualizations/     # Generated plots
│   ├── pca_factors.csv     # PCA results
│   ├── rf_results.csv      # Random Forest results
│   └── AI_Attitude_Analysis_Report.docx  # Final report
└── ai_attitude_analysis.py # Main analysis script
```

## Usage

1. Place your survey data in `data/survey_data.csv`
2. Run the analysis script:

```bash
python ai_attitude_analysis.py
```

3. Check the generated outputs in the `outputs` directory

## Data Format

The input data should be a CSV file with the following columns:

- `Scom1` to `Scom42`: Comfortableness scores
- `Scap1` to `Scap42`: Capability scores
- `PR1` to `PR16`: Positive attitude scores
- `NR1` to `NR16`: Negative attitude scores

## Output

The script generates:

1. PCA analysis results
2. Random Forest regression results
3. Visualizations of the analysis
4. A comprehensive Word report

## Analysis Steps

1. Data Loading and Preprocessing
2. PCA Analysis for Comfortableness and Capability
3. Random Forest Regression Analysis
4. Results Visualization and Report Generation

## License

This project is licensed under the MIT License.

## Contact

For any questions or suggestions, please feel free to contact us by email. Xiaoyanliu@link.cuhk.edu.cn  or  Xinqingli@link.cuhk.edu.cn
