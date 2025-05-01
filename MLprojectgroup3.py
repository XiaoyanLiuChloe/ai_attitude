# Import necessary libraries
import numpy as np
import pandas as pd
import seaborn as sns
import matplotlib.pyplot as plt
from sklearn.preprocessing import StandardScaler
from sklearn.decomposition import PCA
from factor_analyzer import FactorAnalyzer
import matplotlib.cm as cm
from tabulate import tabulate
from colorama import Fore, Back, Style, init
import os
from termcolor import colored
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io
import textwrap

# Initialize colorama
init()

# Create output directory if it doesn't exist
os.makedirs('outputs', exist_ok=True)

# Create a new Word document
doc = Document()
doc.add_heading('PCA Analysis Report', 0)

# Set style for visualizations
plt.style.use('seaborn-v0_8-whitegrid')
colors = sns.color_palette("viridis", 10)

# Function to create safe filename
def create_safe_filename(text):
    # Replace any non-alphanumeric characters (except underscores) with underscores
    safe_name = "".join(c if c.isalnum() or c == '_' else '_' for c in text)
    # Remove multiple consecutive underscores
    while '__' in safe_name:
        safe_name = safe_name.replace('__', '_')
    # Remove leading/trailing underscores
    return safe_name.strip('_')

# Function to create image from text
def create_text_image(text, filename, title=None, width=1200, font_size=30, padding=50):
    # Create safe filename
    safe_filename = create_safe_filename(filename)
    
    # Calculate required height based on text content
    font = ImageFont.truetype("/System/Library/Fonts/Arial Unicode.ttf", font_size)
    lines = text.split('\n')
    line_height = font_size + 10
    text_height = len(lines) * line_height
    
    # Add space for title if provided
    title_height = font_size + 30 if title else 0
    
    # Create image with white background
    height = text_height + 2 * padding + title_height
    img = Image.new('RGB', (width, height), 'white')
    draw = ImageDraw.Draw(img)
    
    # Draw title if provided
    y = padding
    if title:
        title_font = ImageFont.truetype("/System/Library/Fonts/Arial Unicode.ttf", font_size + 10)
        draw.text((padding, y), title, font=title_font, fill='black')
        y += title_height + 20
    
    # Draw text
    for line in lines:
        draw.text((padding, y), line, font=font, fill='black')
        y += line_height
    
    # Save image
    img.save(f'outputs/tables/{safe_filename}.png')

# Function to print nice title and save as image
def print_title(title):
    print("\n" + "="*80)
    print(" "*30 + colored(title, 'cyan', attrs=['bold']))
    print("="*80 + "\n")
    
    # Create image version
    text = "="*80 + "\n" + " "*30 + title + "\n" + "="*80
    create_text_image(text, f"title_{create_safe_filename(title.lower())}", font_size=36)

# Function to print pretty tables and save as image
def print_pretty_table(data, headers, title):
    # Print to terminal
    print(colored(f"\n{title}", 'green', attrs=['bold']))
    table = tabulate(data, headers=headers, tablefmt="fancy_grid", floatfmt=".4f")
    print(table)
    
    # Create image version
    create_text_image(table, f"table_{create_safe_filename(title.lower())}", title=title)

# Function to add a section title to the document
def add_section_title(title):
    doc.add_heading(title, level=1)
    print("\n" + "="*80)
    print(" "*30 + colored(title, 'cyan', attrs=['bold']))
    print("="*80 + "\n")

# Function to add a table to the document
def add_table_to_doc(data, headers, title):
    # Print to terminal
    print(colored(f"\n{title}", 'green', attrs=['bold']))
    print(tabulate(data, headers=headers, tablefmt="fancy_grid", floatfmt=".4f"))
    
    # Add title to document
    doc.add_heading(title, level=2)
    
    # Create table in document
    table = doc.add_table(rows=1, cols=len(headers))
    table.style = 'Table Grid'
    
    # Add headers
    header_cells = table.rows[0].cells
    for i, header in enumerate(headers):
        header_cells[i].text = str(header)
        header_cells[i].paragraphs[0].runs[0].bold = True
    
    # Add data
    for row_data in data:
        row_cells = table.add_row().cells
        for i, cell_data in enumerate(row_data):
            row_cells[i].text = f"{cell_data:.4f}" if isinstance(cell_data, float) else str(cell_data)
    
    doc.add_paragraph()

# Function to add a figure to the document
def add_figure_to_doc(fig, title, filename):
    # Save figure to a temporary buffer
    buf = io.BytesIO()
    fig.savefig(buf, format='png', dpi=300, bbox_inches='tight')
    buf.seek(0)
    
    # Save figure to file
    fig.savefig(f'outputs/{filename}.png', dpi=300, bbox_inches='tight')
    
    # Add figure to document
    doc.add_heading(title, level=2)
    doc.add_picture(buf, width=Inches(6))
    doc.add_paragraph()
    
    plt.close()

# Load the dataset
add_section_title("LOADING DATASET")
data = pd.read_csv('csv/chbr_100014_AI_R_Data_mmc1.csv')
doc.add_paragraph(f"Successfully loaded dataset with {data.shape[0]} rows and {data.shape[1]} columns")
print(colored(f"Successfully loaded dataset with {data.shape[0]} rows and {data.shape[1]} columns", 'yellow'))

# Define variables for comfortableness and capability
com_vars = ['Scom3', 'Scom6', 'Scom7', 'Scom8', 'Scom11', 'Scom12', 'Scom13', 
            'Scom14', 'Scom15', 'Scom16', 'Scom17', 'Scom20', 'Scom21', 'Scom22', 
            'Scom27', 'Scom28', 'Scom29', 'Scom31', 'Scom32', 'Scom33', 'Scom36', 
            'Scom37', 'Scom40']

cap_vars = ['Scap1', 'Scap4', 'Scap6', 'Scap7', 'Scap8', 'Scap10', 'Scap11', 
            'Scap15', 'Scap16', 'Scap19', 'Scap20', 'Scap25', 'Scap26', 'Scap27', 
            'Scap30', 'Scap31', 'Scap32', 'Scap36', 'Scap37', 'Scap38', 'Scap40']

# Check if all variables are in the dataset
missing_com = [var for var in com_vars if var not in data.columns]
missing_cap = [var for var in cap_vars if var not in data.columns]

if missing_com:
    warning = f"Warning: Missing comfortableness variables: {missing_com}"
    print(f"{Fore.RED}{warning}{Style.RESET_ALL}")
    doc.add_paragraph(warning)
if missing_cap:
    warning = f"Warning: Missing capability variables: {missing_cap}"
    print(f"{Fore.RED}{warning}{Style.RESET_ALL}")
    doc.add_paragraph(warning)

# Filter to only include available variables
com_vars = [var for var in com_vars if var in data.columns]
cap_vars = [var for var in cap_vars if var in data.columns]

# Function to perform PCA and generate plots
def perform_pca_analysis(data, variables, title_prefix):
    add_section_title(f"PCA ANALYSIS: {title_prefix.upper()}")
    
    # Extract data for the selected variables
    X = data[variables].dropna()
    doc.add_paragraph(f"Analyzing {len(variables)} variables with {X.shape[0]} valid samples")
    print(colored(f"Analyzing {len(variables)} variables with {X.shape[0]} valid samples", 'yellow'))
    
    # Standardize the data
    scaler = StandardScaler()
    X_scaled = scaler.fit_transform(X)
    
    # PCA for Scree Plot
    pca = PCA()
    pca.fit(X_scaled)
    
    # Scree Plot
    plt.figure(figsize=(10, 6))
    components = np.arange(1, min(len(variables) + 1, 11))
    plt.bar(components, pca.explained_variance_ratio_[:len(components)], 
            alpha=0.7, align='center', color=colors)
    plt.plot(components, np.cumsum(pca.explained_variance_ratio_[:len(components)]), 
             'o-', color='red', linewidth=2, label='Cumulative')
    plt.axhline(y=0.1, color='red', linestyle='--', label='Threshold (10%)')
    plt.title(f'{title_prefix} Scree Plot', fontsize=15, fontweight='bold')
    plt.xlabel('Principal Component', fontsize=12)
    plt.ylabel('Explained Variance Ratio', fontsize=12)
    plt.grid(True, alpha=0.3)
    plt.xticks(components)
    plt.legend(loc='best')
    
    # Add scree plot to document
    add_figure_to_doc(plt.gcf(), f'{title_prefix} - Scree Plot', 
                      f'{title_prefix.lower().replace(" ", "_")}_scree_plot')
    
    # Calculate cumulative explained variance
    cum_explained_variance = np.cumsum(pca.explained_variance_ratio_)
    
    # Create a table of explained variance
    var_table = []
    for i, var in enumerate(pca.explained_variance_ratio_[:10], 1):
        var_table.append([f"PC{i}", var, cum_explained_variance[i-1]])
    
    # Add variance table to document
    add_table_to_doc(var_table, 
                     ["Component", "Explained Variance", "Cumulative Explained Variance"],
                     f"{title_prefix} - Explained Variance by Component")
    
    # Factor Loadings with 2 components
    pca_2 = PCA(n_components=2)
    pca_2.fit(X_scaled)
    
    # Create a DataFrame for loadings
    loadings = pd.DataFrame(pca_2.components_.T, 
                          columns=['Component 1', 'Component 2'], 
                          index=variables)
    
    # Add loadings table to document
    loadings_table = []
    for var, row in loadings.iterrows():
        loadings_table.append([var, row['Component 1'], row['Component 2']])
    
    add_table_to_doc(loadings_table,
                     ["Variable", "Component 1 (Big Data/Automation)", "Component 2 (Human Judgment)"],
                     f"{title_prefix} - Component Loadings")
    
    # Visualize loadings
    plt.figure(figsize=(12, 8))
    plt.scatter(loadings['Component 1'], loadings['Component 2'], 
               s=100, color=colors[0], alpha=0.7)
    
    # Plot vectors for each variable
    for i, (var, loading) in enumerate(loadings.iterrows()):
        plt.arrow(0, 0, loading['Component 1'], loading['Component 2'], 
                  head_width=0.02, head_length=0.02, fc='blue', ec='blue', alpha=0.5)
        plt.text(loading['Component 1']*1.1, loading['Component 2']*1.1, var, 
                 color='darkblue', ha='center', va='center', fontsize=9)
    
    # Add circle
    circle = plt.Circle((0, 0), 1, facecolor='none', edgecolor='gray', alpha=0.5)
    plt.gca().add_patch(circle)
    
    plt.axhline(y=0, color='k', linestyle='-', alpha=0.3)
    plt.axvline(x=0, color='k', linestyle='-', alpha=0.3)
    plt.grid(alpha=0.3)
    plt.xlim(-1.2, 1.2)
    plt.ylim(-1.2, 1.2)
    plt.xlabel('Component 1 (Big Data/Automation)', fontsize=12)
    plt.ylabel('Component 2 (Human Judgment)', fontsize=12)
    plt.title(f'{title_prefix} - Component Loadings Plot', fontsize=15)
    
    # Add loadings plot to document
    add_figure_to_doc(plt.gcf(), f'{title_prefix} - Component Loadings Plot', 
                      f'{title_prefix.lower().replace(" ", "_")}_loadings_plot')
    
    # PCA Scatter Plot
    X_pca = pca_2.transform(X_scaled)
    plt.figure(figsize=(10, 8))
    sc = plt.scatter(X_pca[:, 0], X_pca[:, 1], alpha=0.7, s=50,
                    c=X_pca[:, 0]*X_pca[:, 1], cmap='viridis')
    plt.colorbar(sc, label='PC1 × PC2')
    plt.xlabel('Principal Component 1 (Big Data/Automation)', fontsize=12)
    plt.ylabel('Principal Component 2 (Human Judgment)', fontsize=12)
    plt.title(f'{title_prefix} - PCA Scatter Plot', fontsize=15)
    plt.grid(alpha=0.3)
    
    # Add scatter plot to document
    add_figure_to_doc(plt.gcf(), f'{title_prefix} - PCA Scatter Plot', 
                      f'{title_prefix.lower().replace(" ", "_")}_scatter_plot')
    
    # Return PCA results
    pca_df = pd.DataFrame(X_pca, columns=[f'{title_prefix}Factor1', f'{title_prefix}Factor2'], 
                         index=X.index)
    return pca_df, pca_2, loadings

# Perform PCA for comfortableness
com_pca_df, com_pca, com_loadings = perform_pca_analysis(data, com_vars, "Comfortableness")

# Perform PCA for capability
cap_pca_df, cap_pca, cap_loadings = perform_pca_analysis(data, cap_vars, "Capability")

# Create final DataFrame with the extracted factors
add_section_title("FACTOR SCORES AND STATISTICS")

# First, reset the index of the PCA DataFrames to merge them correctly
com_pca_df = com_pca_df.reset_index()
cap_pca_df = cap_pca_df.reset_index()

# Rename the columns to the specified names
com_pca_df.columns = ['index', 'ScomFactor1', 'ScomFactor2']
cap_pca_df.columns = ['index', 'ScapFactor1', 'ScapFactor2']

# Merge the DataFrames
final_df = pd.merge(com_pca_df, cap_pca_df, on='index')
final_df.set_index('index', inplace=True)

# Save the extracted factors to a CSV file
final_df.to_csv('outputs/pca_factors.csv')
doc.add_paragraph(f"Factor components saved to 'outputs/pca_factors.csv' with {final_df.shape[0]} samples")

# Add preview of factor scores
add_table_to_doc(final_df.head(10).values, 
                 final_df.columns, 
                 "Preview of Factor Scores (First 10 rows)")

# Add descriptive statistics
stats_df = final_df.describe().round(4)
add_table_to_doc(stats_df.values, 
                 stats_df.index, 
                 "Descriptive Statistics for Factor Scores")

# Function to interpret components
def interpret_components(loadings, title):
    add_section_title(f"{title.upper()} COMPONENTS INTERPRETATION")
    
    # Component 1
    comp1_pos = loadings.nlargest(5, 'Component 1').index.tolist()
    comp1_neg = loadings.nsmallest(5, 'Component 1').index.tolist()
    
    component1_table = [
        ["Highest positive loadings", ", ".join(comp1_pos)],
        ["Highest negative loadings", ", ".join(comp1_neg)]
    ]
    
    add_table_to_doc(component1_table,
                     ["Loadings Type", "Variables"],
                     f"Component 1 - Big Data/Automation")
    
    # Component 2
    comp2_pos = loadings.nlargest(5, 'Component 2').index.tolist()
    comp2_neg = loadings.nsmallest(5, 'Component 2').index.tolist()
    
    component2_table = [
        ["Highest positive loadings", ", ".join(comp2_pos)],
        ["Highest negative loadings", ", ".join(comp2_neg)]
    ]
    
    add_table_to_doc(component2_table,
                     ["Loadings Type", "Variables"],
                     f"Component 2 - Human Judgment")

# Interpret the components
interpret_components(com_loadings, "Comfortableness")
interpret_components(cap_loadings, "Capability")

# Add correlation analysis
add_section_title("FACTOR CORRELATIONS")

# Create correlation heatmap
plt.figure(figsize=(10, 8))
factor_corr = final_df.corr()
mask = np.triu(np.ones_like(factor_corr, dtype=bool))
sns.heatmap(factor_corr, mask=mask, cmap='coolwarm', annot=True, 
            fmt=".2f", linewidths=0.5, cbar_kws={"shrink": .8})
plt.title('Correlation Between PCA Factors', fontsize=16, fontweight='bold')

# Add correlation plot to document
add_figure_to_doc(plt.gcf(), 'Factor Correlations Heatmap', 'factor_correlations')

# Add correlation table to document
corr_table = []
for col1 in final_df.columns:
    row = []
    for col2 in final_df.columns:
        row.append(factor_corr.loc[col1, col2])
    corr_table.append([col1] + row)

add_table_to_doc(corr_table,
                 ['Factor'] + list(final_df.columns),
                 'Factor Correlation Matrix')

# Save the document
doc.save('outputs/PCA_Analysis_Report.docx')
print(colored("\nPCA analysis completed successfully! Report saved as 'PCA_Analysis_Report.docx'", 'green', attrs=['bold']))

