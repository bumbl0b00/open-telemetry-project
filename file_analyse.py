import pandas as pd
import numpy as np
def load_excel(file_path):
    try:
        data = pd.read_excel(file_path, sheet_name=input, engine='openpyxl')  # Load all sheets
        return data
    except Exception as e:
        print(f"Error loading Excel file: {e}")
        return None
def analyze_data(dataframes):
    analysis_results = {}
    for sheet_name, df in dataframes.items():
        print(f"Analyzing sheet: {sheet_name}")
        sheet_analysis = {}
        # Basic information
        sheet_analysis['Shape'] = df.shape
        sheet_analysis['Columns'] = df.columns.tolist()
        sheet_analysis['Missing Values'] = df.isnull().sum().to_dict()    
        # Data type analysis
        sheet_analysis['Data Types'] = df.dtypes.to_dict() 
        # Identify duplicates
        sheet_analysis['Duplicate Rows'] = df.duplicated().sum()
        # Summary statistics
        try:
            sheet_analysis['Summary Statistics'] = df.describe(include='all').to_dict()
        except:
            sheet_analysis['Summary Statistics'] = "Unable to generate summary statistics."
        # Patterns and anomalies
        anomalies = {}
        for column in df.select_dtypes(include=np.number):  # Analyze numeric columns
            mean, std = df[column].mean(), df[column].std()
            anomalies[column] = df[(df[column] < mean - 3 * std) | (df[column] > mean + 3 * std)].index.tolist()
        sheet_analysis['Anomalies'] = anomalies
        analysis_results[sheet_name] = sheet_analysis
    return analysis_results
# Save analysis to a file
def save_analysis(results, output_file):
    try:
        with open(output_path, 'w') as f:
            for sheet_name, analysis in results.items():
                f.write(f"Sheet: {sheet_name}\n")
                for key, value in analysis.items():
                    f.write(f"{key}:\n{value}\n\n")
        print(f"Analysis saved to {output_file}")
    except Exception as e:
        print(f"Error saving analysis: {e}")
if __name__ == "__main__":
    file_path ="input.xlsx"  # Replace with your file path
    output_file = "output.txt"
    # Load and analyze data
    dataframes = load_excel(file_path)
    if dataframes:
        analysis_results = analyze_data(dataframes)
        save_analysis(analysis_results, output_file)
