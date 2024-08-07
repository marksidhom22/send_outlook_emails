import pandas as pd
from fuzzywuzzy import process
import re

# Load the CSV files with explicit column names and check the delimiters
pdf_files = pd.read_csv('pdf_files_with_grandfather_folders_1.csv')
directors = pd.read_csv('Directors of Graduate Studies.csv', delimiter=',')
assistants = pd.read_csv('Assistants DoGS.csv', delimiter=',')

# Print column names for verification
print("PDF Files Columns:", pdf_files.columns)
print("Directors Columns:", directors.columns)
print("Assistants Columns:", assistants.columns)

# Preprocess department names by removing parentheses and extra spaces
def preprocess_department_name(name):
    if pd.isnull(name):
        return ''
    # Remove content in parentheses and trim spaces
    return re.sub(r'\s*\(.*?\)', '', name).strip()

# Apply preprocessing to the department names
pdf_files['Department Name'] = pdf_files['Department Name'].apply(preprocess_department_name)
directors['Department Name'] = directors['Department Name'].apply(preprocess_department_name)
assistants['Department Name'] = assistants['Department Name'].apply(preprocess_department_name)

# Define a function to get the best match for a department name
def get_best_match(department_name, choices, score_cutoff=80):
    if pd.isnull(department_name) or department_name == '':
        return None
    result = process.extractOne(department_name, choices)
    if result and result[1] >= score_cutoff:
        return result[0]
    return None

# Get unique department names from directors and assistants
director_departments = directors['Department Name'].dropna().unique()
assistant_departments = assistants['Department Name'].dropna().unique()

# Add new columns to the pdf_files DataFrame
pdf_files['Director Name'] = None
pdf_files['Director Email'] = None
pdf_files['Director Department Name'] = None
pdf_files['Assistant Name'] = None
pdf_files['Assistant Email'] = None
pdf_files['Assistant Department Name'] = None

# Iterate through each row in pdf_files
for index, row in pdf_files.iterrows():
    department_name = row['Department Name']
    
    # Find the best match for directors
    best_director_match = get_best_match(department_name, director_departments)
    if best_director_match:
        director_info = directors[directors['Department Name'] == best_director_match].iloc[0]
        pdf_files.at[index, 'Director Name'] = director_info['Director Name']
        pdf_files.at[index, 'Director Email'] = director_info['Director Email']
        pdf_files.at[index, 'Director Department Name'] = director_info['Department Name']
    
    # Find the best match for assistants
    best_assistant_match = get_best_match(department_name, assistant_departments)
    if best_assistant_match:
        assistant_info = assistants[assistants['Department Name'] == best_assistant_match].iloc[0]
        pdf_files.at[index, 'Assistant Name'] = assistant_info['Assistant Name']
        pdf_files.at[index, 'Assistant Email'] = assistant_info['Assistant Email']
        pdf_files.at[index, 'Assistant Department Name'] = assistant_info['Department Name']

# Save the updated DataFrame to a new CSV file
pdf_files.to_csv('updated_pdf_files_with_grandfather_folders_1.csv', index=False)
