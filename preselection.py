import pandas as pd
import numpy as np
import streamlit as st
from datetime import datetime
import io

# Streamlit file upload component
employment_file = st.file_uploader("Upload Employment data file (Excel format)", type="xlsx")
education_file = st.file_uploader("Upload Education data file (Excel format)", type="xlsx")

# Ensure processing happens only after both files are uploaded
if employment_file is not None and education_file is not None:
    # Read the uploaded Excel files
    df = pd.read_excel(employment_file)
    edu_df = pd.read_excel(education_file)

    # Display the first few rows of Employment and Education data
    st.write("Employment Data:")
    st.write(df.head())  # Display first few rows of Employment data

    st.write("Education Data:")
    st.write(edu_df.head())  # Display first few rows of Education data

    def calc_work_length(row):
        if pd.isnull(row['Start Date']) or pd.isnull(row['End Date']):
            return ''
        start = pd.to_datetime(row['Start Date'], dayfirst=True, errors='coerce')
        end = pd.to_datetime(row['End Date'], dayfirst=True, errors='coerce')
        if pd.isnull(start) or pd.isnull(end):
            return ''
        years = end.year - start.year - ((end.month, end.day) < (start.month, start.day))
        months = (end.year - start.year) * 12 + end.month - start.month
        months = months % 12
        result = ''
        if years > 0:
            result += f'{years} year(s) '
        if months > 0:
            result += f'{months} month(s)'
        return result.strip()

    # Calculate work length for each entry
    df['Work Length'] = df.apply(calc_work_length, axis=1)
    df.dropna(subset=['Start Date', 'End Date'], how='any', inplace=True)

    # Concatenate work experience information
    df['Previous working experience'] = (
        df['Employer'].fillna('') + ', ' +
        df['Country'].fillna('') + ', ' +
        df['Job Title'].fillna('') + ', ' +
        df['Work Length'].replace('', 'N/A')
    ).str.replace(r'(, ){2,}', ', ', regex=True).str.strip(', ')

    # Summarize work experience information
    summary_df = df.groupby(['First Name', 'Last Name']).agg({
        'Previous working experience': lambda x: '\n'.join(x),
        'Grade (for UN staff)': 'first'  # Retain grade
    }).reset_index()

    summary_df.rename(columns={'Grade (for UN staff)': 'Current Grade'}, inplace=True)
    summary_df['Previous working experience'] = summary_df['Previous working experience'].str.replace(r',\s*,', ',', regex=True)
    summary_df['Previous working experience'] = summary_df['Previous working experience'].str.replace(r'\s+,', ',', regex=True)

    # Calculate age
    edu_df['Age'] = edu_df['Date of Birth'].apply(
        lambda x: (datetime.now() - pd.to_datetime(x, dayfirst=True)).days // 365 if not pd.isnull(x) else np.nan
    )

    # Handle Institution and Internal/External values
    edu_df['Institution'] = edu_df.apply(
        lambda row: row['Other Institution'] if row['Institution'] == '* Other – Cannot find my school in the list' else row['Institution'],
        axis=1
    )

    edu_df['Internal/External'] = edu_df['Is Internal'].map({'Yes': 'Internal', 'No': 'External'})

    # Map education levels
    edu_level_map = {
        "5 Master Degree": "MA",
        "6 PhD Doctorate Degree": "PHD",
        "4 Bachelor's Degree": "BA",
        "3 Technical Diploma": "TD",
        "1 Non-Degree Programme": "Non-Degree",
        "2 High School diploma": "High School"
    }
    edu_df['Education Level'] = edu_df['Education Level'].map(edu_level_map).fillna(edu_df['Education Level'])

    # Calculate the highest education level
    level_score = {
        "PHD": 5,
        "MA": 4,
        "BA": 3,
        "TD": 2,
        "Non-Degree": 0,
        "High School": 1
    }
    edu_df['Level Score'] = edu_df['Education Level'].map(level_score)
    edu_df['Max Level Score'] = edu_df.groupby(['First Name', 'Last Name'])['Level Score'].transform('max')
    edu_df['Keep'] = np.where(edu_df['Level Score'] == edu_df['Max Level Score'], 'Keep', '')

    # Handle Main subject
    edu_df['Main subject'] = edu_df['Main subject'].astype(str).str.split(',').str[0]

    # Filter to keep the highest education level
    edu_keep_df = edu_df[edu_df['Keep'] == 'Keep'].copy()

    # Concatenate education information
    edu_keep_df['Highest Education'] = (
        edu_keep_df['Education Level'].fillna('') + ' - ' +
        edu_keep_df['Institution'].fillna('') + ' (' +
        edu_keep_df['Country'].fillna('') + ') - ' +
        edu_keep_df['Main subject'].fillna('')
    )
    
    # Group by First Name, Last Name to merge multiple highest education records into one cell
    edu_summary_df = edu_keep_df.groupby(['First Name', 'Last Name']).agg({
        'Age': 'first',
        'Internal/External': 'first',
        'Geo Dist. Representation': 'first',
        'Primary Nationality': 'first',
        'Highest Education': lambda x: '\n'.join(x)
    }).reset_index()
    

    # Merge with the final summary data
    final_df = summary_df.merge(edu_summary_df, on=['First Name', 'Last Name'], how='left')

    # Adjust 'Geo Dist. Representation' if Internal
    final_df['Geo Dist. Representation'] = final_df.apply(
        lambda row: 'Non impact' if row['Internal/External'] == 'Internal' else row['Geo Dist. Representation'], axis=1
    )

    # Fill 'Current Grade' with 'N/A' where missing
    final_df['Current Grade'] = final_df['Current Grade'].fillna('N/A')

    # Reorder the final DataFrame
    final_df = final_df[[
        'First Name', 'Last Name', 'Age', 'Current Grade', 'Internal/External', 'Primary Nationality', 
        'Geo Dist. Representation', 'Previous working experience', 'Highest Education'
    ]]
    
    # Display the final cleaned data
    st.write("Final Cleaned Data:")
    st.write(final_df)
    
    # Create an in-memory buffer to save the DataFrame as Excel
    output = io.BytesIO()
    final_df.to_excel(output, index=False, engine='openpyxl')
    
    # Seek to the beginning of the BytesIO buffer
    output.seek(0)
    
    # Option to download the cleaned data
    st.download_button(
        label="Download Cleaned Data",
        data=output,
        file_name='final_cleaned_data.xlsx',
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
