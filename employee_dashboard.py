import pandas as pd
import os
from glob import glob
import streamlit as st

# === Function to read daily report files ===
def read_daily_reports(folder):
    daily_data = []
    files = glob(os.path.join(folder, 'Daily report_*.xlsx'))
    for file in files:
        # Extract team member name from filename
        filename = os.path.basename(file)
        parts = filename.replace('.xls', '').split('_')
        team_member = f"{parts[2]} {parts[3]}"
        
        # Read Excel
        df = pd.read_excel(file)
        df.columns = [col.strip() for col in df.columns]  # Clean column names
        df['Team Member'] = team_member
        
        # Filter for passed interviews
        filtered = df[
            (df['Interview Status'].str.strip().str.lower() == 'yes') &
            (df['Remark'].str.strip().str.lower() == 'pass')
        ]
        daily_data.append(filtered)
    return pd.concat(daily_data, ignore_index=True)

# === Function to read new employee file ===
def read_new_employees(file_path):
    df = pd.read_excel(file_path)
    df.columns = [col.strip() for col in df.columns]  # Clean column names
    return df

# === Match new employees with team members ===
def match_employees(interview_df, new_emp_df):
    result = []
    for _, row in new_emp_df.iterrows():
        match = interview_df[
            (interview_df['Candidate Name'].str.strip().str.lower() == row['Employee Name'].strip().lower()) &
            (interview_df['Role'].str.strip().str.lower() == row['Role'].strip().lower())
        ]
        if not match.empty:
            team_member = match.iloc[0]['Team Member']
            result.append({
                'Employee Name': row['Employee Name'],
                'Join Date': row['Join Date'],
                'Role': row['Role'],
                'Team Member': team_member
            })
    return pd.DataFrame(result)

# === Main App ===
def main():
    st.title("New Employees Matched to Team Members")
    
    st.info("Make sure your files are in a folder named `data`")
    
    # File paths
    daily_folder = './data'
    new_emp_file = './data/New Employee_202502.xlsx'
    
    # Read data
    interview_df = read_daily_reports(daily_folder)
    new_emp_df = read_new_employees(new_emp_file)
    final_df = match_employees(interview_df, new_emp_df)

    st.subheader("Matched New Employees")
    st.dataframe(final_df)

    # Export to Excel
    if st.button("Export to Excel"):
        final_df.to_excel("Final_Employee_Team_List.xlsx", index=False)
        st.success("Exported to Final_Employee_Team_List.xlsx")

if __name__ == "__main__":
    main()
