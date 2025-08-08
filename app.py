import streamlit as st
import pandas as pd
import random
from io import BytesIO

# --- CORE LOGIC FUNCTIONS ---

def find_lowest_competencies(row, competencies):
    """Identifies the two lowest scoring competencies for a candidate."""
    scores = {comp: row[comp] for comp in competencies if pd.notna(row[comp])}
    sorted_competencies = sorted(scores.items(), key=lambda item: item[1])
    
    lowest_1 = sorted_competencies[0][0] if len(sorted_competencies) > 0 else None
    lowest_2 = sorted_competencies[1][0] if len(sorted_competencies) > 1 else None
    
    return lowest_1, lowest_2

def get_random_tips(repo_df, competency):
    """Gets two random tips (70% and 20%) for a given competency."""
    repo_df.dropna(subset=['Competency Name'], inplace=True)
    
    comp_df = repo_df[repo_df['Competency Name'].str.strip().str.lower() == competency.strip().lower()]
    
    tips_70 = comp_df['70% Development Tips'].dropna().tolist()
    tips_20 = comp_df['20% Development Tips'].dropna().tolist()
    
    tip_70 = random.choice(tips_70) if tips_70 else "N/A"
    tip_20 = random.choice(tips_20) if tips_20 else "N/A"
    
    return tip_70, tip_20

# --- NEW, SIMPLIFIED OUTPUT FUNCTION ---
def generate_table_output(results):
    """Creates a single-sheet Excel table from the results."""
    if not results:
        return None

    # Convert the list of dictionaries to a pandas DataFrame
    df = pd.DataFrame(results)

    # Define the final column names and order
    column_mapping = {
        "Candidate Name": "Candidate Name",
        "Level": "Level",
        "Lowest Competency 1": "Focus Area 1",
        "Tip 1 (70%)": "Dev Action 70 (Area 1)",
        "Tip 1 (20%)": "Dev Action 20 (Area 1)",
        "Lowest Competency 2": "Focus Area 2",
        "Tip 2 (70%)": "Dev Action 70 (Area 2)",
        "Tip 2 (20%)": "Dev Action 20 (Area 2)"
    }
    
    # Rename columns to match the desired output
    df.rename(columns=column_mapping, inplace=True)
    
    # Ensure all desired columns are present and in the correct order
    final_df = df[list(column_mapping.values())]

    # Save the DataFrame to an in-memory Excel file
    output = BytesIO()
    final_df.to_excel(output, index=False, sheet_name='Development Plans')
    output.seek(0)
    return output

# --- STREAMLIT UI ---

st.set_page_config(page_title="Development Plan Generator", layout="wide")
st.title("Development Plan Generator ⚙️")

st.info("Upload your Excel files below. The app will generate a single-tab Excel report.")

col1, col2 = st.columns(2)

with col1:
    st.header("Candidate Data")
    uploaded_candidates_file = st.file_uploader("1. Upload Candidate Data (Excel)", type=["xlsx"])

    @st.cache_data
    def create_sample_candidate_file():
        sample_data = {
            'Candidate Name': ['John Doe', 'Jane Smith'], 'Level': ['Apply', 'Guide'],
            'Manages Stakeholders': [2.5, 4.1], 'Steers Change': [3.1, 4.2], 'Leads People': [1.8, 4.8],
            'Drives Results': [4.5, 3.2], 'Solves Challenges': [2.1, 4.3], 'Thinks Strategically': [3.3, 4.5]
        }
        df = pd.DataFrame(sample_data)
        output = BytesIO()
        df.to_excel(output, index=False, sheet_name='Candidates')
        output.seek(0)
        return output

    st.download_button(
        label="📥 Download Sample Candidate File", data=create_sample_candidate_file(),
        file_name="sample_candidate_data.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

with col2:
    st.header("Tip Repository")
    uploaded_repo_files = st.file_uploader(
        "2. Upload ALL THREE Tip Repositories (Apply, Guide, Shape Excel files)", 
        type=["xlsx"], 
        accept_multiple_files=True
    )
    
    @st.cache_data
    def create_sample_repo_file(level):
        """Creates a sample repository Excel file for a given level."""
        sample_data = {
            'Competency Name': ['Manages Stakeholders', 'Manages Stakeholders', 'Leads People', 'Leads People', ''],
            '70% Development Tips': [f'70% Tip A for {level}', f'70% Tip B for {level}', f'70% Tip C for {level}', f'70% Tip D for {level}', ''],
            '20% Development Tips': [f'20% Tip X for {level}', f'20% Tip Y for {level}', f'20% Tip Z for {level}', f'20% Tip W for {level}', '']
        }
        df = pd.DataFrame(sample_data)
        output = BytesIO()
        df.to_excel(output, index=False, sheet_name=level)
        output.seek(0)
        return output

    st.download_button(
        label="📥 Download Sample 'Apply' Repo", data=create_sample_repo_file('Apply'),
        file_name="sample_repository_apply.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key='apply'
    )
    st.download_button(
        label="📥 Download Sample 'Guide' Repo", data=create_sample_repo_file('Guide'),
        file_name="sample_repository_guide.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key='guide'
    )
    st.download_button(
        label="📥 Download Sample 'Shape' Repo", data=create_sample_repo_file('Shape'),
        file_name="sample_repository_shape.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key='shape'
    )

st.markdown("---")

if st.button("Generate Development Reports", type="primary"):
    if uploaded_candidates_file is None:
        st.error("⚠️ Please upload the Candidate Data Excel file.")
    elif len(uploaded_repo_files) != 3:
        st.error(f"⚠️ Please upload exactly 3 repository Excel files. You have uploaded {len(uploaded_repo_files)}.")
    else:
        with st.spinner('Processing...'):
            try:
                candidates_df = pd.read_excel(uploaded_candidates_file)
                
                repos = {}
                for f in uploaded_repo_files:
                    if 'apply' in f.name.lower():
                        repos['Apply'] = pd.read_excel(f)
                    elif 'guide' in f.name.lower():
                        repos['Guide'] = pd.read_excel(f)
                    elif 'shape' in f.name.lower():
                        repos['Shape'] = pd.read_excel(f)
                
                if len(repos) != 3:
                    st.error("❌ Could not identify 'Apply', 'Guide', and 'Shape' files from the filenames.")
                    st.stop()
                
                st.write("✅ Repositories loaded successfully!")

                COMPETENCIES = [
                    'Manages Stakeholders', 'Steers Change', 'Leads People', 
                    'Drives Results', 'Solves Challenges', 'Thinks Strategically'
                ]
                final_results = []
                
                for index, row in candidates_df.iterrows():
                    level = row['Level']
                    if level not in repos:
                        st.warning(f"Skipping candidate {row['Candidate Name']}: level '{level}' not found.")
                        continue
                    
                    repo_df = repos[level].copy() 
                    low_comp_1, low_comp_2 = find_lowest_competencies(row, COMPETENCIES)
                    
                    if not low_comp_1 or not low_comp_2:
                        st.warning(f"Skipping candidate {row['Candidate Name']}: could not determine two lowest competencies.")
                        continue
                    
                    tip1_70, tip1_20 = get_random_tips(repo_df, low_comp_1)
                    tip2_70, tip2_20 = get_random_tips(repo_df, low_comp_2)
                    
                    final_results.append({
                        "Candidate Name": row['Candidate Name'], "Level": level,
                        "Lowest Competency 1": low_comp_1, "Tip 1 (70%)": tip1_70, "Tip 1 (20%)": tip1_20,
                        "Lowest Competency 2": low_comp_2, "Tip 2 (70%)": tip2_70, "Tip 2 (20%)": tip2_20,
                    })

                if final_results:
                    # MODIFIED: Call the new table output function
                    excel_data = generate_table_output(final_results)
                    st.success("✅ Success! Your development plan table is ready for download.")
                    st.download_button(
                        label="📥 Download Excel Table", data=excel_data,
                        file_name="Development_Plans_Table.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                else:
                    st.warning("No candidates were processed. Please check your input files.")

            except Exception as e:
                st.error(f"An unexpected error occurred. Please check that your Excel files are formatted correctly. Error details: {e}")
