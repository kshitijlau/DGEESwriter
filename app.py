import streamlit as st
import pandas as pd
import openai
import io

# --- Helper Function to convert DataFrame to Excel in memory ---
def to_excel(df):
    """
    Converts a pandas DataFrame to an Excel file in memory (bytes).
    This function is used to prepare the final output for download.
    """
    output = io.BytesIO()
    # Use 'xlsxwriter' engine for better compatibility and features.
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Summaries')
    processed_data = output.getvalue()
    return processed_data

# --- The Master Prompt Template ---
# This prompt is engineered based on the DGE guidelines and human-written examples.
def create_master_prompt(person_name, person_data):
    """
    Dynamically creates the detailed, expert-level prompt for the Azure OpenAI API.
    This function structures all the rules, examples, and data into a single request.

    Args:
        person_name (str): The first name of the person being evaluated.
        person_data (str): A string representation of the person's scores and competencies.

    Returns:
        str: A fully constructed prompt ready to be sent to the AI model.
    """
    return f"""
You are a senior talent assessment consultant specializing in writing executive summaries for leadership assessment centers. Your writing style adheres strictly to British English and maintains a professional, constructive, and positive tone. Your primary skill is to weave competency scores into a fluid, professional narrative.

## Core Objective
Your task is to synthesize the provided competency scores for {person_name} into a formal, narrative-based executive summary.

## Input Data for {person_name}
{person_data}

## Primary Directives & Rules

1.  **Writing Style: Describe, Don't Label.**
    * **CRITICAL RULE:** Do not simply name the competency. Instead, describe the behaviors that were observed, using the competency name as your guide. The summary must be a story of the person's actions.
    * **Incorrect (Labeling):** "{person_name} demonstrated Strategic Thinking."
    * **Correct (Describing):** "{person_name} consistently displayed the ability to anticipate future challenges and proactively aligned his actions with the organisation’s long-term goals and vision."

2.  **Structure Based on Profile Type:**
    * **High-Scorer Profile (No scores of 1 or 2):** Write a single, flowing narrative paragraph. Start with a general opening like "{person_name} displayed strengths in multiple competencies assessed." Then, seamlessly describe the observed strengths for all competencies.
    * **Mixed-Score Profile (Contains scores of 1 or 2):** Use an integrated narrative structure. Begin with the main strengths, then transition smoothly to growth opportunities using phrases like "...however, he may need to..." or "To further develop his competence...". The goal is a single, cohesive story, not two separate lists.

3.  **Tense and Perspective:**
    * Use **past tense** for all observed behaviors (e.g., "He demonstrated...", "She showcased...").
    * Use **present or future tense** for all development suggestions ("He could enhance...", "She would benefit from...").
    * Use the candidate's first name, "{person_name}", only at the beginning. Use pronouns (He/She) thereafter.

4.  **Actionable Development (For Mixed Profiles Only):**
    * For each growth opportunity (scores 1, 2), you MUST provide a **specific, actionable, and context-relevant suggestion**. Avoid generic advice.
    * **Incorrect (Generic):** "He could benefit from a course on decision-making."
    * **Correct (Specific):** "He may strengthen his decision-making skills by conducting in-depth analysis, comprehensive risk assessments and offering broader alternatives while gaining buy-in beforehand, to minimise resistance."

5.  **General Rules:**
    * **Word Count:** The entire summary must not exceed 280 words.
    * **Coverage:** All 8 competencies must be addressed, either as a described strength or a growth opportunity.
    * **Language:** British English. Professional synonyms for "good/bad". Correct punctuation.
    * **ABSOLUTE RULE:** Never mention the numerical scores in the final text.

## Deconstruction of Examples (Internalize this Logic)

**Example 1: Khalifa (High-Scorer Profile)**
* **Analysis:** This profile has no low scores. The summary is a single, positive narrative.
* **Correct Logic to Emulate:**
    1.  **Opening:** Start with a broad, confident opening: "Khalifa displayed strengths in multiple competencies assessed."
    2.  **Behavioral Narrative:** Describe the behaviors for each competency seamlessly. Instead of saying "Strategic Thinker," write "He consistently displayed the ability to anticipate future challenges...". This is the gold standard.

**Example 2: Rashed (Mixed-Score Profile - Integrated Narrative)**
* **Analysis:** This profile has strengths and multiple development needs. The ideal summary is a single, integrated paragraph.
* **Correct Logic to Emulate:**
    1. **Opening:** Start by highlighting the primary strength: "Rashed presented as a results driver."
    2. **Integrated Flow:** Describe the other demonstrated strengths first. Then, transition smoothly into development areas using phrases like "...however, he may need to foster a more customer-centric approach..." and "To further develop his competence, Rashed may benefit from...". The entire summary feels like one cohesive professional assessment.
    3. **Specific Actions:** The development actions are highly specific and business-relevant (e.g., "analysing latest industry and market trends to identify further growth opportunities and develop solid strategies, rather than focusing on operational tactics.").

**Example 3: Khasiba (Mixed-Score Profile - Balanced Feedback)**
* **Analysis:** This profile has clear strengths and some 'competent' areas (score 3) that can be enhanced.
* **Correct Logic to Emulate:**
    1. **Opening:** Start by naming the top 1-2 strengths: "Khasiba displayed strengths in strategic thinking and driving results."
    2. **Balanced Feedback:** For 'competent' skills, describe the observed positive behavior (past tense), then immediately provide a future-focused development action. Example: "She showcased the ability to work effectively with peers... As a next step, she could focus on enhancing her influence in group settings..."

## Final Instruction for {person_name}
Now, process the provided competency scores for {person_name}. First, determine if it is a High-Scorer or Mixed-Score profile. Then, write a concise executive summary under 280 words, adhering to all the rules above. Focus on describing behaviors, not labeling competencies. Use the correct narrative structure for the profile type and NEVER mention the scores.
"""

# --- API Call Function for Azure OpenAI ---
def generate_summary_azure(prompt, api_key, endpoint, deployment_name):
    """
    Calls the Azure OpenAI API to generate a summary.
    """
    try:
        # Initialize the AzureOpenAI client with credentials from secrets
        client = openai.AzureOpenAI(
            api_key=api_key,
            azure_endpoint=endpoint,
            api_version="2024-02-01"  # A common, stable API version
        )
        
        # Make the API call to the chat completions endpoint
        response = client.chat.completions.create(
            model=deployment_name,  # Use the deployment name for the model
            messages=[
                {"role": "system", "content": "You are a senior talent assessment consultant and report writer following British English standards."},
                {"role": "user", "content": prompt}
            ],
            temperature=0.2,
            max_tokens=600, # Sufficient for the summary length
            top_p=1.0,
            frequency_penalty=0.0,
            presence_penalty=0.0
        )
        return response.choices[0].message.content.strip()
    except Exception as e:
        st.error(f"An error occurred while contacting Azure OpenAI: {e}")
        return None

# --- Streamlit App Main UI ---
st.set_page_config(page_title="DGE Executive Summary Generator", layout="wide")

st.title("📄 DGE Executive Summary Generator")
st.markdown("""
This application generates professional executive summaries based on leadership competency scores, adhering to DGE guidelines.
1.  **Set up your secrets**. Enter your Azure OpenAI credentials in the Streamlit Cloud app settings.
2.  **Download the Sample Template** to see the required Excel format.
3.  **Upload your completed Excel file**.
4.  **Click 'Generate Summaries'** to process the file.
""")

# --- Create and provide a sample file for download ---
sample_data = {
    'email': ['jane.doe@example.com', 'john.roe@example.com'],
    'first_name': ['Jane', 'John'],
    'last_name': ['Doe', 'Roe'],
    'level': ['Director', 'Manager'],
    'Strategic Thinker': [4, 2],
    'Impactful Decision Maker': [5, 5],
    'Effective Collaborator': [2, 3],
    'Talent Nurturer': [4, 2],
    'Results Driver': [3, 4],
    'Customer Advocate': [2, 4],
    'Transformation Enabler': [3, 1],
    'Innovation Catalyst': [1, 3] # Assuming 8th competency is Innovation Catalyst
}
sample_df = pd.DataFrame(sample_data)
sample_excel_data = to_excel(sample_df)


st.download_button(
    label="📥 Download Sample Template File",
    data=sample_excel_data,
    file_name="dge_summary_template.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

st.divider()

# --- File Uploader ---
uploaded_file = st.file_uploader("Upload your completed Excel file here", type="xlsx")

if uploaded_file is not None:
    try:
        df = pd.read_excel(uploaded_file)
        st.success("Excel file loaded successfully. Ready to generate summaries.")
        st.dataframe(df.head())

        if st.button("Generate Summaries", key="generate"):
            try:
                azure_api_key = st.secrets["azure_openai"]["api_key"]
                azure_endpoint = st.secrets["azure_openai"]["endpoint"]
                azure_deployment_name = st.secrets["azure_openai"]["deployment_name"]
            except (KeyError, FileNotFoundError):
                st.error("Azure OpenAI credentials not found. Please configure them in your Streamlit secrets.")
                st.stop()
            
            # --- Identify identifier and competency columns ---
            identifier_cols = ['email', 'first_name', 'last_name', 'level']
            # Assume all other columns are competency scores
            competency_columns = [col for col in df.columns if col not in identifier_cols]
            
            generated_summaries = []
            progress_bar = st.progress(0)
            
            # Iterate through each row of the DataFrame
            for i, row in df.iterrows():
                # **MODIFICATION**: Use only the 'first_name' for the summary
                first_name = row['first_name']
                st.write(f"Processing summary for: {first_name}...")
                
                scores_data = []
                for competency in competency_columns:
                    # Check if the competency column exists and the value is not null
                    if competency in row and pd.notna(row[competency]):
                        scores_data.append(f"- {competency}: {row[competency]}")
                person_data_str = "\n".join(scores_data)

                # Create and run the prompt for the current person using their first name
                prompt = create_master_prompt(first_name, person_data_str)
                summary = generate_summary_azure(prompt, azure_api_key, azure_endpoint, azure_deployment_name)
                
                if summary:
                    generated_summaries.append(summary)
                    st.success(f"Successfully generated summary for {first_name}.")
                else:
                    generated_summaries.append("Error: Failed to generate summary.")
                    st.error(f"Failed to generate summary for {first_name}.")

                # Update the progress bar
                progress_bar.progress((i + 1) / len(df))

            if generated_summaries:
                st.balloons()
                st.subheader("Generated Summaries")
                
                # Append the new summary column to the original dataframe
                output_df = df.copy()
                output_df['Executive Summary'] = generated_summaries
                
                st.dataframe(output_df)
                
                # Provide download for the modified dataframe with results
                results_excel_data = to_excel(output_df)
                st.download_button(
                    label="📥 Download Results as Excel",
                    data=results_excel_data,
                    file_name="Generated_Executive_Summaries.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

    except Exception as e:
        st.error(f"An error occurred: {e}")
