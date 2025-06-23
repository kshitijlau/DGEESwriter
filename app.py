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
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Summaries')
    processed_data = output.getvalue()
    return processed_data

# --- The RE-ENGINEERED Master Prompt Template (Version 5.1 - Syntax Fix) ---
def create_master_prompt(person_name, pronoun, person_data):
    """
    Dynamically creates the new, highly-constrained prompt for the Azure OpenAI API.
    VERSION 5.1: Corrects a syntax error in the f-string definition. The logic is identical to V5.

    Args:
        person_name (str): The first name of the person being evaluated.
        pronoun (str): The correct pronoun for the person (e.g., 'He' or 'She').
        person_data (str): A string representation of the person's scores and competencies.

    Returns:
        str: A fully constructed prompt ready to be sent to the AI model.
    """
    # SYNTAX FIX: The f-string has been carefully reconstructed to avoid parsing errors.
    prompt_text = f"""
You are an elite talent management consultant from a top-tier firm, known for writing insightful, nuanced, and actionable executive summaries in flawless British English. Your writing is strategic and indistinguishable from a human expert.

## Core Objective
Synthesize the provided competency data for {person_name} into a detailed, two-part executive summary.

## Input Data for {person_name}
<InputData>
{person_data}
</InputData>

## ---------------------------------------------
## CRITICAL DIRECTIVES FOR SUMMARY STRUCTURE & TONE
## ---------------------------------------------

1.  **Structure for Mixed-Score Profiles (Most Common Case):**
    * You MUST write a **two-paragraph summary**.
    * **Paragraph 1 (The Strengths Narrative):** This paragraph weaves the candidate's strengths into a fluid, professional story. It should describe *how* they demonstrated their capabilities. Start with a holistic opening sentence.
    * **Paragraph 2 (The Development Plan):** This paragraph details the growth opportunities. It must be constructive and provide specific, actionable suggestions. Start this paragraph with a clear transition phrase like "However," or "To strengthen..."
    * **High-Scorer Profile (No scores of 1 or 2):** In this special case, you can use a single, extended paragraph focusing entirely on the strengths narrative.

2.  **Opening Sentence Protocol (NEW):**
    * **AVOID** the rigid format: "{person_name} evidenced a strong ability to..."
    * **INSTEAD,** start with a more holistic, professional, and narrative opening that introduces the candidate's key strengths.
    * **Example Opening:** "Ebrahim demonstrated a committed and well-intentioned approach, drawing on his operational experience and strong sense of responsibility to driving results." OR "Mahra demonstrated strengths especially in strategic thinking and talent development, with growing capability in driving innovation."

3.  **The Development Plan (Paragraph 2) - CRITICAL DETAIL:**
    * Your development advice must be **highly specific and actionable**.
    * For each development area, clearly explain *how* the person can improve.
    * Use lead-in phrases to structure your points: "To strengthen {pronoun} contribution...", "{pronoun} may also consider enhancing...", "{pronoun} decision-making could be strengthened by...", "To drive results, {pronoun} may consider..."
    * **Incorrect (Generic):** "He needs to improve decision making."
    * **Correct (Specific & Actionable):** "His decision-making could be strengthened by improving risk forecasting and systematically evaluating long-term options."

4.  **Closing Statement (NEW):**
    * For mixed-score profiles, conclude the summary with a brief, positive, and forward-looking statement.
    * This final sentence should summarize the candidate's potential if the development areas are addressed.
    * **Example Closing:** "With more focus on structured execution and outcome-driven planning, he has the potential to strengthen his influence and impact."

5.  **General Style Rules:**
    * Use clear, simple sentences. Aggressively avoid long sentences connected by many commas or conjunctions.
    * Describe behaviors, do not just label competencies.
    * ABSOLUTELY NO mention of scores, levels, or the assessment itself.
    * Use the candidate's first name once at the beginning, then the correct pronoun ({pronoun}) thereafter.

## ---------------------------------------------
## ANALYSIS OF HUMAN-WRITTEN EXAMPLES (INTERNALIZE THIS LOGIC)
## ---------------------------------------------

**Example 1: Ebrahim (Mixed Profile)**
* **Analysis:** This is the gold standard. It uses a two-paragraph structure.
* **Paragraph 1 Logic:** It opens with a narrative sentence, then describes his strengths (responsibility, collaboration, innovation) by explaining his actions. It tells a story.
* **Paragraph 2 Logic:** It starts with "However," to transition. It then gives very specific advice for different competencies: for planning ("...limited evidence of structured planning..."), for decision-making ("...tended to lack depth, with minimal reference to data analysis...").
* **Closing Logic:** It ends with a positive, forward-looking summary of his potential. "Ebrahim brought a positive mindset... With more focus... he has the potential to strengthen his influence..."

**Example 2: Mahra (Mixed Profile)**
* **Analysis:** Also a two-paragraph structure, with highly structured development advice.
* **Paragraph 1 Logic:** Opens by naming her key strengths ("strategic thinking and talent development"). It then provides evidence for each, such as "...aligning initiatives with organisational goals, applying benchmarks..." and "...fostered learning by mentoring peers..."
* **Paragraph 2 Logic:** Uses specific lead-in phrases for each development point ("To strengthen her strategic contribution...", "Her decision-making could be strengthened by..."). Each piece of advice is a concrete action ("...clarify any cost constraints and define more detailed budget plans.").

## ---------------------------------------------
## FINAL INSTRUCTIONS
## ---------------------------------------------

Now, process the provided competency data for {person_name}. Determine the profile type. For mixed profiles, create a two-paragraph summary following all the new rules. For high-scorer profiles, create a single detailed paragraph. The total word count should be **between 250 and 280 words** to ensure comprehensive detail. Be descriptive, specific, and write in the sophisticated, nuanced style of the examples provided.
"""
    return prompt_text

# --- API Call Function for Azure OpenAI ---
def generate_summary_azure(prompt, api_key, endpoint, deployment_name):
    """
    Calls the Azure OpenAI API to generate a summary.
    """
    try:
        client = openai.AzureOpenAI(
            api_key=api_key,
            azure_endpoint=endpoint,
            api_version="2024-02-01"
        )
        response = client.chat.completions.create(
            model=deployment_name,
            messages=[
                {"role": "system", "content": "You are an elite talent management consultant from a top-tier firm, known for writing insightful, nuanced, and actionable executive summaries in flawless British English."},
                {"role": "user", "content": prompt}
            ],
            temperature=0.25,
            max_tokens=800,
            top_p=1.0,
            frequency_penalty=0.4,
            presence_penalty=0.0
        )
        return response.choices[0].message.content.strip()
    except Exception as e:
        st.error(f"An error occurred while contacting Azure OpenAI: {e}")
        return None

# --- Streamlit App Main UI ---
st.set_page_config(page_title="DGE Executive Summary Generator v5.1", layout="wide")

st.title("📄 DGE Executive Summary Generator (V5.1)")
st.markdown("""
This application generates professional executive summaries based on leadership competency scores. **Version 5.1 incorporates advanced, human-expert examples for a more nuanced, two-paragraph output.**
1.  **Set up your secrets**. Enter your Azure OpenAI credentials in the Streamlit Cloud app settings.
2.  **Download the Sample Template**. The format requires a `gender` column (enter 'M' for Male, 'F' for Female).
3.  **Upload your completed Excel file**.
4.  **Click 'Generate Summaries'** to process the file.
""")

# --- Create and provide a sample file for download ---
sample_data = {
    'email': ['jane.doe@example.com', 'john.roe@example.com'],
    'first_name': ['Jane', 'John'],
    'last_name': ['Doe', 'Roe'],
    'gender': ['F', 'M'],
    'level': ['Director', 'Manager'],
    'Strategic Thinker': [4, 2],
    'Impactful Decision Maker': [5, 3],
    'Effective Collaborator': [2, 5],
    'Talent Nurturer': [4, 2],
    'Results Driver': [3, 4],
    'Customer Advocate': [2, 4],
    'Transformation Enabler': [3, 1],
    'Innovation Catalyst': [1, 3]
}
sample_df = pd.DataFrame(sample_data)
sample_excel_data = to_excel(sample_df)

st.download_button(
    label="📥 Download Sample Template File (V5.1)",
    data=sample_excel_data,
    file_name="dge_summary_template_v5.1.xlsx",
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
            
            identifier_cols = ['email', 'first_name', 'last_name', 'level', 'gender']
            all_known_competencies = [
                'Strategic Thinker', 'Impactful Decision Maker', 'Effective Collaborator',
                'Talent Nurturer', 'Results Driver', 'Customer Advocate',
                'Transformation Enabler', 'Innovation Catalyst'
            ]
            competency_columns = [col for col in df.columns if col in all_known_competencies]
            
            if 'gender' not in df.columns:
                st.error("Error: The uploaded file is missing the required 'gender' column. Please download the new template and try again.")
                st.stop()

            generated_summaries = []
            progress_bar = st.progress(0)
            
            for i, row in df.iterrows():
                first_name = row['first_name']
                gender_input = str(row['gender']).upper()
                pronoun = 'They'
                if gender_input == 'M':
                    pronoun = 'He'
                elif gender_input == 'F':
                    pronoun = 'She'
                else:
                    st.warning(f"Invalid or missing gender '{row['gender']}' for {first_name}. Defaulting to pronoun 'They'.")

                st.write(f"Processing summary for: {first_name}...")
                
                scores_data = []
                for competency in competency_columns:
                    if competency in row and pd.notna(row[competency]):
                        scores_data.append(f"- {competency}: {int(row[competency])}")
                person_data_str = "\n".join(scores_data)

                prompt = create_master_prompt(first_name, pronoun, person_data_str)
                summary = generate_summary_azure(prompt, azure_api_key, azure_endpoint, azure_deployment_name)
                
                if summary:
                    generated_summaries.append(summary)
                    st.success(f"Successfully generated summary for {first_name}.")
                else:
                    generated_summaries.append("Error: Failed to generate summary.")
                    st.error(f"Failed to generate summary for {first_name}.")

                progress_bar.progress((i + 1) / len(df))

            if generated_summaries:
                st.balloons()
                st.subheader("Generated Summaries (V5.1)")
                
                output_df = df.copy()
                output_df['Executive Summary'] = generated_summaries
                
                st.dataframe(output_df)
                
                results_excel_data = to_excel(output_df)
                st.download_button(
                    label="📥 Download V5.1 Results as Excel",
                    data=results_excel_data,
                    file_name="Generated_Executive_Summaries_V5.1.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

    except Exception as e:
        st.error(f"An error occurred: {e}")
