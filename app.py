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

# --- The RE-ENGINEERED Master Prompt Template (Version 6.1 - Confirmed One Paragraph) ---
def create_master_prompt(person_name, pronoun, person_data):
    """
    Dynamically creates the new, highly-constrained prompt for the Azure OpenAI API.
    VERSION 6.1: Confirms the reversion to a strict ONE-PARAGRAPH rule, while
    explicitly preserving the content quality and detail from previous versions.

    Args:
        person_name (str): The first name of the person being evaluated.
        pronoun (str): The correct pronoun for the person (e.g., 'He' or 'She').
        person_data (str): A string representation of the person's scores and competencies.

    Returns:
        str: A fully constructed prompt ready to be sent to the AI model.
    """
    prompt_text = f"""
You are an elite talent management consultant from a top-tier firm, known for writing insightful, nuanced, and actionable executive summaries in flawless British English. Your writing is strategic and indistinguishable from a human expert.

## Core Objective
Synthesize the provided competency data for {person_name} into a **single, detailed, and unified narrative paragraph.** The content should be as rich and detailed as our best two-paragraph examples, but presented in a single block of text.

## Input Data for {person_name}
<InputData>
{person_data}
</InputData>

## ---------------------------------------------
## CRITICAL DIRECTIVES FOR SUMMARY STRUCTURE & TONE
## ---------------------------------------------

1.  **CRITICAL STRUCTURE: Single Paragraph ONLY.**
    * The entire summary **MUST** be a single, unbroken paragraph. Do not use line breaks.
    * The narrative must flow seamlessly from strengths into development areas.

2.  **Narrative Flow for Mixed-Score Profiles:**
    * **Opening:** Start with a holistic, professional, and narrative opening that introduces the candidate's key strengths.
    * **Strengths Section:** Describe the candidate's strengths first, providing rich behavioral detail.
    * **Seamless Transition:** After describing the strengths, you must move smoothly into the development areas. Use transition phrases like "To further enhance..." or "As a next step, {pronoun} could focus on..." to connect the parts of the narrative *without* starting a new paragraph.
    * **Development Section:** Detail the growth opportunities with the same level of specific, actionable suggestions as before.
    * **Concluding Sentence:** The final sentence of the paragraph should be a brief, positive, and forward-looking statement about the candidate's potential.

3.  **General Style Rules:**
    * Use clear, simple sentences. Aggressively avoid long sentences connected by many commas or conjunctions.
    * Describe behaviors with nuance; do not just label competencies.
    * ABSOLUTELY NO mention of scores, levels, or the assessment itself.
    * Use the candidate's first name once at the beginning, then the correct pronoun ({pronoun}) thereafter.

## ---------------------------------------------
## ANALYSIS OF A GOLD-STANDARD EXAMPLE (INTERNALIZE THIS LOGIC)
## ---------------------------------------------

**This example demonstrates the required SINGLE-PARAGRAPH structure for a mixed profile.**

* **Correct Output Example:** "Mahra demonstrated strengths especially in strategic thinking and talent development, with growing capability in driving innovation and making informed decisions. Her strategic thinking was evident in aligning initiatives with organisational goals and applying benchmarks, while she fostered learning by mentoring peers and sharing her expertise. To strengthen her strategic contribution, Mahra is encouraged to clarify any cost constraints and define more detailed budget plans. Her decision-making could be strengthened by improving risk forecasting and systematically evaluating long-term options, and she could drive results more effectively by setting clearer goals and monitoring progress more consistently. By embedding structured coaching and feedback to support long-term talent growth, Mahra has the potential to significantly elevate her leadership impact."

* **Analysis of the Logic:**
    * **Single Paragraph:** The entire text is one unified paragraph.
    * **Holistic Opening:** It starts by summarizing her key strengths in a narrative way.
    * **Seamless Transition:** It moves from strengths ("...mentoring peers and sharing her expertise.") directly into development needs ("To strengthen her strategic contribution..."). The transition is smooth and occurs mid-paragraph.
    * **Specific Feedback:** The development advice is actionable ("...clarify any cost constraints," "improving risk forecasting"). The level of detail is high.
    * **Integrated Closing:** The final sentence ("By embedding structured coaching...") is a forward-looking statement and is the natural conclusion of the single paragraph.

## ---------------------------------------------
## FINAL INSTRUCTIONS
## ---------------------------------------------

Now, process the provided competency data for {person_name}. Create a **strict single-paragraph summary** following all the rules above. The total word count should be **between 250 and 280 words** to ensure comprehensive detail is not lost. The final output should have the content and detail of our best prior examples, but formatted as one paragraph.
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
st.set_page_config(page_title="DGE Executive Summary Generator v6.1", layout="wide")

st.title("📄 DGE Executive Summary Generator (V6.1)")
st.markdown("""
This application generates professional executive summaries based on leadership competency scores. **Version 6.1 uses a strict one-paragraph structure.**
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
    label="📥 Download Sample Template File (V6.1)",
    data=sample_excel_data,
    file_name="dge_summary_template_v6.1.xlsx",
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
                st.subheader("Generated Summaries (V6.1)")
                
                output_df = df.copy()
                output_df['Executive Summary'] = generated_summaries
                
                st.dataframe(output_df)
                
                results_excel_data = to_excel(output_df)
                st.download_button(
                    label="📥 Download V6.1 Results as Excel",
                    data=results_excel_data,
                    file_name="Generated_Executive_Summaries_V6.1.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

    except Exception as e:
        st.error(f"An error occurred: {e}")
