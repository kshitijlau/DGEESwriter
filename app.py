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

# --- The RE-ENGINEERED Master Prompt Template (Version 2) ---
# This prompt is unchanged as the logic change happens before the prompt is called.
def create_master_prompt(person_name, pronoun, person_data):
    """
    Dynamically creates the new, highly-constrained prompt for the Azure OpenAI API.
    This incorporates specific rules about tone, structure, phrasing, and forbidden content.

    Args:
        person_name (str): The first name of the person being evaluated.
        pronoun (str): The correct pronoun for the person (e.g., 'He' or 'She').
        person_data (str): A string representation of the person's scores and competencies.

    Returns:
        str: A fully constructed prompt ready to be sent to the AI model.
    """
    return f"""
You are a senior talent assessment consultant and expert report writer, specializing in executive summaries for leadership assessments. Your writing style is indistinguishable from a human expert, characterized by concise, clear, and professional British English.

## Core Objective
Synthesize the provided competency data for {person_name} into a single, fluid narrative paragraph.

## Input Data for {person_name}
<InputData>
{person_data}
</InputData>

## ---------------------------------------------
## ABSOLUTE DIRECTIVES & RULES OF WRITING
## ---------------------------------------------

1.  **Structure: Single Paragraph ONLY.**
    * The entire summary MUST be a single, unbroken paragraph. Do not use multiple paragraphs or line breaks.

2.  **CRITICAL RULE - Writing Style: Describe Behavior, Don't Label Competency.**
    * You must translate the competency name into a behavioral description.
    * **Incorrect (Labeling):** "{person_name} demonstrated Strategic Thinker."
    * **Correct (Describing):** "{person_name} evidenced a strong ability to think strategically." or "{pronoun} demonstrated the capacity to drive for results."

3.  **CRITICAL RULE - Narrative Voice: Be Concise and Human-like.**
    * AVOID long, complex sentences connected by 'while', 'and', or 'however'.
    * Prefer two shorter, direct sentences over one long, rambling one. This is essential.
    * DO NOT add generic, "fluff" phrases to the end of sentences. Be direct and impactful.
    * **Incorrect (AI-like, long):** "His ability to nurture talent was reflected in his supportive interactions, where he encouraged growth and development among team members, fostering a positive and inclusive environment."
    * **Correct (Human-like, concise):** "His ability to nurture talent was reflected in his supportive interactions and encouragement of team member growth."

4.  **Opening Sentence Protocol:**
    * The opening sentence MUST be positive and framed around the candidate's highest-scoring competency.
    * Use one of these specific formats: "{person_name} evidenced a strong ability to [describe highest competency]." OR "{person_name} demonstrated the capacity to [describe highest competency]."
    * NEVER open with a balanced statement about strengths and development areas.

5.  **Closing Sentence Protocol:**
    * DO NOT use a closing or summary sentence. The narrative should end naturally after describing the last behavior or development area.
    * **AVOID sentences like:** "Overall, [Name]'s profile is strong..." or "By addressing these areas, [Name] can improve..."

6.  **Pronoun Usage:**
    * Use the candidate's first name, "{person_name}", only once at the very beginning of the summary.
    * Thereafter, use ONLY the provided pronoun: **{pronoun}**.

## ---------------------------------------------
## FORBIDDEN CONTENT - YOU MUST NOT INCLUDE THESE
## ---------------------------------------------

1.  **NO MENTION OF THE ASSESSMENT:**
    * Do not use phrases like: "several leadership competencies assessed", "as observed during the assessment", "In his interactions", "during the activity". The summary should read as a general character assessment.

2.  **NO MENTION OF SCORES OR LEVELS (Directly or Indirectly):**
    * ABSOLUTELY NO numerical scores.
    * Do not use "proxy" words that imply a score, such as: "showed promise in", "demonstrated a foundational ability to", "is developing in". Describe the behavior neutrally or as a development area.

3.  **NO "COGNITIVE" ABILITIES:**
    * Do not mention "cognitive abilities", "cognitive agility", "cognitive tools", or similar phrases. Focus only on the behavioral competencies provided.

4.  **AVOID OVERUSED WORDS:**
    * Do not use the words "foster" or "underscored". Find more varied and professional synonyms.

## ---------------------------------------------
## REVISED EXAMPLES (INTERNALIZE THIS NEW LOGIC)
## ---------------------------------------------

**Example 1: High-Scorer Profile**
* **Logic:** Single, positive narrative. Opens with the highest strength. No closing summary. Sentences are concise.
* **Correct Output:** "Amal demonstrated the capacity to think strategically, consistently anticipating future challenges and aligning actions with long-term organisational goals. She showcased a strong ability to drive for results, setting ambitious goals and maintaining a disciplined focus on achieving them. Her decision-making reflected a balanced consideration of risks and opportunities. Furthermore, she effectively collaborated with peers to achieve shared objectives."

**Example 2: Mixed-Score Profile**
* **Logic:** Integrated narrative. Opens with a clear strength. Transitions smoothly to actionable development needs. Uses concise language. No closing summary.
* **Correct Output:** "Rashed evidenced a strong ability to drive for results by focusing the team on clear objectives. He effectively collaborated with his peers and demonstrated a capacity for making impactful decisions. To further enhance his leadership, Rashed may benefit from developing a more customer-centric approach by actively seeking client feedback to better inform service delivery. He could also strengthen his ability to think strategically by analysing industry trends to identify growth opportunities, rather than focusing solely on operational tactics."

## Final Instruction for {person_name}
Now, process the provided competency data for {person_name}. Adhering to every rule above, write a single, concise, human-like paragraph of no more than 280 words. Begin with the specified opening sentence format and do not include a closing summary sentence.
"""

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
                {"role": "system", "content": "You are a senior talent assessment consultant and expert report writer following British English standards. You write concisely and with a human touch, avoiding jargon."},
                {"role": "user", "content": prompt}
            ],
            temperature=0.2,
            max_tokens=600,
            top_p=1.0,
            frequency_penalty=0.2,
            presence_penalty=0.0
        )
        return response.choices[0].message.content.strip()
    except Exception as e:
        st.error(f"An error occurred while contacting Azure OpenAI: {e}")
        return None

# --- Streamlit App Main UI ---
st.set_page_config(page_title="DGE Executive Summary Generator v3", layout="wide")

st.title("📄 DGE Executive Summary Generator (V3)")
# CHANGE: Updated instructions for the user.
st.markdown("""
This application generates professional executive summaries based on leadership competency scores.
1.  **Set up your secrets**. Enter your Azure OpenAI credentials in the Streamlit Cloud app settings.
2.  **Download the Sample Template**. The format now requires a `gender` column (enter 'M' for Male, 'F' for Female).
3.  **Upload your completed Excel file**.
4.  **Click 'Generate Summaries'** to process the file.
""")

# --- Create and provide a sample file for download ---
# CHANGE: The 'pronoun' column is now 'gender', and values are 'F' and 'M'.
sample_data = {
    'email': ['jane.doe@example.com', 'john.roe@example.com'],
    'first_name': ['Jane', 'John'],
    'last_name': ['Doe', 'Roe'],
    'gender': ['F', 'M'], # NEW required column with 'M' or 'F'
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

# CHANGE: Updated label and file name for the new template.
st.download_button(
    label="📥 Download Sample Template File (with Gender column)",
    data=sample_excel_data,
    file_name="dge_summary_template_v3.xlsx",
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
            
            # CHANGE: Identifier column is now 'gender'.
            identifier_cols = ['email', 'first_name', 'last_name', 'level', 'gender']
            all_known_competencies = [
                'Strategic Thinker', 'Impactful Decision Maker', 'Effective Collaborator',
                'Talent Nurturer', 'Results Driver', 'Customer Advocate',
                'Transformation Enabler', 'Innovation Catalyst'
            ]
            competency_columns = [col for col in df.columns if col in all_known_competencies]
            
            # CHANGE: Check for the new 'gender' column.
            if 'gender' not in df.columns:
                st.error("Error: The uploaded file is missing the required 'gender' column. Please download the new template and try again.")
                st.stop()

            generated_summaries = []
            progress_bar = st.progress(0)
            
            for i, row in df.iterrows():
                first_name = row['first_name']
                # CHANGE: New logic to map gender ('M'/'F') to a pronoun ('He'/'She').
                gender_input = str(row['gender']).upper()
                pronoun = 'They' # Default pronoun
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

                # Pass the correctly determined pronoun to the prompt function.
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
                st.subheader("Generated Summaries (V3)")
                
                output_df = df.copy()
                output_df['Executive Summary'] = generated_summaries
                
                st.dataframe(output_df)
                
                results_excel_data = to_excel(output_df)
                st.download_button(
                    label="📥 Download V3 Results as Excel",
                    data=results_excel_data,
                    file_name="Generated_Executive_Summaries_V3.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

    except Exception as e:
        st.error(f"An error occurred: {e}")
