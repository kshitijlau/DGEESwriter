Understood. A new critical rule will be implemented to strictly govern the format of the first sentence, ensuring it's one of four approved options and based on the single highest-scoring competency.

I will replace the existing 'Opening Sentence Protocol' in the prompt with this new, mandatory directive. The prompt will instruct the AI to programmatically identify the top score and use it to construct the opening sentence with no deviation from the four formats. The rest of the "Integrated Feedback Loop" logic will remain intact to elaborate on all the competencies after the first sentence.

Here is the complete and updated `app.py` file (`v7.1`) with this critical rule added.

```python
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

# --- The RE-ENGINEERED Master Prompt Template (Version 7.1 - Strict Opening Sentence) ---
def create_master_prompt(salutation_name, pronoun, person_data):
    """
    Dynamically creates the new, highly-constrained prompt for the Azure OpenAI API.
    VERSION 7.1: Implements a new, critical rule that forces the opening sentence
    into one of four exact formats, based on the single highest competency score.

    Args:
        salutation_name (str): The name to be used for the person, including titles (e.g., "Dr. Jonas", "Irene").
        pronoun (str): The correct pronoun for the person (e.g., 'He' or 'She').
        person_data (str): A string representation of the person's scores and competencies.

    Returns:
        str: A fully constructed prompt ready to be sent to the AI model.
    """
    prompt_text = f"""
You are an elite talent management consultant from a top-tier firm. Your writing is strategic, cohesive, and you follow instructions with precision. You will now follow a new, highly specific rule for the opening sentence.

## Core Objective
Synthesize the provided competency data for {salutation_name} into a single, cohesive, and integrated narrative paragraph.

## Input Data for {salutation_name}
<InputData>
{person_data}
</InputData>

## ---------------------------------------------
## CRITICAL DIRECTIVES FOR SUMMARY STRUCTURE & TONE
## ---------------------------------------------

1.  **CRITICAL OPENING SENTENCE PROTOCOL (MANDATORY):**
    * The very first sentence **MUST** be constructed based on the competency with the **single highest numerical score** from the input data.
    * This sentence **MUST** follow one of these four exact formats, with absolutely no deviation, additions, or elaboration.
        1.  `{salutation_name} evidenced a strong ability to [highest scoring competency verb phrase].`
        2.  `{salutation_name} evidenced a strong capacity to [highest scoring competency verb phrase].`
        3.  `{salutation_name} demonstrated a strong ability to [highest scoring competency verb phrase].`
        4.  `{salutation_name} demonstrated a strong capacity to [highest scoring competency verb phrase].`
    * **Example:** If the input data shows 'Results Driver: 3.55' is the highest score, the opening sentence must be one of the four options, such as: "Khasiba demonstrated a strong ability to drive results."
    * All other information and competency descriptions must begin from the **second sentence onward**.

2.  **STRUCTURE AFTER OPENING: The Integrated Feedback Loop.**
    * After the mandatory opening sentence, the rest of the paragraph **MUST** address each competency one by one in a logical flow.
    * For each competency, you will first describe the **observed positive behavior** (the strength).
    * Then, **IMMEDIATELY AFTER** describing the behavior, you will provide the **related development area** for that same competency, introduced with a phrase like "As a next step...", "To build on this...", or "To further develop...".
    * This structure of `[Strength 1] -> [Development for 1] -> [Strength 2] -> [Development for 2]` is mandatory for the body of the paragraph.

3.  **Name and Pronoun Usage:**
    * Use the candidate's full salutation name, **{salutation_name}**, only in the first sentence. Thereafter, use the pronoun **{pronoun}**.

4.  **Linguistic Variety:**
    * You MUST vary your descriptive language. Avoid reusing the same phrases for the same competency across different summaries.

5.  **General Style Rules:**
    * The entire summary must be a **single, unified paragraph**.
    * Use clear, simple sentences.
    * Provide **highly specific and actionable** development advice.
    * ABSOLUTELY NO mention of scores or levels.

## ---------------------------------------------
## ANALYSIS OF A GOLD-STANDARD EXAMPLE (INTERNALIZE THE NEW LOGIC)
## ---------------------------------------------

**This example demonstrates the required structure: A strict opening sentence, followed by the Integrated Feedback Loop, all within a SINGLE PARAGRAPH.**

* **Correct Output Example:** "Khasiba demonstrated a strong ability to drive results. She consistently evidenced the ability to analyse complex scenarios, align initiatives with organisational objectives, and anticipate broader implications, showcasing a thoughtful and forward-looking approach. Her decision-making was impactful, marked by a balance of pragmatism and insight. To build on this, she could focus on refining her ability to evaluate complex, high-stakes scenarios where clarity is limited. Khasiba showcased the ability to foster productive relationships across teams. As a next step, she could focus on enhancing her influence in group settings. Her customer-centric mindset was evident in her advocacy for solutions that prioritised client needs. To further develop this area, she could focus on identifying emerging customer expectations and embedding these insights into design processes."

* **Analysis of the Integrated Logic:**
    * **Strict Opening:** The summary begins with a sentence that follows the mandatory format, based on her highest score ('Results Driver').
    * **Cohesion:** After the opening, the summary flows logically through the other competencies.
    * **Integrated Feedback:** The development point for a competency comes *immediately* after its description. For example, `...foster productive relationships across teams.` (Strength) is immediately followed by `As a next step, she could focus on enhancing her influence...` (Related Development).

## ---------------------------------------------
## FINAL INSTRUCTIONS
## ---------------------------------------------

Now, process the data for {salutation_name}. Create a **strict single-paragraph summary**. The first sentence MUST follow the mandatory protocol based on the single highest score. The rest of the paragraph must follow the **Integrated Feedback Loop** structure. The total word count should remain between 250-280 words.
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
                {"role": "system", "content": "You are an elite talent management consultant. Your writing is strategic and cohesive. You follow all instructions with precision, especially the format of the opening sentence."},
                {"role": "user", "content": prompt}
            ],
            temperature=0.3,
            max_tokens=800,
            top_p=1.0,
            frequency_penalty=0.5,
            presence_penalty=0.2
        )
        return response.choices[0].message.content.strip()
    except Exception as e:
        st.error(f"An error occurred while contacting Azure OpenAI: {e}")
        return None

# --- Streamlit App Main UI ---
st.set_page_config(page_title="DGE Executive Summary Generator v7.1", layout="wide")

st.title("📄 DGE Executive Summary Generator (V7.1)")
st.markdown("""
This application generates professional executive summaries based on leadership competency scores.
**Version 7.1 uses a mandatory, strict format for the opening sentence.**
1.  **Set up your secrets**.
2.  **Download the Sample Template**. The format requires a `salutation_name` column.
3.  **Upload your completed Excel file**.
4.  **Click 'Generate Summaries'**.
""")

# --- Create and provide a sample file for download ---
sample_data = {
    'email': ['irene.a@example.com', 'jonas.k@example.com', 'khasiba.m@example.com'],
    'salutation_name': ['Irene', 'Dr. Jonas', 'Khasiba'],
    'gender': ['F', 'M', 'F'],
    'level': ['Director', 'Manager', 'Specialist'],
    'Strategic Thinker': [3.66, 3.23, 3.56],
    'Impactful Decision Maker': [3.51, 3.52, 3.11],
    'Effective Collaborator': [3.53, 3.28, 3.08],
    'Talent Nurturer': [3.38, 2.9, 2.93],
    'Results Driver': [3.3, 3.06, 3.55],
    'Customer Advocate': [3.29, 3.2, 3.34],
    'Transformation Enabler': [2.97, 3.02, 3],
    'Innovation Explorer': [3.42, 3.29, 3.24]
}
sample_df = pd.DataFrame(sample_data)
sample_excel_data = to_excel(sample_df)

st.download_button(
    label="📥 Download Sample Template File (V7.1)",
    data=sample_excel_data,
    file_name="dge_summary_template_v7.1.xlsx",
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
            
            identifier_cols = ['email', 'salutation_name', 'gender', 'level']
            all_known_competencies = [
                'Strategic Thinker', 'Impactful Decision Maker', 'Effective Collaborator',
                'Talent Nurturer', 'Results Driver', 'Customer Advocate',
                'Transformation Enabler', 'Innovation Explorer'
            ]
            competency_columns = [col for col in df.columns if col in all_known_competencies]
            
            if 'salutation_name' not in df.columns:
                st.error("Error: The uploaded file is missing the required 'salutation_name' column. Please download the new template and try again.")
                st.stop()

            generated_summaries = []
            progress_bar = st.progress(0)
            
            for i, row in df.iterrows():
                salutation_name = row['salutation_name']
                gender_input = str(row['gender']).upper()
                pronoun = 'They'
                if gender_input == 'M':
                    pronoun = 'He'
                elif gender_input == 'F':
                    pronoun = 'She'
                else:
                    st.warning(f"Invalid or missing gender '{row['gender']}' for {salutation_name}. Defaulting to pronoun 'They'.")

                st.write(f"Processing summary for: {salutation_name}...")
                
                scores_data = []
                for competency in competency_columns:
                    if competency in row and pd.notna(row[competency]):
                        scores_data.append(f"- {competency}: {float(row[competency])}")
                person_data_str = "\n".join(scores_data)

                prompt = create_master_prompt(salutation_name, pronoun, person_data_str)
                summary = generate_summary_azure(prompt, azure_api_key, azure_endpoint, azure_deployment_name)
                
                if summary:
                    generated_summaries.append(summary)
                    st.success(f"Successfully generated summary for {salutation_name}.")
                else:
                    generated_summaries.append("Error: Failed to generate summary.")
                    st.error(f"Failed to generate summary for {salutation_name}.")

                progress_bar.progress((i + 1) / len(df))

            if generated_summaries:
                st.balloons()
                st.subheader("Generated Summaries (V7.1)")
                
                output_df = df.copy()
                output_df['Executive Summary'] = generated_summaries
                
                st.dataframe(output_df)
                
                results_excel_data = to_excel(output_df)
                st.download_button(
                    label="📥 Download V7.1 Results as Excel",
                    data=results_excel_data,
                    file_name="Generated_Executive_Summaries_V7.1.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

    except Exception as e:
        st.error(f"An error occurred: {e}")
```
