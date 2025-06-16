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
        person_name (str): The name of the person being evaluated.
        person_data (str): A string representation of the person's scores and competencies.

    Returns:
        str: A fully constructed prompt ready to be sent to the AI model.
    """
    return f"""
You are a senior talent assessment consultant specializing in writing executive summaries for leadership assessment centers. Your writing style adheres strictly to British English and maintains a professional, constructive, and positive tone.

## Core Objective
Your task is to synthesize the provided competency scores for {person_name} into a formal, two-part executive summary. The summary must be a cohesive narrative, not a list of points.

## Input Data for {person_name}
{person_data}

## Primary Directives & Rules

1.  **Overall Structure & Word Count:**
    * The entire summary must not exceed 280 words.
    * The summary must address all 8 competencies provided in the input data.
    * The summary is divided into two main parts: Strengths and Growth Opportunities.

2.  **Opening Sentence:**
    * You MUST begin the summary with a single introductory sentence that highlights the 1-2 highest-scoring competencies.
    * Start with: "{person_name} presented as a..." or "{person_name} displayed strengths in...".
    * Example: "{person_name} presented as a strong collaborator and impactful decision-maker."

3.  **Strengths Section:**
    * This section follows the opening sentence.
    * Discuss all competencies with scores of 4 or 5.
    * If a competency has a score of 3, you may position it as a strength, but it should be mentioned after the clear strengths (scores 4 & 5).
    * Use **past tense** for all observed behaviors (e.g., "He demonstrated...", "She displayed...", "He showcased...").

4.  **Growth Opportunities Section:**
    * This section follows the strengths.
    * Discuss all competencies with scores of 1 or 2.
    * Frame these areas constructively (e.g., "To further her development...", "He could elevate his impact by...", "As a next step, she could focus on...").
    * Use **present or future tense** for all development suggestions.
    * **CRITICAL:** For each growth opportunity, you must provide a **specific, actionable, and measurable suggestion**. Do not just state the weakness; provide a tangible development action.

5.  **Language and Tone:**
    * Use **British English** only.
    * The name of the candidate must be spelled exactly as provided: "{person_name}".
    * Avoid using generic words like "good" or "bad." Use professional synonyms like "effective," "strong," "adept," etc.
    * Ensure all punctuation is used correctly.

## Deconstruction of Examples (Internalize this Logic)

**Example 1: Jane (High-Scoring Profile)**
* **Scores (Inferred):** High scores in Effective Collaborator, Impactful Decision Maker, Talent Nurturer, Results Driver, Strategic Thinker. Low scores in some areas requiring development suggestions.
* **Correct Logic:**
    1.  **Opening:** Start by highlighting the top two strengths: "Jane presented as a strong collaborator and impactful decision-maker..."
    2.  **Strengths Narrative:** Weave a story about how she demonstrated these strengths in the past tense (e.g., "Jane demonstrated strong data-driven decision-making...").
    3.  **Development:** Frame growth opportunities constructively with specific actions: "To further her development, Jane could enhance her collaboration skills by actively seeking feedback from colleagues, aiming for at least three diverse perspectives on key decisions each month."

**Example 2: Dr. Jonas (Mixed-Score Profile)**
* **Scores (Inferred):** High score in Impactful Decision Maker, with several other competencies requiring development.
* **Correct Logic:**
    1.  **Opening:** Start with the single highest strength: "Dr. Jonas presented as an impactful decision maker."
    2.  **Strengths Narrative:** Detail his other strengths in the past tense (e.g., "He evidenced making bold decisions...").
    3.  **Development:** Address low scores with actionable advice. For collaboration, suggest: "He may strengthen collaboration by using more thoughtful and tailored influencing and persuading techniques...". For change, suggest: "Gathering early reactions to change, raising awareness and communicating the vision clearly to others, could help smoother transitions."

**Example 3: Khasiba (Balanced Profile with 'Competent' Scores)**
* **Scores (Inferred):** Strengths in Strategic Thinking & Results Driver, with many other areas being 'competent' (score 3) requiring balanced feedback.
* **Correct Logic:**
    1.  **Opening:** Start with the top strengths: "Khasiba displayed strengths in strategic thinking and driving results."
    2.  **Balanced Feedback:** For 'competent' skills, present the demonstrated ability first (past tense), then immediately provide a future-focused development action. Example: "She showcased the ability to work effectively with peers and stakeholders... As a next step, she could focus on enhancing her influence in group settings..."
    3.  **Actionable Detail:** Development suggestions are specific: "To deepen this capability, she could take a more active role in shaping the development culture of her team by introducing coaching and supporting stretch assignments."

## Final Instruction for {person_name}
Now, process the provided competency scores for {person_name}. Write a concise, two-part executive summary under 280 words, adhering to all the rules above. Start with the required opening line, use the correct tenses, cover all 8 competencies, and provide specific, actionable development suggestions for every growth opportunity.
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
                # Construct full name and format scores for the prompt
                full_name = f"{row['first_name']} {row['last_name']}"
                st.write(f"Processing summary for: {full_name}...")
                
                scores_data = []
                for competency in competency_columns:
                    # Check if the competency column exists and the value is not null
                    if competency in row and pd.notna(row[competency]):
                        scores_data.append(f"- {competency}: {row[competency]}")
                person_data_str = "\n".join(scores_data)

                # Create and run the prompt for the current person
                prompt = create_master_prompt(full_name, person_data_str)
                summary = generate_summary_azure(prompt, azure_api_key, azure_endpoint, azure_deployment_name)
                
                if summary:
                    generated_summaries.append(summary)
                    st.success(f"Successfully generated summary for {full_name}.")
                else:
                    generated_summaries.append("Error: Failed to generate summary.")
                    st.error(f"Failed to generate summary for {full_name}.")

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
