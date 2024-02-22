import streamlit as st
import openai
import pandas as pd
from io import BytesIO

# Initialize the OpenAI client securely
openai.api_key = st.secrets["api_key"]

# Define the Streamlit app layout
st.title('VC Outreach Email Personalization Tool')

# Input fields for customizing the AI prompt
custom_role = st.text_input("Custom Role Description", value="VC analyst")
custom_instruction = st.text_input("What would you like the email to be? *don't remove COMPANY_NAME and COMPANY_DESCRIPTION from intstruction", value="In 30 words, describe the synergy between COMPANY_NAME's work on COMPANY_DESCRIPTION and our VC's focus on VC_DESCRIPTION, excluding subject, greeting, or signature.")

# Initialize session state for storing the DataFrame if it doesn't exist
if 'personalized_df' not in st.session_state:
    st.session_state.personalized_df = pd.DataFrame(columns=['Company Name', 'Personalized Section'])

# File upload for the company descriptions
company_file = st.file_uploader("Upload Company Descriptions File (CSV or Excel)", type=['csv', 'xlsx'])

# File upload for the VC description
vc_description_file = st.file_uploader("Upload Company Description File (Text)", type=['txt'])

# Function to read VC description file
def read_vc_description(vc_file):
    if vc_file is not None:
        vc_description = vc_file.getvalue().decode("utf-8")
        return vc_description
    return ""

vc_description = read_vc_description(vc_description_file) if vc_description_file else ""

# Function to use the OpenAI ChatCompletion API for generating personalized sections
def generate_personalized_section(company_name, company_description, vc_description, custom_role, custom_instruction):
    # Replace placeholders in custom_instruction with actual values
    instruction = custom_instruction.replace("COMPANY_NAME", company_name).replace("COMPANY_DESCRIPTION", company_description).replace("VC_DESCRIPTION", vc_description)
    
    conversation = [
        {"role": "system", "content": f"You are a {custom_role} at {{vc_description}}"},
        {"role": "user", "content": instruction}
    ]
    
    response = openai.ChatCompletion.create(
        model="gpt-4-0125-preview",
        messages=conversation
    )
    
    return response.choices[0].message['content']

# Function to check if the description exceeds word limit
def description_exceeds_limit(description, limit=250):
    return len(description.split()) > limit

# Function to process uploaded files and generate personalized emails
def process_and_generate_emails(company_df, vc_desc, custom_role, custom_instruction):
    if company_df.shape[0] > 20:
        st.error("The CSV file contains more than 20 rows. Please limit your input to 20 companies or contact jadrima1@gmail.com for assistance.")
        return
    
    data_list = []

    if not company_df.empty and vc_desc:
        for _, row in company_df.iterrows():
            company_name = row['Company Name']
            description = row['Description']
            
            if description_exceeds_limit(description):
                st.error(f"The description for {company_name} exceeds 250 words. Please reduce the length.")
                continue
            
            personalized_section = generate_personalized_section(company_name, description, vc_desc, custom_role, custom_instruction)
            data_list.append({'Company Name': company_name, 'Personalized Section': personalized_section})

    # Convert list to DataFrame and update session_state
    st.session_state.personalized_df = pd.concat([st.session_state.personalized_df, pd.DataFrame(data_list)], ignore_index=True)

# Main logic to generate and display emails, and prepare Excel download
if company_file is not None:
    df = pd.read_csv(company_file) if company_file.type == "text/csv" else pd.read_excel(company_file)
    process_and_generate_emails(df, vc_description, custom_role, custom_instruction)
    st.write(st.session_state.personalized_df)

# Function to convert DataFrame to Excel for download
def convert_df_to_excel():
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        st.session_state.personalized_df.to_excel(writer, index=False)
    output.seek(0)
    return output

# Download button for Excel file
if not st.session_state.personalized_df.empty:
    excel_file = convert_df_to_excel()
    st.download_button(
        label="Download Personalized Emails Excel",
        data=excel_file,
        file_name="personalized_emails.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
