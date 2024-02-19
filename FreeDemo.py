import streamlit as st
import openai
import pandas as pd
from io import BytesIO

# Initialize the OpenAI client securely
openai.api_key = st.secrets["api_key"]

# Define the Streamlit app layout
st.title('VC Outreach Email Personalization Tool')

# Initialize session state for storing the DataFrame if it doesn't exist
if 'personalized_df' not in st.session_state:
    st.session_state.personalized_df = pd.DataFrame(columns=['Company Name', 'Personalized Section'])

# File upload for the company descriptions
company_file = st.file_uploader("Upload Company Descriptions File (CSV or Excel)", type=['csv', 'xlsx'])

# File upload for the VC description
vc_description_file = st.file_uploader("Upload VC Description File (Text)", type=['txt'])

# Function to read VC description file
def read_vc_description(vc_file):
    if vc_file is not None:
        vc_description = vc_file.getvalue().decode("utf-8")
        return vc_description
    return ""

vc_description = read_vc_description(vc_description_file) if vc_description_file else ""

# Function to use the OpenAI ChatCompletion API for generating personalized sections
def generate_personalized_section(company_name, company_description, vc_description):
    conversation = [
        {"role": "system", "content": "You are a VC analyst at {vc_description}"},
        {"role": "user", "content": f"In 30 words, describe the synergy between {company_name}'s work on {company_description} and our VC's focus on {vc_description}, excluding subject, greeting, or signature."}
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
def process_and_generate_emails(company_df, vc_desc):
    if company_df.shape[0] > 20:
        st.error("The CSV file contains more than 20 rows. Please limit your input to 20 companies or contact jadrima1@gmail.com for assistance.")
        return
    
    data_list = []  # Collect data in a list to avoid inefficient DataFrame.append()

    if not company_df.empty and vc_desc:
        for _, row in company_df.iterrows():
            company_name = row['Company Name']
            description = row['Description']
            
            if description_exceeds_limit(description):
                st.error(f"The description for {company_name} exceeds 250 words. Please reduce the length or contact jadrima1@gmail.com for assistance.")
                continue  # Skip processing this row
            
            personalized_section = generate_personalized_section(company_name, description, vc_desc)
            data_list.append({'Company Name': company_name, 'Personalized Section': personalized_section})

    # Convert list to DataFrame and update session_state
    st.session_state.personalized_df = pd.concat([st.session_state.personalized_df, pd.DataFrame(data_list)], ignore_index=True)

# Main logic to generate and display emails, and prepare Excel download
if company_file is not None:
    df = pd.read_csv(company_file) if company_file.type == "text/csv" else pd.read_excel(company_file)
    if df.shape[0] > 20:
        st.error("The CSV file contains more than 20 rows. Please limit your input to 20 companies or contact jadrima1@gmail.com for assistance.")
    else:
        process_and_generate_emails(df.iloc[:20], vc_description)  # Limit processing to the first 20 rows
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
