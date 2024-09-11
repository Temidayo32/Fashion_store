from openai import OpenAI
import gspread
from gspread.exceptions import WorksheetNotFound
from oauth2client.service_account import ServiceAccountCredentials
from initpandas import emails_df
from configureOpenai import client

# Open the existing spreadsheet by name
spreadsheet_name = 'Copy of Solving Business Problems with AI'
document_id = '1F6tf8E1R5Sga2MgJlpUAH_h7GK-UI2fVyju9aPu3FPY'
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name('/content/useful-mile-401221-4c4cae9e8d0a.json', scope)
gclient = gspread.authorize(creds)

# Open the existing spreadsheet by name
spreadsheet_name = 'Copy of Solving Business Problems with AI'
spreadsheet = gclient.open(spreadsheet_name)
# Define the function to classify emails
def classify_email(text):
    prompt = f"Classify the following email as either 'Product Inquiry' or 'Order Request':\n\n{text}\n\nRespond with only 'Product Inquiry' or 'Order Request'."
    try:
        completion = client.chat.completions.create(
          model="gpt-4o",
          messages=[
            {"role": "user", "content": prompt}
          ]
        )
        # Extract the classification result from the response
        classification = completion.choices[0].message.content.strip().lower()
        return classification
    except Exception as e:
        print(f"An error occurred: {e}")
        return "error"

# Function to classify
def classify(row):
    text = f"Subject: {row['subject']}\nBody: {row['message']}"

    # Classify the content
    category = classify_email(text)
    return category

# Apply the classify function to each row
emails_df['category'] = emails_df.apply(classify, axis=1)

# Create a new worksheet named 'email-classification'
def get_or_create_worksheet(spreadsheet, title, rows='1000', cols='20'):
    try:
        # Check if the worksheet already exists
        worksheet = spreadsheet.worksheet(title)
        print(f"Worksheet '{title}' already exists.")
    except WorksheetNotFound:
        # If not, create a new one
        worksheet = spreadsheet.add_worksheet(title=title, rows=rows, cols=cols)
        print(f"New worksheet '{title}' created.")

    return worksheet

# emails_df already populated with 'email_id' and 'category' columns
classification_df = emails_df[['email_id', 'category']]

# Convert DataFrame to list of lists for uploading
data_to_upload = classification_df.values.tolist()
header = ['email_id', 'category']
data_to_upload.insert(0, header)

# Update the new worksheet with the data
worksheet = get_or_create_worksheet(spreadsheet, 'email-classificatio')
worksheet.update('A1', data_to_upload)

print("Email Classification data uploaded successfully!")

