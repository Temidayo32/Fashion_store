import pandas as pd
import openai
from fuzzywuzzy import fuzz, process
from configureOpenai import client
from gspread.exceptions import WorksheetNotFound
from initpandas import emails_df, products_df
from classifyEmails import spreadsheet

def find_best_match_product(inquiry, products_df):
    # Initialize variables to store the best match
    best_match = None
    best_score = 80
    product_id, product_name, product_price = None, None, None

    for index, row in products_df.iterrows():
        # Prepare product name and ID for matching
        name = str(row['name'])
        id = str(row['product_id'])
        price = row['price']  # Assume there's a price column

        # Check for exact match with product_id
        if id in inquiry:
            return id, name, price

        # Fuzzy matching for product name
        score = fuzz.partial_ratio(inquiry.lower(), name.lower())
        if score > best_score:
            best_score = score
            best_match = name
            product_id = id
            product_price = price

    if best_match:
        return product_id, best_match, product_price

    return None, None, None


# Define your function to generate AI-powered responses
def generate_inquiry_response(inquiry, products_df):
    # Find the best match product using fuzzy matching
    product_id, product_name, product_price = find_best_match_product(inquiry, products_df)

    if product_id or product_name:
        # a prompt for the AI model with the matched product
        prompt = (f"A customer has inquired about a product with the following information: "
                  f"Product Name: {product_name}, Product Price: ${product_price}. "
                  f"Please generate a friendly and informative response thanking them and providing details about the product."
                  f"Use a generic greeting and use Lucious Stores as company name in the closing text")
    else:
        # a prompt for the AI model when no product match is found
        prompt = ("A customer has inquired about a product, but no matching product was found in the catalog. "
                  "Please generate a response asking for more details and offering assistance in finding a suitable product. "
                  "Encourage them to provide more specifics or check our catalog for alternative options."
                  "Use a generic greeting and use Lucious Stores as company name in the closing text")

    # Example using OpenAI's GPT model to generate the response

    response = client.chat.completions.create(
      model="gpt-4o",
      messages=[
        {"role": "user", "content": prompt}
      ]
    )

    return response.choices[0].message.content.strip()

# Define your function to read data from Google Sheets
def read_data_frame(document_id, sheet_name):
    export_link = f"https://docs.google.com/spreadsheets/d/{document_id}/gviz/tq?tqx=out:csv&sheet={sheet_name}"
    return pd.read_csv(export_link, header=0)

# Document ID for Google Sheets
document_id = '1F6tf8E1R5Sga2MgJlpUAH_h7GK-UI2fVyju9aPu3FPY'

# Read sheets
classification_df = read_data_frame(document_id, 'email-classification')

# Filter out rows with category 'product inquiry'
inquiry_df = classification_df[classification_df['category'].str.contains('product inquiry', case=False)]
# Prepare a list to collect responses
inquiry_responses = []

# Process each product inquiry
for index, row in inquiry_df.iterrows():
    email_id = row['email_id']

    # Fetch email details
    email_row = emails_df[emails_df['email_id'] == email_id].iloc[0]
    subject = str(email_row['subject']) if email_row['subject'] is not None else ''
    body = str(email_row['message']) if email_row['message'] is not None else ''

    # Merge subject and body to form the full inquiry
    inquiry = subject + ' ' + body

    # Generate AI-powered response
    response = generate_inquiry_response(inquiry, products_df)

    # Append the response to the list
    inquiry_responses.append({
        'email_id': email_id,
        'response': response
    })

# Convert the list to a DataFrame
inquiry_response_df = pd.DataFrame(inquiry_responses)

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

# Convert DataFrame to list of lists for uploading
data_to_upload = inquiry_response_df.values.tolist()
header = ['email_id', 'response']
data_to_upload.insert(0, header)

# Update the worksheet with the data
worksheet = get_or_create_worksheet(spreadsheet, 'inquiry-response')
worksheet.update('A1', data_to_upload)

print("Inquiry response data uploaded successfully!")
