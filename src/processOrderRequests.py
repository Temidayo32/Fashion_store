import pandas as pd
from openai import OpenAI
from fuzzywuzzy import fuzz, process
from gspread.exceptions import WorksheetNotFound
import re
from initpandas import emails_df, products_df
from classifyEmails import spreadsheet
from configureOpenai import client

def read_data_frame(document_id, sheet_name):
    export_link = f"https://docs.google.com/spreadsheets/d/{document_id}/gviz/tq?tqx=out:csv&sheet={sheet_name}"
    return pd.read_csv(export_link, header=0)

# Read the classification, emails, and products sheets
document_id = '1F6tf8E1R5Sga2MgJlpUAH_h7GK-UI2fVyju9aPu3FPY'
classification_df = read_data_frame(document_id, 'email-classification')


# Filter order requests
order_requests = classification_df[classification_df['category'].str.contains('order request', case=False)]

#Function to find product ID based on product name/product ID in email subject or body
def find_product_ids(subject, body, products_df):
    subject = str(subject) if subject is not None else ''
    body = str(body) if body is not None else ''
    combined_text = f"{subject} {body}".lower()

    # Initialize an empty list to store matches
    matches = []
    seen_product_ids = set()  # To track seen product IDs and avoid duplicates
    seen_product_names = set()  # To track seen product names and avoid duplicates

    # Check for product IDs directly
    product_ids = products_df['product_id'].tolist()

    for product_id in product_ids:
        # Use fuzzy matching to find product IDs in the combined text
        if fuzz.partial_ratio(product_id.lower(), combined_text) > 90:  # Adjust threshold as needed
            if product_id not in seen_product_ids:
                row = products_df[products_df['product_id'] == product_id].iloc[0]
                matches.append((row['product_id'], row['name']))
                seen_product_ids.add(product_id)

    # Fuzzy matching for product names
    product_names = products_df['name'].tolist()

    # Iterate over product names and find the best match
    best_matches = process.extractBests(combined_text, product_names, scorer=fuzz.partial_ratio)

    for best_match, score in best_matches:
        if score > 80:  # Adjust threshold as needed
            matched_row = products_df[products_df['name'].str.lower() == best_match.lower()].iloc[0]
            if matched_row['product_id'] not in seen_product_ids and best_match not in seen_product_names:
                matches.append((matched_row['product_id'], matched_row['name']))
                seen_product_ids.add(matched_row['product_id'])
                seen_product_names.add(best_match)
    return matches


def extract_quantity_from_text(text, product_name, stock_level):

    prompt = f"""
    Extract the quantity for the following product from the text. If the quantity is explicitly stated, use that number. If the quantity is not explicitly stated but can be inferred to mean all available stock (e.g., "all the remaining," "as many as possible," etc.), respond with the current stock level provided: {stock_level}. If there's no clear indication, assume a default quantity of '1'.

    Product: {product_name}

    Text:
    {text}

    Provide the quantity in the following format:
    - Product Name: Quantity (e.g., '5', '1', or '{stock_level}')
    """


    try:
        response = client.chat.completions.create(
            model="gpt-4o",
            messages=[
                {"role": "user", "content": prompt}
            ]
        )

        response_text = response.choices[0].message.content.strip()

        # parsing logic
        lines = response_text.split('\n')
        for line in lines:
            if product_name.lower() in line.lower():
                parts = line.split(':', 1)
                if len(parts) == 2:
                    quantity_text = parts[1].strip()
                    return int(quantity_text) if quantity_text.isdigit() else None

        return None

    except Exception as e:
        print(f"An error occurred: {e}")
        return None

def determine_purchase_intent(product_name_or_id, text):
    prompt = f"""
    Based on the following email text, determine if the customer intends to purchase or place an order for the product '{product_name_or_id}'. Focus specifically on phrases that indicate a desire to buy, order, or request the product. Ignore mentions of satisfaction with previous purchases, compliments, or future purchase plans.

    Email text: "{text}"

    If the customer clearly intends to purchase '{product_name_or_id}', respond with 'yes'. If the text does not indicate a clear intention to purchase, respond with 'no'.
    """


    try:
        response = client.chat.completions.create(
            model="gpt-4o",
            messages=[
                {"role": "user", "content": prompt}
            ]
        )

        result = response.choices[0].message.content.strip()

        return result.lower() == 'yes'

    except Exception as e:
        print(f"An error occurred while determining purchase intent: {e}")
        return False


# Process each order request
order_status_df = pd.DataFrame(columns=['email_id', 'product_id', 'quantity', 'status'])

# Create a list to collect new rows
order_status_rows = []

# Process each order request
for index, row in order_requests.iterrows():
    email_id = row['email_id']
    email_row = emails_df[emails_df['email_id'] == email_id].iloc[0]
    subject = str(email_row['subject']) if email_row['subject'] is not None else ''
    body = str(email_row['message']) if email_row['message'] is not None else ''

    # Match product name and get stock
    products = find_product_ids(subject, body, products_df)
    # print(products)

    if products:
      for product_id, product_name in products:
          # Use AI to determine if the customer intends to purchase this product
          if determine_purchase_intent(product_id, subject + ' ' + body):
              # Extract quantity for each product
              stock = products_df[products_df['product_id'] == product_id]['stock'].values[0]
              quantity_text = extract_quantity_from_text(subject + ' ' + body, product_name, stock)

              if quantity_text is None:
                  # Handle case where quantity could not be extracted
                  # Default to 1 as single orders often don't specify quantity.
                  quantity = 1
              else:
                  quantity = quantity_text

              # Ensure that stock and quantity are integers
              if isinstance(stock, (int, float)) and isinstance(quantity, (int, float)):
                  stock = int(stock)
                  print(stock)
                  quantity = int(quantity)

              if stock >= quantity:
                  status = 'created'
                  # Update stock in the products DataFrame
                  products_df.loc[products_df['product_id'] == product_id, 'stock'] -= quantity
              else:
                  status = 'out of stock'

              # Collect result for this order
              order_status_rows.append({
                  'email_id': email_id,
                  'product_id': product_id,
                  'quantity': quantity,
                  'status': status
              })
          else:
              print(f"Customer does not intend to purchase: {product_name}")
    else:
        # Handle case where product ID was not found
        order_status_rows.append({
            'email_id': email_id,
            'product_id': 'not found',
            'quantity': 'N/A',
            'status': 'out of stock'
        })



# Convert list of dictionaries to DataFrame
order_status_df = pd.DataFrame(order_status_rows)

# Create or open the 'order-status' worksheet
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
data_to_upload = order_status_df.values.tolist()
header = ['email_id', 'product_id', 'quantity', 'status']
data_to_upload.insert(0, header)

# Update the worksheet with the data
worksheet = get_or_create_worksheet(spreadsheet, 'order-status')
worksheet.update('A1', data_to_upload)

print("Order status data uploaded successfully!")
