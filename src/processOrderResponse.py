import pandas as pd
from initpandas import products_df
from classifyEmails import spreadsheet
from gspread.exceptions import WorksheetNotFound

# Function to read data from Google Sheets
def read_data_frame(document_id, sheet_name):
    export_link = f"https://docs.google.com/spreadsheets/d/{document_id}/gviz/tq?tqx=out:csv&sheet={sheet_name}"
    return pd.read_csv(export_link, header=0)

# Document ID of the Google Sheets
document_id = '1F6tf8E1R5Sga2MgJlpUAH_h7GK-UI2fVyju9aPu3FPY'

# Read the order-status, emails, and products sheets
order_status_df = read_data_frame(document_id, 'order-status')

# Function to get product name by product ID
def get_product_name(product_id, products_df):
    product = products_df[products_df['product_id'] == product_id]
    return product['name'].values[0] if not product.empty else None

# Function to generate email response
def generate_email_response(email_id, product_id, quantity, status, products_df):
    if product_id == 'not found':
        response = (f"Dear Customer,\n\n"
                    f"Unfortunately, your order could not be fulfilled due to insufficient stock. We apologize for the inconvenience.\n\n"
                    f"You may choose to wait for restock or select an alternative product.\n"
                    f"Please contact our support team for further assistance.\n\n"
                    f"Best regards,\n"
                    f"Lucious Stores")
    else:
        product_name = get_product_name(product_id, products_df)

        if status == 'created':
            response = (f"Dear Customer,\n\n"
                        f"Your order for {quantity} unit(s) of {product_name} (Product ID: {product_id}) "
                        f"has been successfully processed. Thank you for shopping with us!\n\n"
                        f"Best regards,\n"
                        f"Lucious Stores")
        elif status == 'out of stock':
            response = (f"Dear Customer,\n\n"
                        f"Unfortunately, your order for {quantity} unit(s) of {product_name} (Product ID: {product_id}) "
                        f"could not be fulfilled due to insufficient stock. We apologize for the inconvenience.\n\n"
                        f"You may choose to wait for restock or select an alternative product.\n"
                        f"Please contact our support team for further assistance.\n\n"
                        f"Best regards,\n"
                        f"Lucious Stores")
        else:
            response = (f"Dear Customer,\n\n"
                        f"We were unable to process your order. Please check the product details in your order and try again.\n\n"
                        f"Best regards,\n"
                        f"Lucious Stores")

    return response

# Generate responses for each order in the order_status_df
order_responses = []
for index, order in order_status_df.iterrows():
    email_id = order['email_id']
    product_id = order['product_id']
    quantity = order['quantity']
    status = order['status']

    response = generate_email_response(email_id, product_id, quantity, status, products_df)
    order_responses.append({
        'email_id': email_id,
        'response': response
    })

# Create a DataFrame for the order responses
order_response_df = pd.DataFrame(order_responses)

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
data_to_upload = order_response_df.values.tolist()
header = ['email_id', 'response']
data_to_upload.insert(0, header)

# Update the worksheet with the data
worksheet = get_or_create_worksheet(spreadsheet, 'order-response')
worksheet.update('A1', data_to_upload)

print("Order response data uploaded successfully!")
