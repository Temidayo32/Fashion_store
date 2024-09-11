import pandas as pd
from IPython.display import display

def read_data_frame(document_id, sheet_name):
    export_link = f"https://docs.google.com/spreadsheets/d/{document_id}/gviz/tq?tqx=out:csv&sheet={sheet_name}"
    return  pd.read_csv(export_link, header=0)

document_id = '1F6tf8E1R5Sga2MgJlpUAH_h7GK-UI2fVyju9aPu3FPY'
products_df = read_data_frame(document_id, 'products')
emails_df = read_data_frame(document_id, 'emails')
