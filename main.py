# =============================================================================
# API // bigcartel.com // Unshipped orders data process
# =============================================================================

# This Python script is designed to interact with the Big Cartel API to process 
# unshipped orders data. It fetches relevant data from the API, organizes it, and 
# performs various operations such as generating invoices, creating finance files, 
# and compiling order lists.This Python script is designed to interact with the 
# Big Cartel API to process unshipped orders data. It fetches relevant data from 
# the API, organizes it, and performs various operations such as generating 
# invoices, creating finance files, and compiling order lists.


# Import Libraries
import customtkinter
import requests
import json
import keyring
import pandas as pd
from requests.auth import HTTPBasicAuth
from docxtpl import DocxTemplate
from docx2pdf import convert
import time

# Global default settings
pd.options.mode.chained_assignment = None

# Definition of variables
subdomain =     "bect"
subdomain_1 =   "api"
excel_file_path = "/Users/username/Desktop/Code/API_Project/Finanzen_Testfile.xlsx"
excel_file_path_temp = "/Users/username/Desktop/Code/API_Project/Temp.xlsx"
excel_file_path_orders = "/Users/username/Desktop/Code/API_Project/Orders.xlsx"
excel_file_path_finance = "/Users/username/Desktop/Code/API_Project/Finance.xlsx"
starting_number = int(input("What's the current Index in the main Finance file? "))
sheet_name = 'Bestellungen'

# APPLICATION: Definition of initial application layout
customtkinter.set_appearance_mode("dark")
customtkinter.set_default_color_theme("dark-blue")

app = customtkinter.CTk()
app.geometry("400x240")
app.title("BECT Dashboard")

label = customtkinter.CTkLabel(master=app, text="Data features only unshipped orders! \n Please Note: All data will be placed into the directory folder.")
label.place(relx=0.5, rely=0.2, anchor=customtkinter.CENTER)

# Get password from keychain
password_keychain = keyring.get_password("bigcartel.com", "bect.bigcartel.com")

# Definition of the API root URL
api_root_url = f"https://{subdomain_1}.bigcartel.com"

# Specify the root URL including the header (-> more info in the big cartel documentation)
account_url = f"{api_root_url}/v1/accounts"
account_headers = {
    "Accept": "application/vnd.api+json",
    "Content-type": "application/vnd.api+json",
    "User-Agent": "Invoice_Generator (https://bect.ch/contact/)"
}

# Requests with exception handling 
try:
    # GET account information
    account_response = requests.get(account_url, headers=account_headers, 
                                    auth=HTTPBasicAuth(subdomain, password_keychain))

    # IF succesful
    if account_response.status_code == 200:
        #Print status code
        print("\n",f"{account_response.status_code}: Success to get Account ID","\n")
        
        # Get account data in json format
        account_data = account_response.json()

        # Extraction of the ID of the account
        account_id = account_data['data'][0]['id']

        # URL to get all orders with shipping_status of unshipped
        orders_url_unshipped = f"{api_root_url}/v1/accounts/{account_id}/orders?filter[shipping_status]=unshipped"
 
        # Make the GET request
        response_unshipped = requests.get(orders_url_unshipped, headers=account_headers, auth=HTTPBasicAuth(subdomain, password_keychain))

        # Exception handling: If successful... make next move
        if response_unshipped.status_code == 200:
            print(f"{response_unshipped.status_code}: Success to get unshipped order data", "\n")
            # Parse the JSON data response
            json_data = response_unshipped.json()
            
            max_line_items = max(len(order['relationships']['items']['data']) for order in json_data['data'])
            # Reorder JSON file
            data_file = []
            
            # For loop to go through the whole json file
            for order in json_data['data']:
                # Extraction of all order information
                order_info = {
                    'id': order['id'],
                    'item_count': order['attributes']['item_count'],
                    'item_total': order['attributes']['item_total'],
                    'discount_total': order['attributes']['discount_total'],
                    'shipping_total': order['attributes']['shipping_total'],
                    'tax_total': order['attributes']['tax_total'],
                    'total': order['attributes']['total'],
                    'customer_first_name': order['attributes']['customer_first_name'],
                    'customer_last_name': order['attributes']['customer_last_name'],
                    'customer_email': order['attributes']['customer_email'],
                    'customer_phone_number': order['attributes']['customer_phone_number'],
                    'customer_opted_in_to_marketing': order['attributes']['customer_opted_in_to_marketing'],
                    'customer_note': order['attributes']['customer_note'],
                    'shipping_address_1': order['attributes']['shipping_address_1'],
                    'shipping_address_2': order['attributes']['shipping_address_2'],
                    'shipping_city': order['attributes']['shipping_city'],
                    'shipping_state': order['attributes']['shipping_state'],
                    'shipping_zip': order['attributes']['shipping_zip'],
                    'shipping_status': order['attributes']['shipping_status'],
                    'payment_status': order['attributes']['payment_status'],
                    'created_at': order['attributes']['created_at'],
                    'updated_at': order['attributes']['updated_at'],
                    'completed_at': order['attributes']['completed_at']
                }
            
                # Extraction of the ID in relationship to get included data (order_line_items)
                item_ids = [item_data['id'] for item_data in order['relationships']['items']['data']]
            
                # Find with the ID in relationship the ID in included items and extract info
                for i in range(max_line_items):
                    item_id = item_ids[i] if i < len(item_ids) else None
                    order_line_item = next(
                        (item for item in json_data['included'] if item['id'] == item_id and item['type'] == 'order_line_items'),
                        None
                    )
            
                    if order_line_item:
                        # Extraction of relevant data if the ID is matching
                        item_details = {
                            'product_name': order_line_item['attributes']['product_name'],
                            'product_option_name': order_line_item['attributes']['product_option_name'],
                            'quantity': order_line_item['attributes']['quantity'],
                            'price': order_line_item['attributes']['price'],
                            'total': order_line_item['attributes']['total'],
                            'image_url': order_line_item['attributes']['image_url']
                        }
            
                        # Update the order_info dictionary with item_details
                        for key, value in item_details.items():
                            order_info[f'order_line_item_{i}_{key}'] = value
            
                # Appending the order information to the data_file
                data_file.append(order_info)
            
            # Convert to DataFrame
            orders_df = pd.DataFrame(data_file)
            
            if 'order_line_item_0_quantity' in orders_df.columns:
                orders_df['order_line_item_0_quantity'].fillna(0, inplace=True)
                orders_df['order_line_item_0_quantity'] = orders_df['order_line_item_0_quantity'].astype(int)
            else:
                pass
            
            if 'order_line_item_1_quantity' in orders_df.columns:
                orders_df['order_line_item_1_quantity'].fillna(0, inplace=True)
                orders_df['order_line_item_1_quantity'] = orders_df['order_line_item_1_quantity'].astype(int)
            else:
                pass
            
            if 'order_line_item_2_quantity' in orders_df.columns:
                orders_df['order_line_item_2_quantity'].fillna(0, inplace=True)
                orders_df['order_line_item_2_quantity'] = orders_df['order_line_item_2_quantity'].astype(int)
            else:
                pass
            
            if 'order_line_item_3_quantity' in orders_df.columns:
                orders_df['order_line_item_3_quantity'].fillna(0, inplace=True)
                orders_df['order_line_item_3_quantity'] = orders_df['order_line_item_3_quantity'].astype(int)
            else:
                pass
                
            if 'order_line_item_4_quantity' in orders_df.columns:
                orders_df['order_line_item_4_quantity'].fillna(0, inplace=True)
                orders_df['order_line_item_4_quantity'] = orders_df['order_line_item_4_quantity'].astype(int)
            else:
                pass
            
            orders_df['completed_at'] = pd.to_datetime(orders_df['completed_at'])
            orders_df['updated_at'] = pd.to_datetime(orders_df['updated_at'])
            orders_df['created_at'] = pd.to_datetime(orders_df['created_at'])
            
            orders_df['completed_at'] = orders_df['completed_at'].dt.strftime('%d.%m.%Y')
            orders_df['updated_at'] = orders_df['updated_at'].dt.strftime('%d.%m.%Y')
            orders_df['created_at'] = orders_df['created_at'].dt.strftime('%d.%m.%Y')
            
            # Save temp orders dataframe to new excel file
            orders_df.to_excel(excel_file_path_temp, index = False)
            print(f"Success: Data saved to {excel_file_path_temp}")


        else:
            # If unsuccessful: print error
            print(f"Error: {response_unshipped.status_code} - {response_unshipped.text}")

    else:
        # If unsuccessful: print error
        print(f"Error: {account_response.status_code} - {account_response.text}")

except requests.RequestException as e:
    # If unsuccessful: print error
    print(f"Request Exception: {e}")

def finance_data():
    orders_df
    finance_df = orders_df
    
    column_order = [
        'id',
        'created_at',
        'customer_first_name',
        'customer_last_name',
        'shipping_address_1',
        'shipping_address_2',
        'shipping_zip',
        'shipping_city',
        'order_line_item_0_quantity',
        'order_line_item_0_product_name',
        'order_line_item_0_product_option_name',
        'order_line_item_0_price',
        'order_line_item_0_total',
        'order_line_item_1_quantity',
        'order_line_item_1_product_name',
        'order_line_item_1_product_option_name',
        'order_line_item_1_price',
        'order_line_item_1_total',
        'order_line_item_2_quantity',
        'order_line_item_2_product_name',
        'order_line_item_2_product_option_name',
        'order_line_item_2_price',
        'order_line_item_2_total',
        'order_line_item_3_quantity',
        'order_line_item_3_product_name',
        'order_line_item_3_product_option_name',
        'order_line_item_3_price',
        'order_line_item_3_total',
        'order_line_item_4_quantity',
        'order_line_item_4_product_name',
        'order_line_item_4_product_option_name',
        'order_line_item_4_price',
        'order_line_item_4_total',
        'item_total',
        'shipping_total',
        'discount_total',
        'total',
        'item_count'
    ]

    # Add in the missing columns
    missing_columns = set(column_order) - set(finance_df.columns)
    for col in missing_columns:
        finance_df[col] = None

    # Reorder the columns based on the column order
    finance_df = finance_df[column_order]
    
    # Ascending order of the dataframe based on the created_at value
    finance_df = finance_df.sort_values(by='created_at')
    
    # Add in the index at the fron column (starting with the value given as input)
    finance_df.insert(0, 'Index', range(starting_number, starting_number + len(finance_df)))
    
    # Saving the data to a temporary file with examption handling
    try:
        finance_df.to_excel(excel_file_path_finance, index=False)
        print(f"Success: Data saved to {excel_file_path_finance}")
    except Exception as e:
        print(f"Error: {e}")


def orderlist_data():
    orderlist_df = orders_df
    
    column_order_1 = [
        'order_line_item_0_product_name',
        'order_line_item_0_product_option_name',
        'order_line_item_0_quantity',
        'order_line_item_1_product_name',
        'order_line_item_1_product_option_name',
        'order_line_item_1_quantity',
        'order_line_item_2_product_name',
        'order_line_item_2_product_option_name',
        'order_line_item_2_quantity',
        'order_line_item_3_product_name',
        'order_line_item_3_product_option_name',
        'order_line_item_3_quantity',
        'order_line_item_4_product_name',
        'order_line_item_4_product_option_name',
        'order_line_item_4_quantity'
    ]
    
    # Add in the missing columns
    missing_columns = set(column_order_1) - set(orderlist_df.columns)
    
    for col in missing_columns:
        orderlist_df[col] = None
    
    # Reorder the columns based on the column order
    orderlist_df = orderlist_df[column_order_1]
    
    
    # Calculation of the total order data
    
    # Initialize dic(kpic)
    product_option_quantities = {}
    
    for index, row in orderlist_df.iterrows():
        for i in range(5):
            product_name_col = f'order_line_item_{i}_product_name'
            option_name_col = f'order_line_item_{i}_product_option_name'
            quantity_col = f'order_line_item_{i}_quantity'
    
            product_name = row[product_name_col]
            option_name = row[option_name_col]
            quantity = row[quantity_col]
    
            if pd.notna(product_name):
                key = (product_name, option_name)
                product_option_quantities[key] = product_option_quantities.get(key, 0) + quantity if pd.notna(quantity) else 0
            else:
                pass
            
    # Convert the dictionary to a list
    result_list = [{'Product_Name': key[0], 'Option_Name': key[1], 'Total_Quantity': value} for key, value in product_option_quantities.items()]
    result_df = pd.DataFrame(result_list)
    result_df = result_df.sort_values('Product_Name')
    result_df    
    
    try:
        result_df.to_excel(excel_file_path_orders, index=False)
        print(f"Success: Data saved to {excel_file_path_orders}")
    except Exception as e:
        print(f"Error: {e}")
                    
def invoice_generation():
    print("Printing Invoices...")
    
    # Gloabal variables
    global starting_number
    global orders_df
    
    # Sorting data
    orders_df = orders_df.sort_values(by='created_at')
    orders_df.fillna('', inplace=True)
    tpl = DocxTemplate('/Users/username/Desktop/Code/API_Project/invoice_template.docx')
    
    # Get the necessary information
    for index, row in orders_df.iterrows():
        first_name = row['customer_first_name']
        last_name = row['customer_last_name']
        address_1 = row['shipping_address_1']
        address_2 = row['shipping_address_2']
        zip_code = row['shipping_zip']
        city = row['shipping_city']
        created_at = row['created_at']
        customer_id = row['id']
        item_total = row['item_total']
        discount_total = row['discount_total']
        shipping_total = row['shipping_total']
        total = row['total']
        starting_number += 1
        
        if not address_2:
            pass
        else:
            address_2 = int(address_2)
        if not shipping_total:
            pass
        else:
            shipping_total = '{:.2f}'.format(float(shipping_total))
        if not total:
            pass
        else:
            total = '{:.2f}'.format(float(total))
        if not item_total:
            pass
        else:
            item_total = '{:.2f}'.format(float(item_total))
        if not discount_total:
            pass
        else:
            discount_total = '{:.2f}'.format(float(discount_total))
            
        
        order_lines = []
        for i in range(6):
            product_name_col = f'order_line_item_{i}_product_name'
            product_option_col = f'order_line_item_{i}_product_option_name'
            quantity_col = f'order_line_item_{i}_quantity'
            price_col = f'order_line_item_{i}_price'
            
            if price_col in row and row[price_col]:
                price = '{:.2f}'.format(float(row[price_col]))

            else:
                print(f"Abnormal condition in {starting_number}: Empty cell detected and skipped")
                continue
            
            if product_name_col in row:
                order_lines.append({
                    'index': i + 1,
                    'product_name': row[product_name_col],
                    'product_option': row[product_option_col],
                    'quantity': row[quantity_col],
                    'price': price,
                    'total': '{:.2f}'.format(float(price) * float(row[quantity_col])),
                })
            
        context = {
            'first_name': first_name,
            'last_name': last_name,
            'address_1': address_1,
            'address_2': address_2,
            'zip': zip_code,
            'city': city,
            'created_at': created_at,
            'id': customer_id,
            'order_number': starting_number,
            'item_total': item_total,
            'discount_total': discount_total,
            'shipping_total': shipping_total,
            'total': total,
            'order_lines': order_lines
        }
        
        # Rendering and saving the docx. documents
        output_path = f'/Users/username/Desktop/Code/API_Project/temp_invoices/Rechnungsset {starting_number}.docx'
        
        tpl.render(context)
        tpl.save(output_path)


    print("Transformation to PDF in progress...")
    time.sleep(3)
    convert('/Users/username/Desktop/Code/API_Project/temp_invoices', '/Users/username/Desktop/Code/API_Project/pdf_invoices')
    print("Successfully saved")
        
# APPLICATION: Button definition
button = customtkinter.CTkButton(master=app, text="Print Financefile", command=finance_data)
button.place(relx=0.5, rely=0.5, anchor=customtkinter.CENTER, y = -20)

button = customtkinter.CTkButton(master=app, text="Print Orderlist", command=orderlist_data)
button.place(relx=0.5, rely=0.5, anchor=customtkinter.CENTER, y = 20)

button = customtkinter.CTkButton(master=app, text="Print Invoices", command=invoice_generation)
button.place(relx=0.5, rely=0.5, anchor=customtkinter.CENTER, y = 60)

app.mainloop()