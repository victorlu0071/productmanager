import csv
import os
from tkinter import messagebox
import chardet  # For automatic encoding detection

def load_products(file_path):
    products = []
    product_dict = {}

    if os.path.exists(file_path):
        try:
            # Detect the file encoding first
            encoding = detect_encoding(file_path)
            
            # Open the file with the detected encoding
            with open(file_path, newline='', encoding=encoding) as file:
                reader = csv.DictReader(file)
                products = list(reader)
                product_dict = {product['Code']: product for product in products}
        
        except FileNotFoundError:
            messagebox.showerror("Error", "The products file was not found.")
        except csv.Error:
            messagebox.showerror("Error", "The products file is corrupted or improperly formatted.")
        except UnicodeDecodeError:
            messagebox.showerror("Error", "The products file has an invalid encoding. Please check the file.")
        except Exception as e:
            messagebox.showerror("Error", f"An unexpected error occurred while loading products: {e}")

    return products, product_dict

def detect_encoding(file_path):
    """
    Attempts to detect the file encoding using the chardet library.
    """
    with open(file_path, 'rb') as f:
        raw_data = f.read()
        result = chardet.detect(raw_data)
        return result['encoding']  # Return the detected encoding

def save_products(file_path, products):
    retries = 3
    for attempt in range(retries):
        try:
            with open(file_path, mode='w', newline='', encoding='utf-8') as file:
                writer = csv.DictWriter(file, fieldnames=["Name", "Cost", "Link", "Quantity", "Code", "LocationCode", "StockQuantity", "AddedDate"])
                writer.writeheader()
                writer.writerows(products)
            return
        except PermissionError:
            if attempt < retries - 1:
                import time
                time.sleep(1)
            else:
                messagebox.showerror("Error", f"Permission denied: {file_path}. Ensure the file is not open in another program.")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save products: {e}")
