import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
from datetime import datetime
import random
import csv
import time
import pyperclip
from threading import Thread

class ProductManagerUI:
    def __init__(self, products, save_products):
        self.products = products
        self.save_products = save_products
        self.current_input_step = 0
        self.current_product = {}
        self.filtered_products = products  # Track filtered products
        self.pasteboard_monitoring = False
        self.init_gui()

    def init_gui(self):
        self.root = tk.Tk()
        self.root.title("Product Management")
        
        window_width = 1080
        window_height = 700
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        position_top = int((screen_height - window_height) / 2)
        position_left = int((screen_width - window_width) / 2)
        self.root.geometry(f"{window_width}x{window_height}+{position_left}+{position_top}")
    
        # Create Search Bar (Separate from Data Entry)
        self.search_label = tk.Label(self.root, text="Search Product:")
        self.search_label.pack(pady=5)
    
        self.search_entry = tk.Entry(self.root, width=50)
        self.search_entry.pack(pady=5)
        self.search_entry.bind('<Return>', self.handle_search)  # Bind to Enter key for search
        self.search_entry.bind("<FocusIn>", self.stop_pasteboard_monitoring)  # Stop monitoring when focus in
        self.search_entry.bind("<FocusOut>", self.start_pasteboard_monitoring)  # Resume monitoring when focus out
    
        # Treeview to display products
        self.tree = ttk.Treeview(self.root, columns=("Name", "Cost", "Link", "Quantity", "Code", "LocationCode", "StockQuantity", "AddedDate"), show='headings')
        for col in ["Name", "Cost", "Link", "Quantity", "Code", "LocationCode", "StockQuantity", "AddedDate"]:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=100, anchor='center')
        self.tree.pack(fill=tk.BOTH, expand=True)
    
        # Add the event binding for column double-click
        self.tree.bind('<Double-1>', self.on_item_double_click)
    
        # Buttons Section (Separate)
        clear_button = tk.Button(self.root, text="Clear Database", command=self.clear_database)
        clear_button.pack(pady=10)
    
        delete_button = tk.Button(self.root, text="Delete Selected Entries", command=self.delete_selected_entries)
        delete_button.pack(pady=10)
    
        back_button = tk.Button(self.root, text="Back", command=self.return_to_main_menu)
        back_button.pack(pady=10)
    
        self.status_label = tk.Label(self.root, text="", anchor="w")
        self.status_label.pack(fill=tk.X, padx=10, pady=5)
    
        # Create Product Data Input Section
        self.entry_label = tk.Label(self.root, text="Enter Product Name:")
        self.entry_label.pack(pady=10)
    
        self.entry_field = tk.Entry(self.root, width=50)
        self.entry_field.pack(pady=5)
        self.entry_field.bind('<Return>', self.handle_entry)
        self.entry_field.bind("<FocusIn>", self.stop_pasteboard_monitoring)  # Stop monitoring when focus in
        self.entry_field.bind("<FocusOut>", self.start_pasteboard_monitoring)  # Resume monitoring when focus out
    
        self.refresh_tree()
    
        self.root.mainloop()

    def stop_pasteboard_monitoring(self, event=None):
        self.pasteboard_monitoring = False  # Stop monitoring

    def start_pasteboard_monitoring(self, event=None):
        if not self.pasteboard_monitoring:
            self.pasteboard_monitoring = True
            self.clear_pasteboard()
            self.monitor_pasteboard()  # Resume monitoring

    def monitor_pasteboard(self):
        def monitor():
            prev_clipboard = ""
            while self.pasteboard_monitoring:
                current_clipboard = pyperclip.paste()
                if current_clipboard != prev_clipboard:
                    prev_clipboard = current_clipboard
                    if self.current_input_step == 0:  # Product Name
                        self.entry_field.delete(0, tk.END)
                        self.entry_field.insert(0, current_clipboard)
                        self.handle_entry(None)
                    elif self.current_input_step == 2:  # Product Link
                        self.entry_field.delete(0, tk.END)
                        self.entry_field.insert(0, current_clipboard)
                        self.handle_entry(None)
                time.sleep(0.5)

        # Start monitoring in a separate thread
        thread = Thread(target=monitor, daemon=True)
        thread.start()
    def clear_pasteboard(self):
        pyperclip.copy("")  # Clear the clipboard initially

    def handle_search(self, event=None):
        search_query = self.search_entry.get().strip().lower()

        if search_query:
            self.filtered_products = [
                product for product in self.products if search_query in product["Name"].lower()
            ]
        else:
            self.filtered_products = self.products  # Show all products if search is empty

        self.refresh_tree()

    def refresh_tree(self):
        # Delete existing rows
        self.tree.delete(*self.tree.get_children())

        # Insert filtered products into the Treeview
        for product in self.filtered_products:
            self.tree.insert("", tk.END, values=tuple(product.get(key, "") for key in ["Name", "Cost", "Link", "Quantity", "Code", "LocationCode", "StockQuantity", "AddedDate"]))

    def on_column_double_click(self, event):
        # Identify the region where the double-click occurred
        region = self.tree.identify('region', event.x, event.y)
        
        if region == 'heading':  # Only trigger for column header
            # Identify the column by position
            col = self.tree.identify_column(event.x)  # This returns something like '#1', '#2', ...
            col = col.lstrip('#')  # Remove the '#' prefix from the column identifier
            self.adjust_column_width_on_double_click(col)  # Call function to adjust column width
    
    def adjust_column_width_on_double_click(self, col):
        # Get the column index
        col_index = int(col) - 1  # The identify_column method returns 1-based index (e.g., '#1' is index 1)
        
        # Get the maximum width based on the values in that column
        max_width = max(len(str(self.tree.item(child)['values'][col_index])) 
                        for child in self.tree.get_children())
        
        # Consider the header length as well
        max_width = max(max_width, len(self.tree.heading(self.tree['columns'][col_index])['text']))
        
        # Calculate new width (adjust multiplier if necessary)
        new_width = max_width * 10  # You can adjust the multiplier (10) based on your preference
        
        # Set the new column width
        self.tree.column(self.tree['columns'][col_index], width=new_width)
    def on_item_double_click(self, event):
        # Identify the region clicked
        region = self.tree.identify('region', event.x, event.y)
        
        if region == 'heading':
            # If the double-click happened on the column header, adjust the column width
            col = self.tree.identify_column(event.x)
            col = col.lstrip('#')  # Remove the '#' prefix
            self.adjust_column_width_on_double_click(col)
        elif region == 'cell':
            # If the double-click happened on a cell, edit the item
            item = self.tree.selection()  # Get the selected row/item
            if not item:
                return
            
            # Get the column index of the clicked cell
            col = self.tree.identify_column(event.x)  # Get the column identifier
            col_index = int(col.lstrip('#')) - 1  # Convert to 0-based index
            
            # Get the current value of the clicked cell
            item_values = self.tree.item(item[0], 'values')
            current_value = item_values[col_index]
            
            # Create an entry widget at the position of the clicked cell
            entry = tk.Entry(self.tree, width=20)
            entry.insert(0, current_value)
            entry.place(x=event.x, y=event.y, anchor="w")
            
            # Function to save the edited value
            def save_edit(event=None):
                new_value = entry.get()
                if new_value != current_value:
                    # Update the product data in the backend (product list)
                    self.update_product(item[0], col_index, new_value)
                    
                entry.destroy()  # Remove the entry widget after editing
                self.refresh_tree()  # Refresh Treeview to show the updated value
            
            # Bind the entry widget to save the edit when pressing Enter
            entry.bind('<Return>', save_edit)
            
            # Bind the entry widget to save the edit when losing focus (clicking outside)
            entry.bind('<FocusOut>', save_edit)
            
            entry.focus()  # Focus on the entry widget
    
    def update_product(self, item_id, col_index, new_value):
        # Get the corresponding product data based on the item (row) in the Treeview
        item_values = self.tree.item(item_id, 'values')
        code = item_values[4]  # Get the product code, assuming it's in the 5th column (index 4)
        
        # Find the product in the product list by matching the code
        for product in self.products:
            if product['Code'] == code:
                column_name = self.tree['columns'][col_index]
                # Update the correct field in the product dictionary
                product[column_name] = new_value
                break
        
        # Save the updated product list to the CSV
        self.save_products()  # Save updated list to CSV

    def generate_unique_code(self):
        while True:
            new_code = str(random.randint(1000000, 9999999))
            if all(product["Code"] != new_code for product in self.products):
                return new_code

    def handle_entry(self, event):
        user_input = self.entry_field.get().strip()
        if self.current_input_step == 0:
            if user_input:
                # Debugging: Print input to ensure it's being correctly captured
                print(f"Step 0 - Product Name: {user_input}")
                self.current_product["Name"] = user_input
                self.current_input_step += 1
                self.entry_label.config(text="Enter Product Cost:")
                self.entry_field.delete(0, tk.END)
            else:
                self.status_label.config(text="Error: Product name cannot be empty.")
        elif self.current_input_step == 1:
            try:
                cost = float(user_input)  # Ensure that cost is a valid number
                # Debugging: Print cost to ensure correct input
                print(f"Step 1 - Product Cost: {cost}")
                self.current_product["Cost"] = f"{cost:.2f}"
                self.current_input_step += 1
                self.entry_label.config(text="Enter Product Link:")
                self.entry_field.delete(0, tk.END)
            except ValueError:
                self.status_label.config(text="Error: Please enter a valid cost.")
        elif self.current_input_step == 2:
            if user_input:
                # Debugging: Print link to ensure it is correctly entered
                print(f"Step 2 - Product Link: {user_input}")
                self.current_product["Link"] = user_input
                self.current_input_step += 1
                self.entry_label.config(text="Enter Product Quantity:")
                self.entry_field.delete(0, tk.END)
            else:
                self.status_label.config(text="Error: Product link cannot be empty.")
        elif self.current_input_step == 3:
            try:
                quantity = int(user_input)
                # Debugging: Print quantity to ensure correct input
                print(f"Step 3 - Product Quantity: {quantity}")
                self.current_product["Quantity"] = str(quantity)
                self.current_product["Code"] = self.generate_unique_code()
                self.current_product["LocationCode"] = ""
                self.current_product["StockQuantity"] = "0"
                self.current_product["AddedDate"] = datetime.now().strftime("%Y-%m-%d")
                self.products.append(self.current_product)
                self.save_products()  # Save to the database
                self.filtered_products = self.products  # Reset filtered products to include the new product
                self.refresh_tree()  # Refresh Treeview to include the newly added product
                self.current_input_step = 0
                self.current_product = {}
                self.entry_label.config(text="Enter Product Name:")
                self.entry_field.delete(0, tk.END)
            except ValueError:
                self.status_label.config(text="Error: Please enter a valid quantity.")
    def delete_selected_entries(self):
        selected_items = self.tree.selection()
        if selected_items:
            product_codes_to_delete = []
            for item in selected_items:
                item_values = self.tree.item(item, 'values')
                if len(item_values) >= 5:  # Check to avoid index errors
                    product_code = item_values[4]
                    product_codes_to_delete.append(product_code)
                    self.tree.delete(item)  # Delete the row from the Treeview
    
            # Remove products from the main list
            self.products = [product for product in self.products if product['Code'] not in product_codes_to_delete]
            
            # Remove products from the filtered list if they were part of the current filter
            self.filtered_products = [product for product in self.filtered_products if product['Code'] not in product_codes_to_delete]
    
            # Update CSV file and refresh the Treeview to reflect changes
            self.save_products()  # Save updated list to CSV
            self.update_csv_file()  # Save to file
            self.refresh_tree()  # Refresh Treeview to show updated list

    def update_csv_file(self):
        try:
            with open('products.csv', mode='w', newline='') as file:
                fieldnames = ["Name", "Cost", "Link", "Quantity", "Code", "LocationCode", "StockQuantity", "AddedDate"]
                writer = csv.DictWriter(file, fieldnames=fieldnames)
                writer.writeheader()
                writer.writerows(self.products)
        except Exception as e:
            self.status_label.config(text=f"Error saving to CSV: {str(e)}")

    def clear_database(self):
        confirmation = simpledialog.askstring("Clear Database", "Type 'delete' to confirm:")
        if confirmation == 'delete':
            self.products.clear()
            self.filtered_products = self.products  # Clear filtered list as well
            self.save_products()  # Clear CSV file
            self.refresh_tree()  # Refresh Treeview to show empty list
        else:
            self.status_label.config(text="Database clearance cancelled.")

    def return_to_main_menu(self):
        self.root.destroy()  # Properly close the current window
        import main  # Assuming the main program logic exists in a file named main.py
        main.ProductManager()  # Initialize the main menu

# Example use:
if __name__ == "__main__":
    # Mock products data to initialize
    products = []
    
    def save_products():
        # Function to save products to CSV
        try:
            with open('products.csv', mode='w', newline='') as file:
                fieldnames = ["Name", "Cost", "Link", "Quantity", "Code", "LocationCode", "StockQuantity", "AddedDate"]
                writer = csv.DictWriter(file, fieldnames=fieldnames)
                writer.writeheader()
                writer.writerows(products)
        except Exception as e:
            print(f"Error saving to CSV: {str(e)}")

    app = ProductManagerUI(products, save_products)