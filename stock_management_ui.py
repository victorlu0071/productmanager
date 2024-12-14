import tkinter as tk
from tkinter import ttk

class StockManagerUI:
    def __init__(self, products, product_dict, save_products):
        self.products = products
        self.product_dict = product_dict
        self.save_products = save_products
        self.current_input_step = 0
        self.current_stock_entry = {}
        self.root = None
        self.init_gui()

    def init_gui(self):
        try:
            self.root = tk.Tk()
            self.root.title("Stock Management")
            
            # 确保窗口关闭时返回主菜单
            self.root.protocol("WM_DELETE_WINDOW", self.return_to_main_menu)
            
            window_width = 1080
            window_height = 700
            screen_width = self.root.winfo_screenwidth()
            screen_height = self.root.winfo_screenheight()
            position_top = int((screen_height - window_height) / 2)
            position_left = int((screen_width - window_width) / 2)
            self.root.geometry(f"{window_width}x{window_height}+{position_left}+{position_top}")

            self.entry_label = tk.Label(self.root, text="Enter Product Code:")
            self.entry_label.pack()

            self.entry_field = tk.Entry(self.root, width=50)
            self.entry_field.pack()
            self.entry_field.bind('<Return>', self.handle_entry)
            self.root.bind('<Key>', self.focus_entry_on_keypress)

            self.tree = ttk.Treeview(self.root, columns=("Code", "ProductName", "LocationCode", "StockQuantity"), show='headings')
            self.tree.heading("Code", text="Product Code")
            self.tree.heading("ProductName", text="Product Name")
            self.tree.heading("LocationCode", text="Location Code")
            self.tree.heading("StockQuantity", text="Stock Quantity")
            self.tree.pack(fill=tk.BOTH, expand=True)

            self.tree.bind('<Double-1>', self.edit_selected_item)

            back_button = tk.Button(self.root, text="Back", command=self.return_to_main_menu)
            back_button.pack(pady=10)

            self.status_label = tk.Label(self.root, text="", anchor="w")
            self.status_label.pack(fill=tk.X, padx=10, pady=5)

            self.update_tree_view()

            self.root.mainloop()
        except tk.TclError as e:
            print(f"Error initializing GUI: {e}")
            self.return_to_main_menu()

    def update_tree_view(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
        for product_code, product in self.product_dict.items():
            self.tree.insert("", tk.END, values=(product_code, product.get("Name", ""), product.get("LocationCode", ""), product.get("StockQuantity", 0)))

    def edit_selected_item(self, event):
        selected_item = self.tree.selection()[0]
        column_index = self.tree.identify_column(event.x)
        column_number = int(column_index[1:]) - 1  # Convert to zero-indexed
        if column_number >= 0:
            entry_window = tk.Toplevel(self.root)
            entry_window.title("Edit Value")

            current_value = self.tree.item(selected_item, 'values')[column_number]
            entry_field = tk.Entry(entry_window, width=50)
            entry_field.insert(0, current_value)
            entry_field.pack(pady=10)

            def save_changes():
                new_value = entry_field.get()
                try:
                    if column_number == 3:  # StockQuantity
                        new_value = float(new_value)
                    values = list(self.tree.item(selected_item, 'values'))
                    values[column_number] = new_value
                    self.tree.item(selected_item, values=values)
                    product_code = values[0]
                    if product_code in self.product_dict:
                        if column_number == 1:
                            self.product_dict[product_code]["Name"] = new_value
                        elif column_number == 2:
                            self.product_dict[product_code]["LocationCode"] = new_value
                        elif column_number == 3:
                            self.product_dict[product_code]["StockQuantity"] = new_value
                    self.save_products()
                    self.update_tree_view()
                    self.status_label.config(text="Success: Changes saved successfully!")
                    entry_window.destroy()
                except ValueError:
                    self.status_label.config(text="Error: Stock quantity must be a valid number.")
                    return  # Do not close the window if there's an error

            save_button = tk.Button(entry_window, text="Save", command=save_changes)
            save_button.pack(pady=5)

    def focus_entry_on_keypress(self, event):
        if self.root.focus_get() != self.entry_field:
            self.entry_field.focus()
            self.entry_field.delete(0, tk.END)  # Clear existing text
            self.entry_field.insert(tk.END, event.char)

    def handle_entry(self, event):
        if not hasattr(self, 'current_input_step'):
            self.current_input_step = 0
            self.current_stock_entry = {}

        user_input = self.entry_field.get().strip()

        if self.current_input_step == 0:
            if user_input in self.product_dict:
                self.current_stock_entry["Code"] = user_input
                self.current_input_step += 1
                self.entry_field.delete(0, tk.END)
                self.entry_label.config(text="Enter Location Code:")
                self.status_label.config(text="")
            else:
                self.status_label.config(text="Error: Product code not found.")

        elif self.current_input_step == 1:
            self.current_stock_entry["LocationCode"] = user_input
            self.current_input_step += 1
            self.entry_field.delete(0, tk.END)
            self.entry_label.config(text="Enter Stock Quantity:")
            self.status_label.config(text="")

        elif self.current_input_step == 2:
            try:
                self.current_stock_entry["StockQuantity"] = float(user_input)
                product = self.product_dict[self.current_stock_entry["Code"]]
                product["LocationCode"] = self.current_stock_entry["LocationCode"]
                product["StockQuantity"] = self.current_stock_entry["StockQuantity"]

                self.save_products()
                self.update_tree_view()

                self.status_label.config(text="Success: Stock entry recorded successfully!")
                self.current_input_step = 0
                self.current_stock_entry = {}
                self.entry_field.delete(0, tk.END)
                self.entry_label.config(text="Enter Product Code:")
            except ValueError:
                self.status_label.config(text="Error: Stock quantity must be a valid number.")

    def return_to_main_menu(self):
        try:
            if self.root:
                self.root.quit()
                self.root.destroy()
            import main
            main.ProductManager()
        except Exception as e:
            print(f"Error returning to main menu: {e}")
            # 确保在发生错误时也能退出
            if self.root:
                self.root.destroy()
