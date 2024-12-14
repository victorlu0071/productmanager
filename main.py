import tkinter as tk
from file_operations import load_products, save_products
from product_management_ui import ProductManagerUI
from stock_management_ui import StockManagerUI
from datetime import datetime, timedelta
import os
import csv

class ProductManager:
    def __init__(self):
        self.file_path = "products.csv"
        self.products, self.product_dict = load_products(self.file_path)
        self.init_gui()


    def init_gui(self):
        self.root = tk.Tk()
        self.root.title("Main Menu")

        # Center the window
        window_width = 1080
        window_height = 700
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        position_top = int((screen_height - window_height) / 2)
        position_left = int((screen_width - window_width) / 2)
        self.root.geometry(f"{window_width}x{window_height}+{position_left}+{position_top}")

        # Create a frame for button layout
        button_frame = tk.Frame(self.root)
        button_frame.pack(expand=True)

        button1 = tk.Button(button_frame, text="Product Management", command=self.open_product_manager, font=("Helvetica", 14), width=20, height=2)
        button1.pack(side=tk.LEFT, padx=100)

        button2 = tk.Button(button_frame, text="Stock Management", command=self.open_stock_manager, font=("Helvetica", 14), width=20, height=2)
        button2.pack(side=tk.RIGHT, padx=100)

        # Add exit button
        exit_button = tk.Button(self.root, text="Exit", command=self.exit_program, font=("Helvetica", 14), width=20, height=2)
        exit_button.pack(pady=20)

        self.root.mainloop()

    def exit_program(self):
        self.root.destroy()

    def open_product_manager(self):
        self.root.destroy()
        ProductManagerUI(self.products, lambda: save_products(self.file_path, self.products))

    def open_stock_manager(self):
        self.root.destroy()
        StockManagerUI(self.products, self.product_dict, lambda: save_products(self.file_path, self.products))

if __name__ == "__main__":
    ProductManager()
