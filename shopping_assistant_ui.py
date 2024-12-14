import tkinter as tk
from tkinter import ttk
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import os
import requests
from datetime import datetime
import json
import threading
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service

class ShoppingAssistantUI:
    def __init__(self):
        self.root = None
        self.browser = None
        self.recording = False
        self.purchases = []
        self.init_gui()
        
    def init_gui(self):
        try:
            self.root = tk.Tk()
            self.root.title("Shopping Assistant")
            
            self.root.protocol("WM_DELETE_WINDOW", self.return_to_main_menu)
            
            window_width = 1080
            window_height = 700
            screen_width = self.root.winfo_screenwidth()
            screen_height = self.root.winfo_screenheight()
            position_top = int((screen_height - window_height) / 2)
            position_left = int((screen_width - window_width) / 2)
            self.root.geometry(f"{window_width}x{window_height}+{position_left}+{position_top}")

            # 控制按钮
            control_frame = tk.Frame(self.root)
            control_frame.pack(pady=10)

            self.start_button = tk.Button(control_frame, text="Start Recording", command=self.start_recording)
            self.start_button.pack(side=tk.LEFT, padx=5)

            self.stop_button = tk.Button(control_frame, text="Stop Recording", command=self.stop_recording)
            self.stop_button.pack(side=tk.LEFT, padx=5)
            self.stop_button.config(state=tk.DISABLED)

            # 记录列表
            self.tree = ttk.Treeview(self.root, columns=("Time", "Website", "Product", "Price", "Quantity", "Images"), show='headings')
            for col in ["Time", "Website", "Product", "Price", "Quantity", "Images"]:
                self.tree.heading(col, text=col)
                self.tree.column(col, width=150)
            self.tree.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

            # 状态标签
            self.status_label = tk.Label(self.root, text="Ready", anchor="w")
            self.status_label.pack(fill=tk.X, padx=10, pady=5)

            back_button = tk.Button(self.root, text="Back", command=self.return_to_main_menu)
            back_button.pack(pady=10)

            self.root.mainloop()

        except Exception as e:
            print(f"Error initializing GUI: {e}")
            self.return_to_main_menu()

    def start_recording(self):
        try:
            service = Service(ChromeDriverManager().install())
            self.browser = webdriver.Chrome(service=service)
            self.recording = True
            self.start_button.config(state=tk.DISABLED)
            self.stop_button.config(state=tk.NORMAL)
            self.status_label.config(text="Recording started...")
            
            # 在新线程中开始监控
            threading.Thread(target=self.monitor_pages, daemon=True).start()
            
        except Exception as e:
            self.status_label.config(text=f"Error starting browser: {e}")

    def stop_recording(self):
        self.recording = False
        if self.browser:
            self.browser.quit()
        self.start_button.config(state=tk.NORMAL)
        self.stop_button.config(state=tk.DISABLED)
        self.status_label.config(text="Recording stopped")
        self.save_purchases()

    def monitor_pages(self):
        while self.recording:
            try:
                current_url = self.browser.current_url
                if "1688.com" in current_url or "taobao.com" in current_url:
                    self.extract_product_info()
            except Exception as e:
                print(f"Error monitoring pages: {e}")
            self.root.after(2000)  # 每2秒检查一次

    def extract_product_info(self):
        try:
            if "1688.com" in self.browser.current_url:
                self.extract_1688_info()
            elif "taobao.com" in self.browser.current_url:
                self.extract_taobao_info()
        except Exception as e:
            print(f"Error extracting product info: {e}")

    def extract_1688_info(self):
        # 1688.com 特定的提取逻辑
        try:
            product_name = WebDriverWait(self.browser, 10).until(
                EC.presence_of_element_located((By.CLASS_NAME, "product-title"))
            ).text
            price = self.browser.find_element(By.CLASS_NAME, "price").text
            # 获取图片URL
            image_elements = self.browser.find_elements(By.CSS_SELECTOR, ".detail-gallery-img img")
            image_urls = [img.get_attribute("src") for img in image_elements]
            
            self.download_images(image_urls, product_name)
            self.add_purchase_record("1688.com", product_name, price, "1", image_urls)
            
        except Exception as e:
            print(f"Error extracting 1688 info: {e}")

    def extract_taobao_info(self):
        # 淘宝特定的提取逻辑
        try:
            product_name = WebDriverWait(self.browser, 10).until(
                EC.presence_of_element_located((By.CLASS_NAME, "tb-main-title"))
            ).text
            price = self.browser.find_element(By.CLASS_NAME, "tb-rmb-num").text
            # 获取图片URL
            image_elements = self.browser.find_elements(By.CSS_SELECTOR, ".tb-thumb img")
            image_urls = [img.get_attribute("src") for img in image_elements]
            
            self.download_images(image_urls, product_name)
            self.add_purchase_record("taobao.com", product_name, price, "1", image_urls)
            
        except Exception as e:
            print(f"Error extracting Taobao info: {e}")

    def download_images(self, image_urls, product_name):
        # 创建图片保存目录
        save_dir = os.path.join("product_images", product_name)
        os.makedirs(save_dir, exist_ok=True)

        for i, url in enumerate(image_urls):
            try:
                response = requests.get(url)
                if response.status_code == 200:
                    file_path = os.path.join(save_dir, f"image_{i+1}.jpg")
                    with open(file_path, 'wb') as f:
                        f.write(response.content)
            except Exception as e:
                print(f"Error downloading image {url}: {e}")

    def add_purchase_record(self, website, product, price, quantity, image_urls):
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        record = {
            "time": timestamp,
            "website": website,
            "product": product,
            "price": price,
            "quantity": quantity,
            "images": len(image_urls)
        }
        self.purchases.append(record)
        
        # 更新UI
        self.root.after(0, lambda: self.tree.insert("", 0, values=(
            timestamp, website, product, price, quantity, f"{len(image_urls)} images"
        )))

    def save_purchases(self):
        try:
            with open("purchases.json", "w", encoding="utf-8") as f:
                json.dump(self.purchases, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"Error saving purchases: {e}")

    def return_to_main_menu(self):
        self.stop_recording()
        if self.root:
            self.root.quit()
            self.root.destroy()
        import main
        main.ProductManager()
