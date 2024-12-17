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
import time
import openpyxl
from openpyxl.drawing.image import Image as XLImage
from PIL import Image
import io
import random
import tkinter.messagebox as messagebox

class ShoppingAssistantUI:
    def __init__(self):
        self.root = None
        self.browser = None
        self.recording = False
        self.purchases = []
        self.product_cache = {}  # 用于存储产品页面的图片URL
        self.cached_urls = set()  # 添加URL缓存集合
        self.current_url = None  # 添加当前URL记录
        self.processed_products = set()  # 添加已处理商品记录
        self.last_product_id = None  # 添加最后访问的产品ID记录
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

            self.start_button = tk.Button(control_frame, text="Start Browser", command=self.start_recording)
            self.start_button.pack(side=tk.LEFT, padx=5)

            self.continue_button = tk.Button(control_frame, text="Continue Recording", command=self.continue_recording)
            self.continue_button.pack(side=tk.LEFT, padx=5)
            self.continue_button.config(state=tk.DISABLED)

            self.stop_button = tk.Button(control_frame, text="Stop Recording", command=self.stop_recording)
            self.stop_button.pack(side=tk.LEFT, padx=5)
            self.stop_button.config(state=tk.DISABLED)

            # 记录列表
            self.tree = ttk.Treeview(self.root, columns=("Time", "Website", "Product", "Price"), show='headings')
            for col in ["Time", "Website", "Product", "Price"]:
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
            # 设置Chrome选项
            chrome_options = webdriver.ChromeOptions()
            
            # 设置用户数据目录
            user_data_dir = os.path.join(os.getcwd(), "chrome_user_data")
            os.makedirs(user_data_dir, exist_ok=True)
            chrome_options.add_argument(f"user-data-dir={user_data_dir}")
            
            # 添加其他选项
            chrome_options.add_argument("--no-sandbox")
            chrome_options.add_argument("--disable-dev-shm-usage")
            chrome_options.add_argument("--disable-gpu")
            chrome_options.add_argument("--disable-features=TranslateUI")
            chrome_options.add_argument("--disable-extensions")
            chrome_options.add_argument("--disable-popup-blocking")
            
            # 添加WebGL相关的选项
            chrome_options.add_argument("--enable-unsafe-swiftshader")
            chrome_options.add_argument("--ignore-gpu-blocklist")
            chrome_options.add_argument("--enable-gpu-rasterization")
            
            # 初始化浏览器
            service = Service(ChromeDriverManager().install())
            self.browser = webdriver.Chrome(service=service, options=chrome_options)
            
            # 首先访问登录页面
            print("Accessing login page...")
            self.browser.get("https://login.1688.com/member/signin.htm")
            
            # 更新状态标签
            self.status_label.config(text="请在浏览器中登录，然后点击 'Continue Recording' 按钮...")
            
            # 禁用开始按钮，启用继续按钮
            self.start_button.config(state=tk.DISABLED)
            self.continue_button.config(state=tk.NORMAL)
            
        except Exception as e:
            self.status_label.config(text=f"Error starting browser: {e}")
            print(f"Error details: {str(e)}")
            # 如果出错，重新启用开始按钮
            self.start_button.config(state=tk.NORMAL)

    def continue_recording(self):
        try:
            # 确认登录状态
            time.sleep(2)
            
            self.recording = True
            self.continue_button.config(state=tk.DISABLED)
            self.stop_button.config(state=tk.NORMAL)
            self.status_label.config(text="Recording started...")
            
            # 在新线程中开始监控
            threading.Thread(target=self.monitor_pages, daemon=True).start()
            
        except Exception as e:
            self.status_label.config(text=f"Error starting recording: {e}")

    def stop_recording(self):
        try:
            self.recording = False
            if self.browser:
                try:
                    self.browser.quit()
                except:
                    pass  # 忽略关闭浏览器时的错误
            self.browser = None
            self.current_url = None  # 重置当前URL
            self.processed_products.clear()  # 重置处理态
            self.start_button.config(state=tk.NORMAL)
            self.continue_button.config(state=tk.DISABLED)
            self.stop_button.config(state=tk.DISABLED)
            self.status_label.config(text="Recording stopped")
            self.save_purchases()
        except Exception as e:
            print(f"Error stopping recording: {e}")

    def monitor_pages(self):
        def wait_for_page_load():
            try:
                WebDriverWait(self.browser, 10).until(
                    lambda driver: driver.execute_script("return document.readyState") == "complete"
                )
                return True
            except Exception as e:
                print(f"Error waiting for page load: {e}")
                return False

        last_url = None
        current_handle = None
        while self.recording:
            try:
                # 检查浏览器是否还在运行
                handles = self.browser.window_handles
                if not handles:
                    print("All browser windows closed. Stopping recording...")
                    self.root.after(0, self.stop_recording)
                    break

                # 获取当前活动标签页
                try:
                    new_handle = self.browser.current_window_handle
                    if new_handle != current_handle:
                        print(f"Switched to new tab: {new_handle}")
                        current_handle = new_handle
                        last_url = None  # 重置URL以强制检查新标签页
                except:
                    print("No active window handle found, waiting...")
                    time.sleep(0.5)
                    continue

                # 获取当前标签页的URL
                try:
                    current_url = self.browser.current_url
                except:
                    print("Could not get URL, retrying...")
                    time.sleep(0.5)
                    continue

                # 只在URL改变或标签页切换时处理
                if current_url != last_url:
                    print(f"Processing URL in tab {current_handle}: {current_url}")
                    last_url = current_url

                    # 等待页面加载完成
                    if wait_for_page_load():
                        # 检查当前页面类型
                        if "detail.1688.com/offer" in current_url:
                            # 获取商品ID
                            product_id = self.extract_product_id(current_url)
                            if product_id:
                                # 如果是新产品，清除之前的缓存
                                if product_id != self.last_product_id:
                                    print(f"New product detected in tab {current_handle}, clearing cache...")
                                    self.product_cache.clear()
                                    self.cached_urls.clear()
                                    self.last_product_id = product_id

                                print(f"Found product page: {product_id}")
                                # 只缓存图片
                                self.cache_product_images()
                            else:
                                print(f"Invalid product ID in tab {current_handle}")

                        elif "order.1688.com/order/smart_make_order.htm" in current_url:
                            print(f"Found order page in tab {current_handle}, recording product info...")
                            # 在购买页面记录产品信息
                            if self.last_product_id:
                                self.record_product_info()
                            else:
                                print("No product ID found in cache")

                        elif "trade.1688.com/order/trade_flow.htm" in current_url:
                            try:
                                # 等待页面加载并检查是否包含"收银台"文字
                                WebDriverWait(self.browser, 10).until(
                                    lambda driver: "收银台" in driver.page_source
                                )
                                print(f"Found checkout page in tab {current_handle}, saving to database...")
                                # 遍历所有规格的缓存记录
                                saved_any = False
                                for cache_key in list(self.product_cache.keys()):
                                    if cache_key.startswith(self.last_product_id):
                                        if self.save_to_database(cache_key):
                                            saved_any = True
                                if not saved_any:
                                    print("No valid product information found in cache")
                            except Exception as e:
                                print(f"Error checking checkout page: {e}")

            except Exception as e:
                if "no such window" in str(e):
                    print(f"Lost tab {current_handle}, resetting...")
                    current_handle = None
                    last_url = None
                    time.sleep(0.5)
                else:
                    print(f"Error in monitor loop: {e}")
                    time.sleep(0.5)

            time.sleep(0.5)

    def record_product_info(self):
        try:
            # 等待页面加载
            WebDriverWait(self.browser, 10).until(
                EC.presence_of_element_located((By.CLASS_NAME, "offer-link"))
            )

            # 获取产品名称
            product_name = self.browser.find_element(By.CLASS_NAME, "offer-link").text.strip()
            
            # 获取所有cargo-inner元素（每个代表一个产品规格组合）
            cargo_inners = self.browser.find_elements(By.CLASS_NAME, "cargo-inner")
            print(f"Found {len(cargo_inners)} cargo-inner elements")
            
            # 为每个规格组合处理信息
            if self.last_product_id:
                # 获取原始图片信息
                original_cache = self.product_cache.get(self.last_product_id, {})
                image_urls = original_cache.get('images', [])
                
                # 如果有图片，为每个规格创建独立的图片目录
                if image_urls:
                    # 遍历每个cargo-inner
                    for cargo_inner in cargo_inners:
                        try:
                            # 获取价格 (直接从cargo-inner下获取)
                            price_element = cargo_inner.find_element(By.CLASS_NAME, "cargo-unit-price")
                            price = price_element.text.strip()
                        except:
                            print("Could not find price in cargo-inner")
                            price = "N/A"
                        
                        try:
                            # 获取cargo-info
                            cargo_info = cargo_inner.find_element(By.CLASS_NAME, "cargo-info")
                            # 获取cargo-desc
                            cargo_desc = cargo_info.find_element(By.CLASS_NAME, "cargo-desc")
                            # 获取cargo-spec
                            cargo_spec = cargo_desc.find_element(By.CLASS_NAME, "cargo-spec")
                            # 获取规格项
                            spec_items = cargo_spec.find_elements(By.CLASS_NAME, "spec-item")
                            
                            # 收集规格信息
                            current_specs = []
                            for item in spec_items:
                                spec_text = item.text.strip()
                                if spec_text:
                                    # 移除末尾的分号
                                    spec_text = spec_text.rstrip(";")
                                    current_specs.append(spec_text)
                            
                            # 如果找到规格，创建规格组合
                            if current_specs:
                                spec_combination = ";".join(current_specs)
                            else:
                                # 如果没有找到规格，使用默认规格
                                spec_combination = "默认规格:默认"
                            
                            # 创建缓存键
                            cache_key = f"{self.last_product_id}_{spec_combination}"
                            
                            # 检查是否已存在相同规格的记录
                            if cache_key not in self.product_cache:
                                # 为这个规格组合生成新的产品代码
                                product_code = self.generate_unique_code()  # 现在返回整数
                                
                                # 为这个规格创建独立的图片目录
                                spec_image_dir = os.path.join("product_images", str(product_code))  # 转换为字符串用于路径
                                os.makedirs(spec_image_dir, exist_ok=True)
                                
                                # 从临时目录复制图片到规格特定的目录
                                temp_dir = os.path.join("temp_images", self.last_product_id)
                                copied_images = []
                                for i in range(len(image_urls)):
                                    temp_path = os.path.join(temp_dir, f"image_{i+1}.jpg")
                                    if os.path.exists(temp_path):
                                        spec_path = os.path.join(spec_image_dir, f"image_{i+1}.jpg")
                                        import shutil
                                        shutil.copy2(temp_path, spec_path)
                                        copied_images.append(spec_path)
                                
                                self.product_cache[cache_key] = {
                                    'name': product_name,
                                    'specs': spec_combination,
                                    'price': price,
                                    'code': product_code,  # 存储为整数
                                    'images': image_urls
                                }
                                print(f"Recorded product info: {product_name}, {spec_combination}, {price}, {product_code}")
                            else:
                                print(f"Skipping duplicate specification: {spec_combination}")
                        except Exception as e:
                            print(f"Error processing cargo-inner element: {e}")
                else:
                    print("No images found in cache")
            else:
                print("No product ID found in cache")
            
        except Exception as e:
            print(f"Error recording product info: {e}")
            import traceback
            print(traceback.format_exc())

    def save_to_database(self, cache_key):
        try:
            if not cache_key or cache_key not in self.product_cache:
                print(f"No product information to save for cache key: {cache_key}")
                return False
            
            # 检查是否是纯产品ID的缓存（只包含图片信息）
            if cache_key == self.last_product_id:
                print(f"Skipping image-only cache for product ID: {cache_key}")
                return False
            
            cache_data = self.product_cache[cache_key]
            required_keys = ['name', 'specs', 'price', 'images', 'code']
            
            # 打印当前缓存数据的所有键
            print(f"Cache data keys for {cache_key}: {list(cache_data.keys())}")
            
            # 检查每个必���的键
            missing_keys = []
            for key in required_keys:
                if key not in cache_data:
                    missing_keys.append(key)
            
            if missing_keys:
                print(f"Missing required keys in cache for {cache_key}: {missing_keys}")
                return False
            
            # 获取当前URL作为链接
            current_url = self.browser.current_url
            
            # 添加购买记录
            self.add_purchase_record(
                current_url,
                f"{cache_data['name']} [{cache_data['specs']}]",
                cache_data['price'],
                "1",
                cache_data['images'],
                cache_data['code']
            )
            
            print(f"Successfully saved product {cache_data['code']} to database")
            return True
            
        except Exception as e:
            print(f"Error saving to database: {e}")
            import traceback
            print(traceback.format_exc())
            return False

    def extract_product_info(self, driver):
        try:
            # 等待价格元素加载
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "span.amount"))
            )
            
            # 获取价格
            price_element = driver.find_element(By.CSS_SELECTOR, "span.amount")
            price = price_element.text.strip()
            if not price:
                price = "0.00"
            
            # 移除货币符号并转换为浮点数
            try:
                price = float(price.replace('¥', '').replace('￥', '').strip())
            except ValueError:
                price = 0.00
            
            # 获取规格信息
            specs = []
            try:
                spec_elements = driver.find_elements(By.CSS_SELECTOR, ".sku-name")
                for spec in spec_elements:
                    spec_text = spec.text.strip()
                    if spec_text:
                        specs.append(spec_text)
            except Exception as e:
                print(f"Error getting specifications: {e}")
            
            # 获取商品名称
            try:
                name_element = driver.find_element(By.CSS_SELECTOR, ".product-title")
                name = name_element.text.strip()
            except Exception as e:
                print(f"Error getting product name: {e}")
                name = ""
            
            # 获取商品链接
            product_url = driver.current_url
            
            # 生成唯一的商品代码
            product_code = str(random.randint(1000000, 9999999))
            
            # 创建商品信息字典
            product_info = {
                "Name": name,
                "Specs": " | ".join(specs) if specs else "",
                "Cost": f"{price:.2f}",
                "Link": product_url,
                "Code": product_code,
                "LocationCode": "",
                "StockQuantity": "0",
                "AddedDate": datetime.now().strftime("%Y-%m-%d")
            }
            
            return product_info
            
        except Exception as e:
            print(f"Error extracting product info: {e}")
            return None

    def extract_product_id(self, url):
        # 从URL中提取产品ID
        try:
            return url.split('offer/')[1].split('.html')[0]
        except:
            return None

    def generate_unique_code(self):
        while True:
            new_code = random.randint(1000000, 9999999)  # 直接生成整数
            try:
                # 检查Excel文件中是否存在此Code
                wb = openpyxl.load_workbook('products.xlsx')
                ws = wb.active
                codes = [int(str(ws.cell(row=row, column=5).value)) for row in range(2, ws.max_row + 1) if ws.cell(row=row, column=5).value is not None]  # 获取现有的代码并转换为整数
                if new_code not in codes:
                    wb.close()
                    return new_code  # 返回整数
            except FileNotFoundError:
                return new_code  # 如果文件不存在，直接返回新生成的整数
            except Exception as e:
                print(f"Error checking unique code: {e}")
                wb.close()
                continue

    def cache_product_images(self):
        """缓存当前标签页中的产品图片"""
        try:
            current_url = self.browser.current_url
            if current_url in self.cached_urls:  # 检查URL是否已经缓存过
                return
            
            product_id = self.extract_product_id(current_url)
            if not product_id:
                return

            print(f"Caching images for product {product_id} in current tab")
            image_urls = []
            
            # 方法1：尝试获取od-gallery-turn-item-wrapper中的图片
            try:
                WebDriverWait(self.browser, 10).until(
                    EC.presence_of_element_located((By.CLASS_NAME, "od-gallery-turn-item-wrapper"))
                )
                image_containers = self.browser.find_elements(By.CLASS_NAME, "od-gallery-turn-item-wrapper")
                if image_containers:
                    # 只取前6张图片
                    for container in image_containers[:6]:
                        img_element = container.find_element(By.CLASS_NAME, "od-gallery-img")
                        img_url = img_element.get_attribute("src")
                        if img_url and self.check_image_size(img_url):
                            image_urls.append(img_url)
            except Exception as e:
                print(f"Method 1 failed: {e}")

            # 方法2：如果方法1没有找到图片，尝试获取detail-gallery-turn-wrapper中的图片
            if not image_urls:
                try:
                    WebDriverWait(self.browser, 10).until(
                        EC.presence_of_element_located((By.CLASS_NAME, "detail-gallery-turn-wrapper"))
                    )
                    
                    image_containers = self.browser.find_elements(By.CLASS_NAME, "detail-gallery-turn-wrapper")
                    print(f"Found {len(image_containers)} detail-gallery-turn-wrapper elements")
                    
                    for container in image_containers[:6]:
                        try:
                            img_element = container.find_element(By.TAG_NAME, "img")
                            img_url = img_element.get_attribute("src")
                            if img_url and img_url not in image_urls and self.check_image_size(img_url):
                                image_urls.append(img_url)
                                print(f"Found image URL: {img_url}")
                        except Exception as e:
                            print(f"Error processing image container: {e}")
                            continue
                    
                except Exception as e:
                    print(f"Method 2 failed: {e}")

            if not image_urls:
                print("No valid images found in current tab")
                return

            print(f"Found {len(image_urls)} valid images")
            
            # 缓存图片信息
            self.product_cache[product_id] = {
                'images': image_urls
            }
            self.cached_urls.add(current_url)
            
            # 立即下载图片到临时目录
            temp_dir = os.path.join("temp_images", product_id)
            os.makedirs(temp_dir, exist_ok=True)
            
            for i, url in enumerate(image_urls):
                try:
                    response = requests.get(url)
                    if response.status_code == 200:
                        img_data = io.BytesIO(response.content)
                        with Image.open(img_data) as img:
                            width, height = img.size
                            if width >= 100 and height >= 100:
                                img_path = os.path.join(temp_dir, f"image_{i+1}.jpg")
                                with open(img_path, 'wb') as f:
                                    f.write(response.content)
                                print(f"Saved image {i+1} for product {product_id}")
                            else:
                                print(f"Skipped small image {i+1}")
                except Exception as e:
                    print(f"Error downloading image {url}: {e}")

        except Exception as e:
            print(f"Error caching product images: {e}")
            import traceback
            print(traceback.format_exc())

    def check_image_size(self, url):
        """检查图片尺寸是否符合要求"""
        try:
            response = requests.get(url)
            if response.status_code == 200:
                img_data = io.BytesIO(response.content)
                with Image.open(img_data) as img:
                    width, height = img.size
                    if width >= 100 and height >= 100:
                        print(f"Image size check passed: {width}x{height}")
                        return True
                    else:
                        print(f"Image too small: {width}x{height}")
                        return False
        except Exception as e:
            print(f"Error checking image size: {e}")
            return False
        return False

    def extract_1688_info(self):
        try:
            # 加明确的等待时间
            time.sleep(2)
            
            # 等待页面全加载
            WebDriverWait(self.browser, 10).until(
                EC.presence_of_element_located((By.TAG_NAME, "body"))
            )

            # 获取当前页面URL作为链接
            current_url = self.browser.current_url

            # 获取产品链接和名称
            try:
                # 在收银台页面，产品同的位置
                if "trade.1688.com/order/trade_flow.htm" in current_url:
                    # 等待产品信息加载
                    WebDriverWait(self.browser, 10).until(
                        EC.presence_of_element_located((By.CLASS_NAME, "order-item"))
                    )
                    # 获取产品名称
                    product_name = self.browser.find_element(By.CLASS_NAME, "order-item-title").text.strip()
                    # 获取产品链接
                    product_link = self.browser.find_element(By.CLASS_NAME, "order-item-title").get_attribute("href")
                    product_id = self.last_product_id  # 使用之前缓存的产品ID
                else:
                    # 原有的产品页面逻辑
                    product_link = WebDriverWait(self.browser, 10).until(
                        EC.presence_of_element_located((By.CLASS_NAME, "offer-link"))
                    )
                    product_id = self.extract_product_id(product_link.get_attribute("href"))
                    product_name = product_link.text

                print(f"Found product: {product_name}")
            except Exception as e:
                print(f"Error getting product info: {e}")
                return

            # 获取价格
            try:
                if "trade.1688.com/order/trade_flow.htm" in current_url:
                    # 在收银台页面获取价格
                    price_element = self.browser.find_element(By.CLASS_NAME, "order-item-price")
                    price = price_element.text.strip()
                else:
                    # 原有的产品页面价格获取逻辑
                    price = WebDriverWait(self.browser, 10).until(
                        EC.presence_of_element_located((By.CLASS_NAME, "amount"))
                    ).text
                print(f"Found price: {price}")
            except Exception as e:
                print(f"Error getting price: {e}")
                price = "N/A"

            # 获取规格信息
            try:
                if "trade.1688.com/order/trade_flow.htm" in current_url:
                    # 在收银台页面获取规格
                    spec_items = self.browser.find_elements(By.CLASS_NAME, "order-item-sku")
                    specs = []
                    for item in spec_items:
                        spec_text = item.text.strip()
                        if spec_text:
                            specs.append(spec_text)
                else:
                    # 原有的产品页面规格获取逻辑
                    spec_items = self.browser.find_elements(By.CLASS_NAME, "spec-item")
                    specs = []
                    for item in spec_items:
                        spec_text = item.text.strip()
                        if spec_text:
                            specs.append(spec_text)
                specs_str = "; ".join(specs)
                print(f"Found specifications: {specs_str}")
            except Exception as e:
                print(f"Error getting specifications: {e}")
                specs_str = ""

            # 检查是否已经处理过该商品
            if product_id in self.processed_products:
                print(f"Product {product_id} already processed, skipping...")
                return

            # 检查是否有缓存的图片
            if product_id in self.product_cache:
                try:
                    cache_data = self.product_cache[product_id]
                    image_urls = cache_data['images']
                    product_code = cache_data['code']  # 使用缓存中的Code
                    
                    # 在这里取产品名称和价格
                    cache_data['name'] = product_name
                    cache_data['price'] = price
                    cache_data['specs'] = specs_str  # 添加规格信息
                    cache_data['timestamp'] = datetime.now()
                    
                    # 添加购买记录，使用完整URL作为链接
                    # 将产品名称和规格组合在一起
                    full_product_name = f"{product_name} [{specs_str}]" if specs_str else product_name
                    self.add_purchase_record(current_url, full_product_name, price, "1", image_urls, product_code)
                    
                    # 标记为已处理
                    self.processed_products.add(product_id)
                    
                    print(f"Successfully recorded purchase: {full_product_name} with code {product_code}")
                except Exception as e:
                    print(f"Error processing cached images: {e}")
            else:
                print(f"No cached images found for product {product_id}")

        except Exception as e:
            print(f"Error in extract_1688_info: {e}")
            import traceback
            print(traceback.format_exc())

    def extract_taobao_info(self):
        # 淘宝特定提取逻辑
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

    def download_images(self, image_urls, product_code):
        # 创建图片保存目录
        save_dir = os.path.join("product_images", product_code)
        os.makedirs(save_dir, exist_ok=True)

        for i, url in enumerate(image_urls):
            try:
                response = requests.get(url)
                if response.status_code == 200:
                    file_path = os.path.join(save_dir, f"image_{i+1}.jpg")
                    with open(file_path, 'wb') as f:
                        f.write(response.content)
                    print(f"Saved image {i+1} for product {product_code}")
            except Exception as e:
                print(f"Error downloading image {url}: {e}")

    def add_purchase_record(self, website, product, price, quantity, image_urls, product_code):
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
        # 分离产品名称和规格
        if "[" in product and "]" in product:
            product_name = product.split("[")[0].strip()
            specs = product.split("[")[1].rstrip("]").strip()
        else:
            product_name = product
            specs = ""
        
        # 处理价格格式
        try:
            # 移除价格中的货币符号和空格
            price_str = str(price).replace("￥", "").replace("¥", "").strip()
            # 如果价格包含范围（例如：12.00-15.00），取第一个价格
            if "-" in price_str:
                price_str = price_str.split("-")[0]
            # 确保价格是数字格式
            price_float = float(price_str)
            # 格式化价格显示（用于UI显示）
            formatted_price = f"¥{price_float:.2f}"
            # 数据库中存储数字
            db_price = price_float
        except:
            formatted_price = str(price)
            db_price = 0.0
            
        # 处理数量格式
        try:
            quantity_int = int(str(quantity).strip())
        except:
            quantity_int = 0
            
        # 确保product_code是整数
        try:
            product_code = int(str(product_code).strip())
        except:
            product_code = 0
            
        record = {
            "time": timestamp,
            "website": website,
            "product": product_name,
            "specs": specs,
            "price": formatted_price,
            "product_code": product_code  # 存储为整数
        }
        self.purchases.append(record)
        
        # 更新UI
        self.root.after(0, lambda: self.tree.insert("", 0, values=(
            timestamp, website, product, formatted_price
        )))

        max_retries = 3
        retry_count = 0
        
        while retry_count < max_retries:
            try:
                # 准备新的商品数据
                image_paths = []
                save_dir = os.path.join("product_images", str(product_code))  # 转换为字符串用于路径
                os.makedirs(save_dir, exist_ok=True)
                
                # 下载并保存原始图片
                saved_images = []
                temp_files = []  # 用于跟踪临时文件
                
                for i, url in enumerate(image_urls):
                    if i >= 6:  # 最多保存6张图片
                        break
                    try:
                        response = requests.get(url)
                        if response.status_code == 200:
                            # 保存原始图片
                            img_path = os.path.join(save_dir, f"image_{i+1}.jpg")
                            with open(img_path, 'wb') as f:
                                f.write(response.content)
                            image_paths.append(img_path)
                            saved_images.append(response.content)
                    except Exception as e:
                        print(f"Error downloading image {url}: {e}")

                # 创建新的商品记录
                new_product = {
                    "Name": product_name,
                    "Specs": specs,
                    "Cost": db_price,  # 存储数字格式的价格
                    "Link": website,
                    "Code": product_code,  # 存储为整数
                    "LocationCode": "",
                    "StockQuantity": quantity_int,  # 存储数字格式的数量
                    "AddedDate": datetime.now().strftime("%Y-%m-%d")
                }

                # 读取或创建Excel��件
                try:
                    wb = openpyxl.load_workbook('products.xlsx')
                    ws = wb.active
                except FileNotFoundError:
                    wb = openpyxl.Workbook()
                    ws = wb.active
                    # 添加表头
                    headers = ["Name", "Specs", "Cost", "Link", "Code", "LocationCode", "StockQuantity", "AddedDate"]
                    for col, header in enumerate(headers, 1):
                        ws.cell(row=1, column=col, value=header)
                    # 设置数字列的格式
                    cost_col = ws.column_dimensions['C']
                    code_col = ws.column_dimensions['E']  # Code列
                    stock_col = ws.column_dimensions['G']
                    cost_col.number_format = '#,##0.00'
                    code_col.number_format = '0'  # 整数格式
                    stock_col.number_format = '0'
                except PermissionError:
                    retry_count += 1
                    if retry_count >= max_retries:
                        self.status_label.config(text="无法访问Excel文件，请确保文件未被其他程序打开")
                        messagebox.showerror("错误", "请关闭Excel文件后重试")
                        raise
                    time.sleep(1)
                    continue

                # 查找最后一行
                last_row = ws.max_row
                product_row = None

                # 查找是否已存在相同Code的商品
                for row in range(2, last_row + 1):
                    cell_value = ws.cell(row=row, column=5).value  # Code在第5列
                    try:
                        if int(str(cell_value)) == product_code:  # 确保比较时都是整数
                            product_row = row
                            break
                    except (ValueError, TypeError):
                        continue

                if not product_row:
                    product_row = last_row + 1

                # 写入商品数据
                for col, (key, value) in enumerate(new_product.items(), 1):
                    cell = ws.cell(row=product_row, column=col, value=value)
                    # 为价格、代码和数量列设置数字格式
                    if col == 3:  # Cost列
                        cell.number_format = '#,##0.00'
                    elif col == 5:  # Code列
                        cell.number_format = '0'
                    elif col == 7:  # StockQuantity列
                        cell.number_format = '0'

                # 调整行以适应图片
                ws.row_dimensions[product_row].height = 150  # 增加行高以适应多图片

                # 添加图片到Excel（缩略图）
                for i, img_data in enumerate(saved_images):
                    if i >= 6:  # 最多显示6张图片
                        break
                    try:
                        # 将图片数据转换为PIL Image
                        img = Image.open(io.BytesIO(img_data))
                        # 调整片大小用于Excel显示
                        img.thumbnail((100, 100))
                        # 保存调整后的图片到临时文件
                        temp_path = os.path.join(save_dir, f"temp_img_{product_row}_{i}.png")
                        img.save(temp_path)
                        temp_files.append(temp_path)
                        # 创建Excel图片对
                        xl_img = XLImage(temp_path)
                        # 计算图片位置（从第9列开始放置图片，因为添加了Specs列）
                        col_letter = openpyxl.utils.get_column_letter(9 + i)
                        # 添加图片到单元格
                        ws.add_image(xl_img, f"{col_letter}{product_row}")
                    except Exception as e:
                        print(f"Error adding image {i+1} to Excel: {e}")

                try:
                    # 保存Excel文件
                    wb.save('products.xlsx')
                    wb.close()  # 确保闭工作簿
                    break  # 如果成功保存，跳出重试循环
                except PermissionError:
                    retry_count += 1
                    if retry_count >= max_retries:
                        self.status_label.config(text="无法保存Excel文件，请确保文件未被其他程序打开")
                        messagebox.showerror("错误", "请关闭Excel文件后重试")
                        raise
                    time.sleep(1)  # 等待1秒后重试
                    continue
                finally:
                    # 清理临时文件
                    for temp_file in temp_files:
                        try:
                            if os.path.exists(temp_file):
                                os.remove(temp_file)
                        except Exception as e:
                            print(f"Error deleting temp file {temp_file}: {e}")

            except Exception as e:
                print(f"Error saving to products.xlsx: {e}")
                import traceback
                print(traceback.format_exc())
                retry_count += 1
                if retry_count >= max_retries:
                    raise
                time.sleep(1)  # 等待1秒后重试

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
