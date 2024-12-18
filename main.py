import tkinter as tk
from tkinter import ttk
from product_management_ui import ProductManagerUI
from stock_management_ui import StockManagerUI
from file_operations import load_products, save_products

class ProductManager:
    def __init__(self):
        self.file_path = "products.xlsx"
        self.products, self.product_dict = load_products(self.file_path)
        self.root = None
        self.setup_styles()  # 设置样式
        self.init_gui()

    def setup_styles(self):
        """设置UI样式"""
        # 定义颜色
        self.colors = {
            'primary': '#2196F3',  # 主色调
            'secondary': '#64B5F6',  # 次要色调
            'success': '#4CAF50',  # 成功色
            'warning': '#FFC107',  # 警告色
            'error': '#F44336',  # 错误色
            'background': '#F5F5F5',  # 背景色
            'text': '#212121',  # 文本色
            'light_text': '#757575'  # 浅色文本
        }

        # 创建自定义样式
        style = ttk.Style()
        style.theme_use('clam')

        # 配置主按钮样式
        style.configure("Main.TButton",
            background=self.colors['primary'],
            foreground="white",
            padding=(20, 100),  # 大幅增加按钮高度
            font=('Microsoft YaHei UI', 24, 'bold'))  # 增加字体大小
        
        style.map("Main.TButton",
            background=[('active', self.colors['secondary'])])

        # 配置退出按钮样式
        style.configure("Exit.TButton",
            background=self.colors['error'],
            foreground="white",
            padding=(50, 30),  # 适当增加退出按钮高度
            font=('Microsoft YaHei UI', 16))  # 增加退出按钮字体大小
        
        style.map("Exit.TButton",
            background=[('active', '#E57373')])  # 浅红色

    def init_gui(self):
        self.root = tk.Tk()
        self.root.title("库存管理系统")
        
        # 设置窗口样式
        self.root.configure(bg=self.colors['background'])

        # 设置窗口大小和位置
        window_width = 1080
        window_height = 700
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        position_top = int((screen_height - window_height) / 2)
        position_left = int((screen_width - window_width) / 2)
        self.root.geometry(f"{window_width}x{window_height}+{position_left}+{position_top}")

        # 创建主框架
        main_frame = ttk.Frame(self.root)
        main_frame.pack(expand=True, fill=tk.BOTH, padx=40, pady=40)
        
        # 使主框架可以调整大小
        main_frame.grid_columnconfigure(0, weight=1)
        main_frame.grid_rowconfigure(1, weight=1)

        # 创建标题标签
        title_label = ttk.Label(main_frame, 
                              text="欢迎使用库存管理系统",
                              font=('Microsoft YaHei UI', 32, 'bold'),
                              foreground=self.colors['primary'])
        title_label.pack(pady=(0, 40))

        # 创建按钮框架
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.BOTH, expand=True, padx=10)
        
        # 配置按钮框架的行和列权重
        button_frame.grid_columnconfigure(0, weight=1)
        button_frame.grid_columnconfigure(1, weight=1)
        button_frame.grid_columnconfigure(2, weight=1)
        button_frame.grid_rowconfigure(0, weight=1)  # 添加行权重以允许垂直拉伸

        # 主功能按钮
        button1 = ttk.Button(button_frame, 
                           text="产品管理", 
                           command=self.open_product_manager,
                           style="Main.TButton")
        button1.grid(row=0, column=0, sticky="nsew", padx=5, ipady=80)  # 添加ipady来增加按钮高度

        button2 = ttk.Button(button_frame, 
                           text="库存管理", 
                           command=self.open_stock_manager,
                           style="Main.TButton")
        button2.grid(row=0, column=1, sticky="nsew", padx=5, ipady=80)  # 添加ipady来增加按钮高度

        button3 = ttk.Button(button_frame, 
                           text="采购助手", 
                           command=self.open_shopping_assistant,
                           style="Main.TButton")
        button3.grid(row=0, column=2, sticky="nsew", padx=5, ipady=80)  # 添加ipady来增加按钮高度

        # 退出按钮
        exit_button = ttk.Button(main_frame, 
                               text="退出系统", 
                               command=self.exit_program,
                               style="Exit.TButton")
        exit_button.pack(pady=40)

        self.root.mainloop()

    def return_to_main_menu(self):
        try:
            print("Returning to main menu")
            self.is_running = False
            self.pasteboard_monitoring = False
            
            # 取消所有pending的after调用
            if hasattr(self, 'after_id'):
                try:
                    self.root.after_cancel(self.after_id)
                except Exception as e:
                    print(f"Error canceling after: {e}")
            
            # 显示主窗口
            if hasattr(self, 'root') and self.root:
                try:
                    self.root.destroy()
                except Exception as e:
                    print(f"Error destroying window: {e}")
            
            # 重新创建主窗口
            ProductManager()
            
        except Exception as e:
            print(f"Error returning to main menu: {e}")

    def exit_program(self):
        print("Exiting program")
        self.is_running = False
        self.pasteboard_monitoring = False
        
        # 取消所有pending的after调用
        if hasattr(self, 'after_id'):
            try:
                self.root.after_cancel(self.after_id)
            except Exception as e:
                print(f"Error canceling after: {e}")
        
        if self.root:
            try:
                self.root.quit()
                self.root.destroy()
            except Exception as e:
                print(f"Error destroying window: {e}")

    def open_product_manager(self):
        if self.root:
            self.root.withdraw()  # 隐藏主窗口而不是销毁
        ProductManagerUI(self.products, None)

    def open_stock_manager(self):
        if self.root:
            self.root.withdraw()  # 隐藏主窗口而不是销毁
        StockManagerUI(self.products, self.product_dict, lambda: save_products(self.file_path, self.products))

    def open_shopping_assistant(self):
        if self.root:
            self.root.withdraw()  # 隐藏主窗口而不是销毁
        from shopping_assistant_ui import ShoppingAssistantUI
        ShoppingAssistantUI()

if __name__ == "__main__":
    ProductManager()
