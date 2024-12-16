import os
from tkinter import messagebox
import openpyxl
from PIL import Image
import io

def load_products(file_path):
    products = []
    product_dict = {}

    if os.path.exists(file_path):
        try:
            wb = openpyxl.load_workbook(file_path)
            ws = wb.active
            headers = [cell.value for cell in ws[1]]
            
            for row in ws.iter_rows(min_row=2):
                product = {}
                for header, cell in zip(headers, row):
                    if header:
                        product[header] = str(cell.value) if cell.value is not None else ""
                products.append(product)
                if 'Code' in product:
                    product_dict[product['Code']] = product
        
        except FileNotFoundError:
            messagebox.showerror("Error", "The products file was not found.")
        except Exception as e:
            messagebox.showerror("Error", f"An unexpected error occurred while loading products: {e}")

    return products, product_dict

def save_products(file_path, products):
    retries = 3
    for attempt in range(retries):
        try:
            # 创建新的工作簿
            wb = openpyxl.Workbook()
            ws = wb.active
            
            # 写入表头
            headers = ["Name", "Cost", "Link", "Code", "LocationCode", "StockQuantity", "AddedDate"]
            for col, header in enumerate(headers, 1):
                ws.cell(row=1, column=col, value=header)
            
            # 写入数据
            for row, product in enumerate(products, 2):
                for col, header in enumerate(headers, 1):
                    ws.cell(row=row, column=col, value=product.get(header, ""))
                
                # 设置行高以适应图片
                ws.row_dimensions[row].height = 150
                
                # 尝试添加图片
                try:
                    product_code = product.get("Code", "")
                    if product_code:
                        img_dir = os.path.join("product_images", product_code)
                        if os.path.exists(img_dir):
                            for i, img_file in enumerate(sorted(os.listdir(img_dir))):
                                if i >= 6:  # 最多显示6张图片
                                    break
                                if img_file.endswith(('.jpg', '.png')):
                                    img_path = os.path.join(img_dir, img_file)
                                    # 调整图片大小用于Excel显示
                                    img = Image.open(img_path)
                                    img.thumbnail((100, 100))
                                    temp_path = f"temp_img_{i}.png"
                                    img.save(temp_path)
                                    # 添加到Excel
                                    xl_img = openpyxl.drawing.image.Image(temp_path)
                                    col_letter = openpyxl.utils.get_column_letter(8 + i)  # 从第8列开始
                                    ws.add_image(xl_img, f"{col_letter}{row}")
                                    os.remove(temp_path)
                except Exception as e:
                    print(f"Error adding images for product {product_code}: {e}")
            
            # 保存文件
            wb.save(file_path)
            return
            
        except PermissionError:
            if attempt < retries - 1:
                import time
                time.sleep(1)
            else:
                messagebox.showerror("Error", f"Permission denied: {file_path}. Ensure the file is not open in another program.")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save products: {e}")
            break
