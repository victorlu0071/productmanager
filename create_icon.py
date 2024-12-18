from PIL import Image, ImageDraw, ImageFont
import os
import sys

def get_system_font():
    """获取系统中可用的中文字体"""
    if sys.platform.startswith('win'):  # Windows系统
        font_paths = [
            "C:\\Windows\\Fonts\\simhei.ttf",  # 黑体
            "C:\\Windows\\Fonts\\msyh.ttc",    # 微软雅黑
            "C:\\Windows\\Fonts\\simsun.ttc",  # 宋体
            "C:\\Windows\\Fonts\\simkai.ttf",  # 楷体
        ]
    else:  # Linux/Mac系统
        font_paths = [
            "/usr/share/fonts/truetype/droid/DroidSansFallbackFull.ttf",
            "/System/Library/Fonts/PingFang.ttc",
            "/System/Library/Fonts/STHeiti Light.ttc",
        ]
    
    # 尝试每个字体路径
    for font_path in font_paths:
        if os.path.exists(font_path):
            print(f"Found font: {font_path}")
            return font_path
            
    print("No suitable font found, using default font")
    return None

def create_icon():
    # 创建不同尺寸的图标
    sizes = [(16,16), (32,32), (48,48), (64,64), (128,128), (256,256)]
    images = []
    
    # 获取系统字体
    font_path = get_system_font()
    
    for size in sizes:
        print(f"Creating icon size: {size}")
        # 创建新图像，使用RGBA模式支持透明度
        image = Image.new('RGBA', size, (0,0,0,0))
        draw = ImageDraw.Draw(image)
        
        # 计算圆形的尺寸和位置（留出1像素边距）
        circle_size = min(size[0], size[1]) - 2
        circle_x = (size[0] - circle_size) // 2
        circle_y = (size[1] - circle_size) // 2
        
        # 绘制圆形背景
        draw.ellipse([circle_x, circle_y, 
                     circle_x + circle_size, circle_y + circle_size], 
                     fill=(33,150,243,255))  # Material Design蓝色
        
        # 添加文字（仅在32x32及以上尺寸）
        if size[0] >= 32:
            try:
                # 计算字体大小（圆形大小的40%）
                font_size = int(circle_size * 0.4)
                
                # 尝试加载字体
                if font_path:
                    try:
                        font = ImageFont.truetype(font_path, font_size)
                        print(f"Loaded font: {font_path} with size {font_size}")
                    except Exception as e:
                        print(f"Error loading font {font_path}: {e}")
                        font = ImageFont.load_default()
                        font_size = int(circle_size * 0.3)
                else:
                    font = ImageFont.load_default()
                    font_size = int(circle_size * 0.3)
                
                # 添加文字
                text = "库"
                # 获取文字边界框
                text_bbox = draw.textbbox((0, 0), text, font=font)
                text_width = text_bbox[2] - text_bbox[0]
                text_height = text_bbox[3] - text_bbox[1]
                
                # 计算文字位置（居中）
                text_x = (size[0] - text_width) // 2
                text_y = (size[1] - text_height) // 2
                
                # 绘制文字
                draw.text((text_x, text_y), text, fill="white", font=font)
                print(f"Added text at position ({text_x}, {text_y})")
                
            except Exception as e:
                print(f"Error adding text to {size} icon: {e}")
        
        images.append(image)
    
    # 保存为ICO文件
    try:
        images[0].save("icon.ico", format='ICO', sizes=sizes, append_images=images[1:])
        print("Icon created successfully!")
        
        # 同时保存每个尺寸的预览
        for i, img in enumerate(images):
            img.save(f"icon_preview_{sizes[i][0]}x{sizes[i][1]}.png", format='PNG')
        print("Preview images created successfully!")
        
    except Exception as e:
        print(f"Error saving icon: {e}")

if __name__ == "__main__":
    create_icon() 