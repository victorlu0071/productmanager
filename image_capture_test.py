from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import os
import requests
from datetime import datetime
import time

def test_image_capture():
    try:
        # 初始化浏览器
        service = Service(ChromeDriverManager().install())
        browser = webdriver.Chrome(service=service)
        
        # 首先访问登录页面
        print("Accessing login page...")
        browser.get("https://login.1688.com/member/signin.htm")
        
        # 等待用户手动登录
        print("\n请在浏览器中手动登录。")
        print("登录成功后，请在此处按回车键继续...")
        input()
        
        # 确认登录成功
        print("Verifying login status...")
        time.sleep(2)  # 给一些时间让登录状态更新
        
        # 访问产品页面
        test_url = "https://detail.1688.com/offer/831987540185.html"
        print(f"\nAccessing product page: {test_url}")
        browser.get(test_url)
        
        # 等待页面加载
        WebDriverWait(browser, 10).until(
            EC.presence_of_element_located((By.CLASS_NAME, "od-gallery-turn-item-wrapper"))
        )
        
        # 获取产品ID
        product_id = test_url.split('offer/')[1].split('.html')[0]
        print(f"Product ID: {product_id}")
        
        # 获取产品名称
        product_name = browser.find_element(By.TAG_NAME, "h1").text
        print(f"Product Name: {product_name}")
        
        # 获取图片
        image_containers = browser.find_elements(By.CLASS_NAME, "od-gallery-turn-item-wrapper")
        image_urls = []
        
        # 获取前6张图片
        for i, container in enumerate(image_containers[:6]):
            img_element = container.find_element(By.CLASS_NAME, "od-gallery-img")
            img_url = img_element.get_attribute("src")
            if img_url:
                image_urls.append(img_url)
                print(f"Found image {i+1}: {img_url}")
        
        # 下载图片
        save_dir = os.path.join("test_images", f"P{product_id}")
        os.makedirs(save_dir, exist_ok=True)
        
        for i, url in enumerate(image_urls):
            try:
                response = requests.get(url)
                if response.status_code == 200:
                    file_path = os.path.join(save_dir, f"image_{i+1}.jpg")
                    with open(file_path, 'wb') as f:
                        f.write(response.content)
                    print(f"Saved image {i+1} to {file_path}")
            except Exception as e:
                print(f"Error downloading image {i+1}: {e}")
        
        print(f"\nTest completed successfully!")
        print(f"Total images found: {len(image_urls)}")
        print(f"Images saved to: {save_dir}")
        
        # 等待用户确认
        print("\n测试完成。按回车键退出...")
        input()
        
    except Exception as e:
        print(f"Error during test: {e}")
    
    finally:
        browser.quit()

if __name__ == "__main__":
    test_image_capture()
