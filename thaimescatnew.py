from bs4 import BeautifulSoup 
import requests 
import pandas as pd 
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.action_chains import ActionChains
#Fix
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from bs4 import BeautifulSoup
import time
import re 
import pandas as pd 
from datetime import datetime
from selenium.webdriver.chrome.options import Options
import os
import xlsxwriter
from PIL import Image, ImageDraw, ImageFont
import requests
from io import BytesIO


def remove_common_items(list1, list2):
    # Create copies to avoid modifying the original lists
    new_list1 = [item for item in list1 if item not in list2]
    new_list2 = [item for item in list2 if item not in list1]
    
    return new_list1, new_list2

filename = input("Please input filename : ")

start = int(input("หน้าเริ่มต้น : "))
end = int(input("หน้าสุดท้าย : "))
catname = input("Category Name : ")
catID = int(input("Category ID : "))
# water_mark_text = input("Watermark Text : ")
# fsize = int(input("Waterark Font Size : "))

driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))

title_lis = []
url_lis_full = []
desc_lis = []
img_lis = []

company_lis = []
prov_lis = []
brand_lis = []
version_lis = []

phone_lis = []
email_lis = []
weight_lis = []
size_lis = []
price_lis = []
color_lis = []
warranty_lis = []

for page in range(start,end+1):
    url = "https://thaismegp.com/product?currentPage={}&selectedCategory={}".format(page,catID)
    driver.get(url)
    time.sleep(2)

    item_lis = [ l for l in driver.find_element(By.CSS_SELECTOR,'div.profile-content').find_elements(By.CSS_SELECTOR,'div.product-img')]
    


    url_lis = [i.find_element(By.CSS_SELECTOR,'a').get_attribute('href') for i in item_lis]
    for item_link in url_lis: 
        print(item_link)
        url_lis_full.append(item_link)

        driver.get(item_link)
        time.sleep(2)

        title = driver.find_element(By.CSS_SELECTOR,'h3').text 
        print(title)
        title_lis.append(title)

        desc = driver.find_element(By.CSS_SELECTOR,'p.text-muted').text 
        print(desc)
        desc_lis.append(desc)

        img = driver.find_element(By.CSS_SELECTOR,'div.hero-photo').find_element(By.CSS_SELECTOR,'img').get_attribute('src')
        print('Image : ',img)
        img_lis.append(img)

        p_head = [x.text for x in driver.find_elements(By.CSS_SELECTOR,'th')]
        print(p_head)

        p_data = [x.text for x in driver.find_elements(By.CSS_SELECTOR,'td')]
        print(p_data)

        template = ['ชื่อ','แบรนด์','รุ่น','จังหวัด','ขนาด','น้ำหนัก','สี','ราคา']
        table_data = list(zip(p_head,p_data)) 

        # if len(p_head) == 2:

 
        #     print('Version : -')    
        #     version_lis.append('-')
 
        #     print('Brand : -')
        #     brand_lis.append('-')



        for t in table_data: 

            if t[0] == 'ชื่อ': 
                company = t[1]
                company_lis.append(company)
                print('Company : ',company)

            elif t[0] == "รุ่น": 
                version = t[1]
                version_lis.append(version)
                print('Version : ',version)
            elif t[0] == "แบรนด์":
                brand = t[1]
                brand_lis.append(brand)
                print('Brand : ',brand)
            
            elif t[0] == 'จังหวัด': 
                prov = t[1]
                prov_lis.append(prov)
                print("Province : ",prov)

            elif t[0] == 'ขนาด': 
                size = t[1]
                size_lis.append(size)
                print("Size : ",size)


            elif t[0] == 'น้ำหนัก': 
                weight = t[1]
                weight_lis.append(weight)
                print("Weight : ",weight)
            
            elif t[0] == 'สี': 
                color = t[1]
                color_lis.append(color)
                print("Color : ",color)            

            elif t[0] == 'ราคา': 
                price = t[1]
                price_lis.append(price)
                print("Color : ",price)            
              

        # Check if 'แบรนด์' and 'รุ่น' are in table_data
        brand_present = any(t[0] == 'แบรนด์' for t in table_data)
        version_present = any(t[0] == 'รุ่น' for t in table_data)
        size_present = any(t[0] == 'ขนาด' for t in table_data)
        weight_present = any(t[0] == 'น้ำหนัก' for t in table_data)
        color_present = any(t[0] == 'สี' for t in table_data)
        price_present = any(t[0] == 'ราคา' for t in table_data)

        if not brand_present:
            brand_lis.append('-')
            print('Brand : -')

        if not version_present:
            version_lis.append('-')
            print('Version : -')
        
        if not size_present:
            size_lis.append('-')
            print('Size : -')

        if not weight_present:
            weight_lis.append('-')
            print('Weight : -')

        if not color_present:
            color_lis.append('-')
            print('Color : -')

        if not price_present:
            price_lis.append('-')
            print('Price : -')


        contact_info = [ x.find_element(By.CSS_SELECTOR,'p').text for x in driver.find_elements(By.CSS_SELECTOR,'div.detail-item')]
        for k2 in contact_info: 
            print(k2)
        
        phone = contact_info[0]
        phone_lis.append(phone)
        print("Phone : ",phone)

        email = contact_info[1]
        email_lis.append(email)
        print("Email : ",email)

        #warranty
        try:
            # Try to locate the h5 element with the specific text
            h5_element = driver.find_element(By.XPATH, "//h5[text()='การรับประกันสินค้า']")
            
            # Try to locate the next sibling div tag
            try:
                div_element = h5_element.find_element(By.XPATH, "following-sibling::div")
                div_text = div_element.text
            except:
                div_text = '-'
                
            
            # Print the text content of the div tag or the default value
            print('Warranty : ',div_text)
            warranty_lis.append(div_text)

        except:
            # If the h5 element is not found, print the default value
            print('-')
            warranty_lis.append('-')

df = pd.DataFrame()
df['ชื่อสินค้า'] = title_lis 
df['แบรนด์'] = brand_lis 
df['รุ่น'] = version_lis 
df['ขนาด'] = size_lis 
df['น้ำหนัก'] = weight_lis 
df['สี'] = color_lis
df['ราคา'] = price_lis
df['การรับประกัน'] = warranty_lis
df['รายละเอียด'] = desc_lis 
df['รูปภาพ'] = img_lis 
df['ชื่อบริษัท'] = company_lis 
df['ชื่อจังหวัด'] = prov_lis
df['เบอร์โทร'] = phone_lis 
df['Email'] = email_lis 
df['หมวดหมู่'] = [catname]*len(title_lis)
df['URL'] = url_lis_full

# Define the output file
output_file = filename + ".xlsx"

# Create a Pandas Excel writer using XlsxWriter as the engine
writer = pd.ExcelWriter(output_file, engine='xlsxwriter')

# Write the DataFrame to the Excel file
df.to_excel(writer, sheet_name='Sheet1', index=False)

# Get the xlsxwriter workbook and worksheet objects
workbook = writer.book
worksheet = writer.sheets['Sheet1']

# Insert images into the worksheet in a new column (let's say column 'K' which is index 10)
image_column_index = len(df.columns)  # Next column after the last column of the DataFrame

# Set column width to make the images fit better
worksheet.set_column(image_column_index, image_column_index, 20)

# Create a folder to temporarily save images
if not os.path.exists('temp_images'):
    os.makedirs('temp_images')



# Insert images into the worksheet
for row_num, img_url in enumerate(df['รูปภาพ'], start=1):
    response = requests.get(img_url)
    img = Image.open(BytesIO(response.content))
    
    # Convert image mode to RGB if necessary
    if img.mode in ("RGBA", "P", "CMYK"):
        img = img.convert("RGB")
    
   
        
    img_filename = f'temp_images/temp_image_{row_num}.png'
    img.save(img_filename)
    
    # Get image dimensions
    img_width, img_height = img.size
    
    # Calculate the scaling factor to fit the image into the row height
    max_height = 60  # Maximum height of the row in pixels
    scale_factor = max_height / img_height
    
    worksheet.set_row(row_num, max_height)  # Set the row height
    worksheet.insert_image(row_num, image_column_index, img_filename, {'x_scale': scale_factor, 'y_scale': scale_factor})

# Save the workbook
writer.save()

# Clean up temporary images
# for img_filename in os.listdir('temp_images'):
#     os.remove(os.path.join('temp_images', img_filename))
#os.rmdir('temp_images')

print("Finish")