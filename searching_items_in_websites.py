from selenium import webdriver
import time
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException
import pandas as pd

# assign any keyword for searching
product_codes = []
out_of_stock = []
on_the_market = []

excel_file= pd.ExcelFile(r'C:/Users/erdin/Desktop/stok_kontrol.xlsx')
sheets= excel_file.sheet_names


for sheet in sheets:
    
    data = pd.read_excel (r'C:/Users/erdin/Desktop/stok_kontrol.xlsx', sheet_name= sheet)

    columns_length= len(data.columns)

    for column_number in range(columns_length):
        column= data[data.columns[column_number]].dropna().tolist()
        
        product_codes= product_codes+column
                

#%%

# assign the driver path
driver_path = "C:\chromedriver"

# assign your website to scrape
web = 'https://www.merterfashioncenter.com/'

# create a driver object using driver_path as a parameter
driver = webdriver.Chrome(executable_path=driver_path)

driver.get(web)

# wait for the page to download
driver.implicitly_wait(5)

#-------------------------------------------------------------------

# create WebElement for a search box

search = driver.find_element_by_name('q')

for product_code in product_codes:
    
    product_name = ""
    
    search.send_keys(product_code)
    search.send_keys(Keys.RETURN)
    
    # create a WebElement using xpath
    items = WebDriverWait(driver,10).until(EC.presence_of_all_elements_located((By.XPATH, '//div[contains(@class, "px-row product-cards")]')))

    # find name

    try:
        whole_names = items[0].find_element_by_xpath(".//span[@class='item color-gray fix-height']")
        product_name = whole_names.text     
    
        find_name = product_name.find(product_code)
        
        on_the_market.append(product_code)
        search = driver.find_element_by_name('q')
    
    except NoSuchElementException:
        out_of_stock.append(product_code)
        search = driver.find_element_by_name('q')


time.sleep(4)

# keep this line of code at the bottom
driver.quit()

#%%     To create market infos by excel

# Calling DataFrame constructor on list
df = pd.DataFrame(out_of_stock)
df_2= pd.DataFrame(on_the_market)

# Create a Pandas Excel writer using XlsxWriter as the engine.
writer = pd.ExcelWriter('C:/Users/erdin/Desktop/output.xlsx', engine='xlsxwriter')

df.to_excel(writer, sheet_name='out of stock', header=False, index=False)
df_2.to_excel(writer, sheet_name='on the market', header=False, index=False)


# Close the Pandas Excel writer and output the Excel file.
writer.save()

#%%

print("Stoklarda kalmamış ürünler:")
for i in out_of_stock:
    print(i)
print("*"*20)    
print("Satıştaki ürünler:")
for i in on_the_market:
    print(i)






















