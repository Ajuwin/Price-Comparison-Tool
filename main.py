# importing library
from selenium import webdriver
from openpyxl import Workbook

# managing chrome driver
options = webdriver.ChromeOptions()
options.add_experimental_option('excludeSwitches', ['enable-logging'])
driver = webdriver.Chrome(options=options)
driver.maximize_window()

# url of the product Apple iPhone 12 Pro Max 256 GB Pacific Blue from amazon, flipkart, paytmmall.
amazon_url = "https://www.amazon.in/dp/B08L5T31M6?th=1"
flipkart_url = "https://www.flipkart.com/apple-iphone-12-pro-max-pacific-blue-256-gb/p/itm3a0860c94250e?pid=MOBFWBYZ8STJXCVT&lid=LSTMOBFWBYZ8STJXCVT0OKDMO&marketplace=FLIPKART&q=iphone+12+pro+max&store=tyy%2F4io&srno=s_1_2&otracker=AS_QueryStore_OrganicAutoSuggest_1_9_na_na_na&otracker1=AS_QueryStore_OrganicAutoSuggest_1_9_na_na_na&fm=SEARCH&iid=f2c9d338-43fd-495f-92e2-b77c936cacf6.MOBFWBYZ8STJXCVT.SEARCH&ppt=sp&ppn=sp&ssid=dnarke0deo0000001627350256832&qH=5a7a12c4a730c1af"
paytmmall_url = "https://paytmmall.com/apple-iphone-12-pro-max-256-gb-pacific-blue-CMPLXMOBAPPLE-IPHONEDUMM202561B6D39AD-pdp?product_id=333133876&sid=f41365d4-9445-4024-9e7f-145c1f5c3a12&src=consumer_search&svc=-1&cid=66781&tracker=organic%7C66781%7Capple%20iphone%2012%20pro%20max%20256%20gb%20pacific%20blue%7Cgrid%7CSearch_experimentName%3Ddemographics_location%23NA_gender%23NA%7C%7C1%7Cdemographics_location%23NA_gender%23NA&get_review_id=333122484&site_id=2&child_site_id=6"

# best to buy from
best_price = 0
best_website = ''

# open workbook
work_book = Workbook()
worksheet = work_book.active

worksheet['B1'] = "Amazon Details"
worksheet['C1'] = "Flipkart Details"
worksheet['D1'] = "Paytmmall Details"
worksheet['A2'] = "Product Name"
worksheet['A3'] = "Product Price"
worksheet['A4'] = "Best to Buy From"

# to get product details from amazon
def get_amazon_details():
    global best_price, best_website, url

    driver.get(amazon_url)
    driver.implicitly_wait(2)
    amazon_name = driver.find_element_by_id("productTitle").text
    amazon_price = driver.find_element_by_id("priceblock_ourprice").text
    
    worksheet['B2'] = amazon_name
    worksheet['B3'] = amazon_price
    best_website = "Amazon"
    best_price = float(amazon_price[1:].replace(',', ''))
    url = amazon_url

# to get product details from flipkart
def get_flipkart_details():
    global best_price, best_website, url

    driver.get(flipkart_url)
    driver.implicitly_wait(2)
    flipkart_name = driver.find_element_by_class_name("B_NuCI").text
    flipkart_price = driver.find_element_by_xpath('//*[@id="container"]/div/div[3]/div[1]/div[2]/div[2]/div/div[4]/div[1]/div/div[1]').text

    worksheet['C2'] = flipkart_name
    worksheet['C3'] = flipkart_price
    if float(flipkart_price[1:].replace(',', '')) < best_price:
        best_price = float(flipkart_price.text[1:])
        best_website = "Flipkart"
        url = flipkart_url

# to get product details from paytmmall
def get_paytmmall_details():
    global best_website, best_price, url

    driver.get(paytmmall_url)
    driver.implicitly_wait(2)
    paytmmall_name = driver.find_element_by_class_name("NZJI").text
    paytmmall_price = driver.find_element_by_class_name("_1V3w").text

    worksheet['D2'] = paytmmall_name
    worksheet['D3'] = paytmmall_price
    if float(paytmmall_price.replace(',', '')) < best_price:
        best_price = float(paytmmall_price)
        best_website = "Paytmmall"
        url = amazon_url

get_amazon_details()
get_flipkart_details()
get_paytmmall_details()

# update details to excel
worksheet.merge_cells('B4:D4')
worksheet['B4'] = str(best_website) + ' Rs. ' + str(best_price) + "   URL:   " + url 

# save the workbook
file_name = 'product_details.xlsx'
work_book.save(file_name)

print(f"File saved as {file_name} \nDone!")

# close work book and chrome browser
work_book.close()
driver.quit()
