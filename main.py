from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
import xlsxwriter
from time import sleep

chrome_options = Options()
chrome_options.add_argument("--start-maximized")
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()),options=chrome_options)

links_list = [
'https://ieltsmatters.com/product-category/the-latest-ielts-resources/vocabulary/'
,'https://ieltsmatters.com/product-category/the-main-book/'
,'https://ieltsmatters.com/product-category/test-books/'
,'https://ieltsmatters.com/product-category/all/test-books/ielts-reading-tests/'
,'https://ieltsmatters.com/product-category/all/test-books/ielts-writing-tests/'
,'https://ieltsmatters.com/product-category/all/test-books/ielts-speaking-tests/'
,'https://ieltsmatters.com/product-category/all/test-books/ielts-listening-tests/'
,'https://ieltsmatters.com/product-category/%d9%85%d8%ac%d9%84%d8%a7%d8%aa/'
,'https://ieltsmatters.com/product-category/%d9%87%d9%85%d9%87-%d9%85%d8%ad%d8%b5%d9%88%d9%84%d8%a7%d8%aa/%d9%85%d9%86%d8%a7%d8%a8%d8%b9-ielts/speaking/'
,'https://ieltsmatters.com/product-category/%d9%87%d9%85%d9%87-%d9%85%d8%ad%d8%b5%d9%88%d9%84%d8%a7%d8%aa/%d9%85%d9%86%d8%a7%d8%a8%d8%b9-ielts/writing/'
,'https://ieltsmatters.com/product-category/%d9%87%d9%85%d9%87-%d9%85%d8%ad%d8%b5%d9%88%d9%84%d8%a7%d8%aa/%d9%85%d9%86%d8%a7%d8%a8%d8%b9-ielts/reading/'
,'https://ieltsmatters.com/product-category/%d9%87%d9%85%d9%87-%d9%85%d8%ad%d8%b5%d9%88%d9%84%d8%a7%d8%aa/%d9%85%d9%86%d8%a7%d8%a8%d8%b9-ielts/listening/',
'https://ieltsmatters.com/product-category/%d9%87%d9%85%d9%87-%d9%85%d8%ad%d8%b5%d9%88%d9%84%d8%a7%d8%aa/%d9%85%d9%86%d8%a7%d8%a8%d8%b9-ielts/grammar/'
,'https://ieltsmatters.com/product-category/the-latest-ielts-resources/%d9%be%da%a9%db%8c%d8%ac-%d9%87%d8%a7%db%8c-%d8%a2%db%8c%d9%84%d8%aa%d8%b3-%d9%85%d8%aa%d8%b1%d8%b2/'
,'https://ieltsmatters.com/product-category/video-courses/']

cards_data = list()
card_links = list()

for single_link in links_list:
    driver.get(single_link)

    sleep(3)
    try:
        next_button_link = driver.find_element(By.CSS_SELECTOR,'a.next')
    except:

        cards_link_list = driver.find_elements(By.XPATH,'//*[@id="content"]/div/div/div/div[2]/div/div[2]/div[2]/div/div[1]/div/div/div/div/div/div[1]/div/h2/a')

        for single_card_link in cards_link_list:
            card_links.append(single_card_link.get_attribute('href'))
        
        continue

    while next_button_link:

        cards_link_list = driver.find_elements(By.XPATH,'//*[@id="content"]/div/div/div/div[2]/div/div[2]/div[2]/div/div[1]/div/div/div/div/div/div[1]/div/h2/a')

        for single_card_link in cards_link_list:
            card_links.append(single_card_link.get_attribute('href'))


        try:
            next_button_link.click()
            sleep(3)
        except:
            break

        try:
            next_button_link = driver.find_element(By.CSS_SELECTOR,'a.next')
        except:
            pass


for index,single_card_link in enumerate(card_links):
    driver.get(single_card_link)

    print(index)

    sleep(3)

    card_star_count = driver.find_element(By.CSS_SELECTOR,'.rmp-rating-widget__results__votes.js-rmp-vote-count').text
    cards_data.append([driver.current_url,card_star_count])


 # Create a workbook and add a worksheet.
workbook = xlsxwriter.Workbook('work.xlsx')
worksheet = workbook.add_worksheet()

 # Add a bold format to use to highlight cells.
bold = workbook.add_format({'bold': True})


 # Write some data headers.
worksheet.write('A1', 'Count', bold)
worksheet.write('B1', 'Post Url', bold)

# Start from the first cell below the headers.
row = 1
col = 0

for count,post in (cards_data):
     worksheet.write(row, col,     post)
     worksheet.write(row, col + 1, count)
     row += 1


workbook.close()
print('success')