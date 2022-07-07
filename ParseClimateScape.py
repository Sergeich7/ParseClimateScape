"""

Парсер информации об экологических организациях

Программа парсит, в многопоточном режиме, карточки экологических организациях с сайта
https://climatescape.org/organizations/
и сохраняет их в посточно в файл res.xlsx (текст
разделенный табуляциями) в формате
Link Company About Description Employees Homepage Crunchbase LinkedIn Facebook Twitter

Использованы библиотеки:
selenium

Последние изменение: 09.06.2022

"""

from selenium import webdriver
from selenium.webdriver.common.by import By

from openpyxl import Workbook
from openpyxl.styles import Font

from concurrent.futures import ThreadPoolExecutor, wait

num_pages_4_test = 10   # парсим только первые страницы для проверки. если 0, то парсим все страници
chrome_visible = True   # прячем браузер если False
max_thread = 5          # можно максимально запускать браузеров

options = webdriver.ChromeOptions()
options.add_experimental_option("excludeSwitches", ["enable-logging"])
if not chrome_visible:
    options.add_argument("--headless")

data = []

# получаем и разбираем карточку организации
def url2list(oLink, number):
    driver = webdriver.Chrome(options=options, executable_path="chromedriver.exe")
    driver.implicitly_wait(3)
    driver.set_page_load_timeout(15)
    driver.get(oLink)
    oCompany = oAbout = oDescription = oEmployees = oHomepage = oCrunchbase = oLinkedIn = oFacebook = oTwitter = ""
    try:
        oCompany = driver.find_element(By.XPATH, value="//h1[@class='flex-grow text-xl font-semibold']").text
    except:
        pass
    try:
        oAbout = driver.find_element(By.XPATH, value="//p").text
    except:
        pass
    try:
        oDescription = driver.find_element(By.XPATH, value="//div[@class='my-6']").text
    except:
        pass
    try:
        oEmployees = driver.find_element(By.XPATH, value="//span[contains(text(), 'employees')]").text.replace(" employees", "")
    except:
        pass
    try:
        oHomepage = driver.find_element(By.XPATH, value="//span[contains(text(), 'Homepage')]//ancestor-or-self::a").get_attribute("href")
    except:
        pass
    try:
        oLinkedIn = driver.find_element(By.XPATH, value="//span[contains(text(), 'LinkedIn')]//ancestor-or-self::a").get_attribute("href")
    except:
        pass
    try:
        oFacebook = driver.find_element(By.XPATH, value="//span[contains(text(), 'Facebook')]//ancestor-or-self::a").get_attribute("href")
    except:
        pass
    try:
        oCrunchbase = driver.find_element(By.XPATH, value="//span[contains(text(), 'Crunchbase')]//ancestor-or-self::a").get_attribute("href")
    except:
        pass
    try:
        oTwitter = driver.find_element(By.XPATH, value="//span[contains(text(), '" + oLink.split("/")[-1] + "')]//ancestor-or-self::a").get_attribute("href")
    except:
        pass
    data.append((number, oLink, oCompany, oAbout, oDescription, oEmployees, oHomepage, oCrunchbase, oLinkedIn, oFacebook, oTwitter))
    driver.close()
    driver.quit()


if __name__ == "__main__":
    try:
        u = open("urls.txt", "r", encoding="utf8")
    except FileNotFoundError:
        # получаем все урлы на организации и записываем в файл urls.txt
        # chromedriver.exe - находится в папке программы
        driver = webdriver.Chrome(options=options, executable_path="chromedriver.exe")
        driver.implicitly_wait(3)
        driver.set_page_load_timeout(15)
        url = "https://climatescape.org/organizations/"
        driver.get(url)
        er_links = driver.find_elements(By.XPATH, value="//a[@class='flex flex-grow py-2 sm:py-4 sm:pl-2 sm:pr-16 hover:bg-gray-200']")
        urls = list(map(lambda x: x.get_attribute("href"), er_links))
        driver.close()
        driver.quit()
        with open("urls.txt", "w", encoding="utf8") as u:
            list(map(lambda x: u.write(x + "\n"), urls))
    else:
        # читаем урлы, которые нужно спарсить из файл urls.txt
        urls = list(map(lambda x: x.strip(), u.readlines()))
    finally:
        u.close()

    print("Всего " + str(len(urls)) + " организаций")

    print("Старт......")
    futures = []
    with ThreadPoolExecutor(max_workers=max_thread) as executor:
        i = 0
        for url in urls:
            i += 1
            if num_pages_4_test > 0 and i > num_pages_4_test:
                # если режим отладки - заканчиваем парсинг
                break
            futures.append(executor.submit(url2list, url, i))
    wait(futures)
    print("Стоп")

    print(data)

    excel_file = Workbook()
    excel_sheet = excel_file.create_sheet(title="Data", index=0)
    excel_sheet.append(("Num", "Link", "Company", "About", "Description", "Employees", "Homepage", "Crunchbase", "LinkedIn", "Facebook", "Twitter"))
    excel_sheet.row_dimensions[1].font = Font(bold=True)
    excel_sheet.column_dimensions["A"].font = Font(bold=True)
    excel_sheet.freeze_panes = "A2"

    list(map(lambda x: excel_sheet.append(x), data))

    excel_file.save(filename="res.xlsx")
