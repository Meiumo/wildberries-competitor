from selenium import webdriver
import time
import pandas as pd
import tkinter
import tkinter.filedialog
import timeit

start = timeit.default_timer()
root = None
file = tkinter.filedialog.askopenfilename(parent=root, title='Выберите таблицу с артикулами.')
articuls = pd.read_excel(file, usecols=[0])['артикулы'].tolist()


def get_our_prices(url, articul):
    prices_list = []
    op = webdriver.FirefoxOptions()
    op.add_argument('--headless')
    driver = webdriver.Firefox(
        
        executable_path = r"geckodriver.exe", options=op
    
    )

    driver.maximize_window()

    try:
        driver.get(url=url)
        time.sleep(3)
        if driver.find_element_by_class_name('other-offers__price-now'):

            data = driver.find_elements_by_class_name('other-offers__price-now')

            for i in data:
                prices_list.append(i.get_attribute('outerHTML').replace('&nbsp;', '').replace('<b class="other-offers__price-now">', '').replace('</b>', '').replace(' ', '').replace('&nbsp;₽', ''))

            print(f'Артикул {articul}, цена {min(prices_list)}...')
            return min(prices_list)
            
    except Exception as _ex:
        print(articul, 'нет цены конкурента...')
        return 'Нет цены'
    
    finally:
        driver.close()


writer = pd.ExcelWriter(file, mode='a', engine='openpyxl', if_sheet_exists='overlay')
send_back = {'артикулы': [], 'цена_конкурента': []}

for articul in articuls:
    url= rf'https://www.wildberries.ru/catalog/{articul}/other-sellers?size=31719185'
    send_back['артикулы'].append(articul)
    send_back['цена_конкурента'].append(get_our_prices(url, articul))

send_back = pd.DataFrame(send_back)
send_back.to_excel(writer, index=False)
writer._save()
end = timeit.default_timer()
 
print(f"Done! Time taken is {end - start}s")


    




