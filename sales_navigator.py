import time
from selenium import webdriver
from selenium.webdriver.common.by import By
import json
import requests
import pandas as pd
from datetime import datetime
import openpyxl as oxl
import pyperclip

start_time = datetime.now()


data = {"Company":"","Website":"","Country":"","State":"","City":"","Industry":"","Number of employees":"","First name":"","Last name":"",\
        "Title":"","Email":"","Status":"","Linkedin Company":"","Linkedin Person":"" }


def write_data():
    # зαписываем cобранные данные в excel
    print(data)
    try:
        excel = pd.read_excel(r"E:\data.xlsx", index_col=0)
        row = len(excel.index) + 2
        print('exel row: ',row)
        wb = oxl.load_workbook(r"E:\data.xlsx")
        sheet = wb.active
        columns = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J','K', 'L', 'M', 'N']
        col = 0
        for k in data:
            sheet[columns[col] + str(row)] = str(data[k])
            col = col + 1
        wb.save(r"E:\data.xlsx")
    except PermissionError:
        print('--- Oшибка! Вы забыли закрыть excel фаил ---' )



def valid_email():
    # ищем имеил cотрудника с помощью api сервиса https://www.dropcontact.com/
    # и проверяем валидность почты сервисом https://www.zerobounce.net/
    try:
        apiKey = 'api key dropcontact.com'

        new_dict = {'first_name': '', 'last_name': '', 'company': '', }
        new_dict['first_name'] = data["First name"]
        new_dict['last_name'] = data["Last name"]
        new_dict['company'] = data["Company"]
        new_dict.update()

        j_data = {
            'data': [
            ], 'siren': True,
            'language': 'en'}

        j_data['data'].append(new_dict)

        r = requests.post(
            "https://api.dropcontact.io/batch",
            json=j_data,
            headers={
                'Content-Type': 'application/json',
                'X-Access-Token': apiKey
            }
        )

        time.sleep(20)
        first_response = r.content
        first_response = first_response.decode("utf-8")
        print('first_response: ', first_response)
        first_response = json.loads(first_response)
        request_id = first_response["request_id"]

        r2 = requests.get("https://api.dropcontact.io/batch/{}".format(request_id),
                          headers={'X-Access-Token': apiKey})


        second_response = r2.content
        second_response = second_response.decode("utf-8")

        second_response = json.loads(second_response)
        print('second_response: ', second_response)
        message = "--- Почта не найдена! ---"
        if 'email' in second_response['data'][0].keys():
            email = second_response["data"][0]['email'][0]['email']
            print('результат поиска почты : ',email )

            url = "https://api.zerobounce.net/v2/validate"
            api_key = "api key zerobounce.net"
            ip_address = " "  # ip_address can be blank

            params = {"email": email, "api_key": api_key, "ip_address": ip_address}
            response = requests.get(url, params=params)
            valid_email = json.loads(response.content)
            time.sleep(4)
            print('-- проверка на валидацию -- : ', valid_email)
            data["Email"] = valid_email['address']
            data["Status"] = valid_email['status']
            data.update()

        else:
            print(message)

        print('== результат == : ', data)
    except BaseException as e:
        print("Произошла ошибка в valid_email()! ")
        print(e)




def linkedin_company():
    # получаем ссылку на линкедин компании
    element = driver.find_element(By.XPATH,'//*[@id="ember68"]')
    element.click()
    time.sleep(3)
    menu = driver.find_element(By.XPATH,'//*[@id="ember69"]')
    menu.click()
    time.sleep(1)
    url = pyperclip.paste()
    data["Linkedin Company"]= url
    data.update()

def find_k(str):
    if 'K' in str:
        return True
    else:
        return False

def amount_to_integer(str):
    # парсинг количество сотрудников
    index2 = str.index(')')
    index = str.index('(')
    str = str[index + 1:index2]
    str = str.strip()
    str1 = find_k(str)

    if str1==True:
        return True
    else:
        return str1


def amount_of_employees(count):
    # количество сотрудников
    employees = amount_to_integer(count)

    if employees == True:
        data["Number of employees"] = '10000+'
        data.update()

    elif employees == 1:

        data["Number of employees"] = '2-10'
        data.update()
        return employees
    elif employees >= 2 and employees <=10:

        data["Number of employees"] = '2-10'
        data.update()
        return employees

    elif employees >=11 and employees <=50:

        data["Number of employees"] = '11-50'
        data.update()
        return employees

    elif employees >=51 and employees <=200:

        data["Number of employees"] = '51-200'
        data.update()
        return employees


    elif employees >=201 and employees <=500:

        data["Number of employees"] = '201-500'
        data.update()
        return employees


    elif employees >= 501 and employees <= 1000:

        data["Number of employees"] = '501-1000'
        data.update()
        return employees


    elif employees >= 1001 and employees <= 10000:

        data["Number of employees"] = '1001-10000'
        data.update()
        return employees


    elif employees >= 10000:

        data["Number of employees"] = '10000+'
        data.update()
        return employees



def get_industry():
    # индустрия компании
    time.sleep(4)
    web_elem = driver.find_element(By.CLASS_NAME,'artdeco-entity-lockup__subtitle')
    element = web_elem.find_element(By.CLASS_NAME,'t-14')
    el = element.find_element(By.TAG_NAME,'span')
    data["Industry"]= el.text
    data.update()



def get_link():
    # получаем ссылку нa caйт комрпнии
    webelement = driver.find_element(By.CLASS_NAME,'account-actions')
    url = webelement.find_element_by_tag_name('a').get_attribute('href')
    print('url компании : ',url)

    if 'www.' in url :

        index = url.find('.')

        new_url = url[index + 1:]

        if new_url.endswith('/'):
            new_url = new_url.replace('/', '')
            data["Website"]= new_url

        else:
            data["Website"] = new_url


    if not 'www.'in url and url.endswith('/'):
        url = url[:-1]
        data["Website"] = url


    if not 'www.' in url and not url.endswith('/') and not 'https://' in url and not 'http://' in url:
        data["Website"] = url


    if 'https://' in url and not 'www' in url or 'http://' in url and not 'www' in url:
        index = url.find('//')
        new_url = url[index + 2:]

        if new_url.endswith('/'):
            new_url = new_url.replace('/', '')
            data["Website"] = new_url

        else:
            data["Website"] = new_url



def parse_job_position(str):
    # должность сотрудника
    str = str.strip()
    index_first = str.rfind("at")
    title = str[:index_first]
    title = title.strip()
    data["Title"]= title
    data.update()


def parse_name(str):
    #имя и фамилия сотрудника
    str = str.strip()
    count_space = str.count(' ')
    if count_space > 1:
        index_first = str.find(" ")
        index_last = str.rfind(" ")
        name = str[:index_first]
        last_name = str[index_last:]
        last_name = last_name.strip()
        name = name.strip()
        data["First name"] = name
        data.update()
        data["Last name"] = last_name
        data.update()
        return data

    else:
        index = str.find(" ")
        name = str[:index]
        name = name.strip()
        last_name = str[index:]
        last_name = last_name.strip()
        data["First name"] = name
        data.update()
        data["Last name"] = last_name
        data.update()
        return data



def parse_location_data(string):
    #данные о местонахождении компании
    global first_coma
    commas = string.count(',')

    if commas == 1:
        index = string.find(',')
        city = string[:index]
        city = city.strip()
        data["City"] = city
        data.update()
        country = string[index + 1:]
        country = country.strip()
        data["Country"] = country
        data.update()

    if commas == 2:
        index = 0
        comma = 0
        for s in string:
            index += 1
            if s == ',':
                comma += 1
                if comma == 1:
                    first_coma = index
                    city = string[:index - 1]
                    city = city.strip()
                    data["City"]= city
                    data.update()

                if comma == 2:
                    state = string[first_coma:index - 1]
                    state = state.strip()
                    data["State"]= state
                    data.update()
                    country = string[index:]
                    country = country.strip()
                    data["Country"]= country
                    data.update()

        return data



def collect_companies():
    #coбираем все компании на странице
    while True:
              print('Поставьте галочку напртив компании и нажмите 1')
              command = input('Или 0 (ноль) чтобы закрыть скрипт: ')

              if command == '1':
                  try:
                        driver.execute_script("""
                            let employeers = document.querySelectorAll('.artdeco-list__item');
                                 console.log('employeers: ',employeers);
                                 console.log('employeers: ',employeers.length);
                                 for (let i =0; employeers.length > i; i++){
                                        let element = employeers[i].querySelector('.small-input');
                                        console.log('element: ',element);
                                        if (element.checked){
                                            a = employeers[i].querySelector('.artdeco-entity-lockup__title a').href;
                                            console.log('a: ',a);
                                            window.open(a, '_blank').focus();
                                        }
                                      }
                                      """)
                        driver.switch_to.window(driver.window_handles[1])
                        company = driver.find_element(By.CLASS_NAME, 'artdeco-entity-lockup__title')
                        data["Company"] = company.text
                        data.update()
                        location = driver.find_element(By.CLASS_NAME, 't-12').text
                        parse_location_data(location)
                        get_link()
                        get_industry()
                        data["Company"]= company.text
                        data.update()
                        linkedin_company()
                        l = driver.execute_script("""
                                            let l = document.querySelector('.link-without-visited-and-hover-state').href;
                                            console.log('l: ',l);
                                            return l;
                                        """)

                        all_employees = driver.execute_script("""
                                            let all_employees = document.querySelector('.link-without-visited-and-hover-state');
                                            return all_employees;
                                        """)
                        amount_of_employees(all_employees.text)
                        driver.get(l)
                        to_select_employees()
                  except BaseException as e:
                      print("Произошла ошибка в collect_companies(): ")
                      print(e)

              if command == '0':
                  break



def check_title():
    #удаляем лишнее сииволы из должности сотрудника
    s = data['Title']

    if '&amp;' in s:
        a = s.find('&')
        b = s.find(';')
        first_part=s[:a+1]
        second_part= s[b+1:]
        data['Title']=first_part+second_part
        data.update()



def to_select_employees():
    #выбор сотрудника
    global driver
    while True:
        print('Выберите ТОЛЬКО ОДНОГО сотрудникa в браузере,поставьте галочку')
        command = input('После выбора сотрудника нажмите 1 Или 0 (ноль) чтобы вернуться к выбору компании: ')
        if command == '1':
            try:
                d = driver.execute_script("""
                            const d = new Object();
                            let a;
                            let span;

                            let employeers = document.querySelectorAll('.artdeco-list__item');
                            console.log('employeers: ',employeers);
                            console.log('employeers: ',employeers.length);
                            for (let i =0; employeers.length > i; i++){
                                  let element = employeers[i].querySelector('.small-input');
                                  console.log('element: ',element);
                                  if (element.checked){
                                      a = employeers[i].querySelector('.artdeco-entity-lockup__title a');
                                      console.log('a: ',a);
                                      console.log('a: ',a.text);
                                      d.name = a.text
                                      let span = employeers[i].querySelectorAll('.artdeco-entity-lockup__subtitle span');
                                      console.log('span: ',span[1]);
                                      d.title=span[1].innerHTML;
                                      person(a)
                                      return d
                               }
                            }
                            function person(a) {
                              window.open(a, '_blank').focus();
                            };
                            """)

                parse_name(d["name"])
                parse_job_position(d["title"])
                check_title()
                driver.switch_to.window(driver.window_handles[-1])
                time.sleep(5)
                button = driver.execute_script("""
                     let button = document.querySelector('.right-actions-overflow-menu-trigger');
                     console.log('button:',button);
                     return button;
                 """)
                button.click()
                time.sleep(3)
                menu = driver.execute_script("""
                     let div = document.querySelector('.artdeco-dropdown__content-inner');
                     console.log('div:',div);
                     let ul = div.children[0];
                     console.log('ul:',ul);
                     let li = ul.children[3];
                     let menu = li.children[0];
                     console.log('menu: ', menu);
                     return menu  
                 """)
                menu.click()
                time.sleep(1)
                url = pyperclip.paste()
                data["Linkedin Person"] = url
                data.update()
                valid_email()
                write_data()
                driver.close()
                driver.switch_to.window(driver.window_handles[-1])

            except BaseException as e:
                print("Произошла ошибка в to_select_employees(): ")
                print(e)

        if command == '0':
            break
        if command == '3':
            driver.switch_to.window(driver.window_handles[-1])




def enter_to_linkedin():
    #вход нa cαйт sales navigator
    global driver
    opt = webdriver.ChromeOptions()
    # добавить профиль пользователя в браузер
    opt.add_argument(' --profile-directory="Default"')
    # путь к папке chromeprofile
    opt.add_argument(r"--user-data-dir=C:\\Users\\Zhenia\\Desktop\\salesNav\\chrome\\chromedriver_win32\\chromeprofile")
    opt.add_argument("--remote-debugging-port=9222")
    opt.add_argument("--no-sandbox")
    opt.add_argument("--disable-setuid-sandbox")
    opt.add_argument("disable-infobars")
    opt.add_experimental_option("excludeSwitches", ["enable-automation"])
    opt.add_experimental_option('useAutomationExtension', False)
    opt.add_argument('--disable-blink-features=AutomationControlled')
    driver = webdriver.Chrome(
        executable_path="C:\\Users\\Zhenia\\Desktop\\salesNav\\chrome\\chromedriver_win32\\chromedriver.exe",
        chrome_options=opt)
    driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
        "source": """
                    const newProto = navigator.__proto__
                    delete newProto.webdriver
                    navigator.__proto__ = newProto
                    """
    })
    driver.maximize_window()
    driver.get("https://www.linkedin.com/uas/login?session_redirect=/sales&fromSignIn=true&trk=navigator")
    time.sleep(10)
    driver.find_element(By.XPATH, '//*[@id="username"]').send_keys("email")
    time.sleep(15)
    driver.find_element(By.XPATH, '//*[@id="password"]').send_keys("password")
    time.sleep(10)
    button = driver.execute_script("""
                  let button = document.querySelector('.btn__primary--large');
                  console.log('button: ',button);
                  return button;
      """)
    button.click()
    time.sleep(1)
    driver.get("https://www.linkedin.com/sales/search/company?companySize=I%2CE%2CF%2CG%2CH&geoIncluded=101165590&industryIncluded=27&searchSessionId=3iaVlQyVQqae1aWQBQO0eA%3D%3D")
    time.sleep(5)
    collect_companies()


enter_to_linkedin()
















































