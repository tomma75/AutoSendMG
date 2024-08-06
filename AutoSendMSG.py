import datetime
import os
from PyQt5.QtCore import QObject, pyqtSignal
import pyautogui
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import time
import threading
import debugpy
import pandas as pd
from bs4 import BeautifulSoup
import psutil
import json
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from pywinauto.application import Application
from pywinauto.timings import wait_until_passes

class AutoSendMSG(QObject):
    ReturnError = pyqtSignal(str)
    returnWarning = pyqtSignal(str)

    def __init__(self, isdebug):
        super().__init__()
        self.isdebug = isdebug

    def run(self):
        try:
            Manage, Login_ID, Login_PW = self.get_data()
            df, driver = self.find_RV(Login_ID, Login_PW)
            sort_df = self.sort_data(df)
            self.send_MSG(sort_df, Manage)
            # output_dir = "./output"
            # if not os.path.exists(output_dir):
            #     os.makedirs(output_dir)

            # output_file = os.path.join(output_dir, f"추출물(test)_{datetime.datetime.now().strftime('%Y%m%d')}.xlsx")

            # with pd.ExcelWriter(output_file, engine='openpyxl', mode='w') as writer:
            #     df.to_excel(writer, index=False)
        except Exception as e:
            self.ReturnError.emit(f'Main FLOW ERR : {str(e)}')

    def send_MSG(self,df_filtered, Manage):
        for process in psutil.process_iter():
                if process.name() == "chromedriver.exe":
                    process.kill()
                    
        driver_path = ChromeDriverManager().install()
        
        options = webdriver.ChromeOptions()
        options.binary_location = r".\Chrome\Application\chrome.exe"
        options.add_argument("--disable-blink-features=AutomationControlled")
        service = Service(driver_path, log_path="chromedriver.log")
        driver = webdriver.Chrome(service=service, options=options)
        page = 1
        # 웹 페이지 열기
        url = f'https://messages.google.com/web/conversations'
        driver.get(url)
        try:
            # Phone Number 입력 및 메시지 전송
            for index, row in df_filtered.iterrows():
                # 특정 요소가 나올 때까지 대기, 최대 60초
                start_chat_button = WebDriverWait(driver, 60).until(
                    EC.presence_of_element_located((By.XPATH, "//div[contains(@class, 'fab-label') and contains(text(), '채팅 시작')]"))
                )
                time.sleep(3)
                # '채팅 시작' 버튼 클릭
                start_chat_button.click()

                phone_number = row['Phone Number']
                space_name = row['Place Name']
                # 전화번호 입력 필드 찾기 및 입력
                phone_input = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, "//input[@type='text' and @class='input']"))
                )
                phone_input.send_keys(phone_number)
                time.sleep(3)
                message_input_btn = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, f"//span[contains(@class, 'anon-contact-name') and contains(text(), '{phone_number}')]"))
                )
                # '채팅 시작' 버튼 클릭
                message_input_btn.click()
                # 메시지 입력 필드 찾기 및 메시지 입력
                time.sleep(3)
                matching_row = Manage[Manage['Space_NAME'] == space_name]
        
                if not matching_row.empty:
                    msg_contents = matching_row['MSG_contents'].values[0]
                    img_name = matching_row['IMG_name'].values[0]

                message_input = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, "//textarea[@class='input' and @placeholder='문자메시지']"))
                )
                message_input.send_keys(f"{msg_contents}")  # 원하는 메시지 내용 입력
        
                # WebDriverWait을 사용하여 버튼이 존재할 때까지 기다립니다.
                button = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, "//div[contains(@class, 'input-area')]//button[@aria-label='첨부파일 선택']"))
                )
                # JavaScript를 사용하여 버튼 클릭
                driver.execute_script("arguments[0].click();", button)
                
                time.sleep(1)  # 각 메시지 전송 후 잠시 대기
                current_path = os.getcwd()
                full_img_path = os.path.join(current_path,"Send_IMG",img_name)
                full_img_path = os.path.abspath(full_img_path)
                app = Application().connect(title_re="열기")
                dlg = app['열기']  # '열기' 창을 찾습니다.
                dlg['파일 이름(&N):Edit'].set_text(full_img_path)  # 파일 이름 입력란에 파일 경로를 입력합니다.
                time.sleep(1) 
                def click_open_button():
                    try:
                        dlg['열기(&O)'].click()
                        return True
                    except:
                        return False

                # '열기' 버튼을 확실하게 클릭하기 위한 루프
                while not click_open_button():
                    time.sleep(0.5)

                # 또는 wait_until_passes를 사용하여 특정 시간 동안 버튼을 클릭 시도
                wait_until_passes(10, 0.5, click_open_button)

                time.sleep(3) 
                pyautogui.press('enter')
                with open('./Sendlist.txt', 'a') as file:
                    file.write(phone_number + '\n')

                current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                with open('./history.txt', 'a') as file:
                    file.write(f'{current_time} - {phone_number}_{space_name}\n')

        except Exception as e:
            print(f"Error: {e}")

    def sort_data(self,df):
        df_unique = df.drop_duplicates(subset=['Phone Number'])
        df_sorted = df_unique.sort_values(by='Phone Number')
        with open('./Sendlist.txt', 'r') as file:
            sendlist_numbers = file.read().splitlines()
        
        # Sendlist.txt에 없는 전화번호만 필터링
        df_filtered = df_sorted[~df_sorted['Phone Number'].isin(sendlist_numbers)]
        return df_filtered

    def get_data(self):
        Manage = pd.read_excel('Manage.xlsx')

        with open('admin.json', 'r') as file:
            admin_data = json.load(file)
            ID = admin_data['id']
            PW = admin_data['pw']
        
        return Manage, ID, PW

    def find_RV(self,Login_ID, Login_PW):
        for process in psutil.process_iter():
                if process.name() == "chromedriver.exe":
                    process.kill()
                    
        driver_path = ChromeDriverManager().install()
        
        options = webdriver.ChromeOptions()
        options.binary_location = r".\Chrome\Application\chrome.exe"
        options.add_argument("--disable-blink-features=AutomationControlled")
        service = Service(driver_path, log_path="chromedriver.log")
        driver = webdriver.Chrome(service=service, options=options)
        page = 1
        # 웹 페이지 열기
        url = f'https://partner.spacecloud.kr/reservation?page={page}'
        driver.get(url)
        # 페이지 로딩 대기
        time.sleep(5)
        # 상품 링크 저장 리스트
        email_field = driver.find_element(By.ID, 'email')
        email_field.send_keys(Login_ID)
        
        # PW 입력
        pw_field = driver.find_element(By.ID, 'pw')
        pw_field.send_keys(Login_PW)
        
        # 로그인 버튼 클릭
        login_button = driver.find_element(By.XPATH, "//button[@type='submit' and text()='호스트 이메일로 로그인']")
        login_button.click()
        
        # 로그인 후 대기 (페이지가 로딩될 시간을 고려)
        time.sleep(3)
        page = 1
        reservation_list = []
        today = datetime.datetime.today()
        
        while True:
            url = f'https://partner.spacecloud.kr/reservation?page={page}&sort_by=reservation_start_date'
            driver.get(url)
            
            # 페이지 로딩 대기
            time.sleep(5)
            
            articles = driver.find_elements(By.CLASS_NAME, "list_box")
            if not articles:
                break
            
            for article in articles:
                try:
                    place_element = article.find_element(By.CSS_SELECTOR, "dd.place")
                    place_name = place_element.text.strip()
                    
                    date_element = article.find_element(By.CSS_SELECTOR, "dd.date > span.blind")
                    date_text = date_element.find_element(By.XPATH, "..").text.strip()
                    date_str = date_text.split(" ")[0]  # '2024.08.05' 부분만 추출
                    date_str = date_str[:10]  # '2024.09.01' 부분만 추출
                    date_obj = datetime.datetime.strptime(date_str, "%Y.%m.%d")

                    # 날짜가 오늘 이전이면 루프 종료
                    if date_obj < today:
                        driver.quit()
                        df = pd.DataFrame(reservation_list)
                        return df, driver
                    
                    user_element = article.find_element(By.CSS_SELECTOR, "dd.sub_detail > p.user > span.blind")
                    user_name = user_element.find_element(By.XPATH, "..").text.strip()
                    
                    phone_element = article.find_element(By.CSS_SELECTOR, "dd.sub_detail > p.tel > span.blind")
                    phone_number = phone_element.find_element(By.XPATH, "..").text.strip()
                    
                    reservation_list.append({
                        "Place Name": place_name,
                        "User Name": user_name,
                        "Phone Number": phone_number,
                        "Date/Time": date_text
                    })
                except Exception as e:
                    print(f"Error while processing article: {e}")
            
            page += 1

        driver.quit()
        # DataFrame 생성
        df = pd.DataFrame(reservation_list)
        return df, driver



    # output_dir = "./output"
    # if not os.path.exists(output_dir):
    #     os.makedirs(output_dir)

    # output_file = os.path.join(output_dir, f"추출물(test)_{datetime.datetime.now().strftime('%Y%m%d')}.xlsx")

    # with pd.ExcelWriter(output_file, engine='openpyxl', mode='w') as writer:
    #     df.to_excel(writer, index=False)