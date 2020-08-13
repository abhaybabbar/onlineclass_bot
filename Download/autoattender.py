from selenium import webdriver
import openpyxl
import time
import datetime
import schedule
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException

#make changes here, refer to README
lecture = {'IPC': "18BTCS303&#047;18BTNS303&#047;18BTIS303 - Introduction to Processor and Chips CSE II", "COMI": "Computer Organization & Microprocessor Interfacing CSE II", "DM": "18BTCS305&#047;18BTNS305&#047;18BTIS305 - Discrete Mathematics CSE II", "SHD": "German German-2 Year-TUESDAY-09.30-11.30", "EFE": "18BTCS304&#047;18BTNS304&#047;18BTIS304 - Economics and Finance for Engineers CSE II", "DS": "18BTCS301&#047;18BTNS301&#047;18BTIS301 - Data Structures CSE II", "PLI": "18BTCS311&#047;18BTNS311&#047;18BTIS311 - Programming Laboratory – I CSE II","PLII": "18BTCS312&#047;18BTNS312&#047;18BTIS312 - Programming Laboratory II CSE II", "MNI": "18BTCS321&#047;18BTNS321&#047;18BTIS321 - Mini Project –I CSE II"}
workbook = openpyxl.load_workbook("timetable.xlsx")
sheet = workbook['Sheet1']

def timetable():
    Current_Date = datetime.datetime.now().strftime('%Y %m %d')
    currentday = datetime.datetime.strptime(Current_Date, '%Y %m %d').weekday()
    currentday += 1
    periodToday = []
    print(currentday)
    if(currentday<6):
        for i in range(2,7):
            periodToday.append(sheet.cell(row=currentday, column= i).value)
        return periodToday
    else:
        print("You dont have any Class Today.")

def selector():
    global period
    if periodtoday[period]=="None":
        print("No Class")
    else:
        link = lecture[periodtoday[period]]
        period += 1
        if period>=5:         #I have only 5 lectures, you can change it if you have less or more.
            period = 0
        return link

def tcs_login():
    link = selector()
    browser = webdriver.Firefox(executable_path=r'C:\geckodriver.exe')
    browser.get('https://g01.tcsion.com/LX/login')
    browser.find_element_by_id("Usrname").send_keys("#######@mitsoe.ac.in")       #Enter your username
    browser.find_element_by_id("Passwd").send_keys("############")      #Enter your password
    browser.find_element_by_class_name("login-btn").click()
    timeout = 40
    try:
        element_present = EC.presence_of_element_located((By.CSS_SELECTOR, "a.navigationList:nth-child(3)"))
        WebDriverWait(browser, timeout).until(element_present)
        time.sleep(10)
        browser.find_element_by_css_selector("a.navigationList:nth-child(3)").click()
        element_present = EC.presence_of_element_located((By.LINK_TEXT, link))
        WebDriverWait(browser, timeout).until(element_present)
        time.sleep(10)
        browser.find_element_by_link_text(link).click()
        time.sleep(20)
        browser.find_element_by_xpath("/html/body/div[25]/div/div[4]/div[3]/div[2]/div[2]/div[2]/div/div[13]/div/div[1]/div/div[2]/div[1]/div/div/div[6]/a/span").click()
        time.sleep(10)
        browser.switch_to.window(browser.window_handles[1])
        browser.find_element_by_css_selector("button.md-button:nth-child(2)").click()
    except TimeoutException:
        print("Timed out waiting for page to load")

def schd():
    schedule.every().day.at("09:33").do(tcs_login)            #you can change time as per your schedule
    schedule.every().day.at("10:48").do(tcs_login)
    schedule.every().day.at("12:03").do(tcs_login)
    schedule.every().day.at("14:03").do(tcs_login)
    schedule.every().day.at("15:18").do(tcs_login)

period = 0
periodtoday = timetable()


if __name__ == '__main__':
    schd()
    while True:
        schedule.run_pending()
        time.sleep(1)
