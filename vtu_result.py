from selenium import webdriver
import re
from selenium.common.exceptions import NoAlertPresentException
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoAlertPresentException
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import WebDriverException
import time
import csv
import os
import pyautogui
import pandas as pd
import cv2 as cv
import pytesseract

f = open('R:\VTU_web_scrapper\\result\\marks.csv', "w")
f.truncate()
f.close()

option = webdriver.ChromeOptions()
option.add_argument("-incognito")
option.add_experimental_option("excludeSwitches", ['enable-automation'])
option.add_experimental_option("detach",True)

browser = webdriver.Chrome(executable_path=r'C:\\Program Files (x86)\\chromedriver.exe', options=option)
browser.set_window_size(600, 800)
def fillLoginpage(usn, subject_codes):

    try:
        browser.get("https://results.vtu.ac.in/FMEcbcs22/index.php")
    except WebDriverException:
        print("Error loading VTU result page")
        quit()

    textbox = browser.find_element(by=By.XPATH, value="/html/body/div[2]/div[1]/div[2]/div/div[2]/form/div/div[2]/div[1]/div/input")
    captchabox = browser.find_element(by=By.XPATH, value="/html/body/div[2]/div[1]/div[2]/div/div[2]/form/div/div[2]/div[2]/div[1]/input")
    button = browser.find_element(by=By.XPATH, value="//*[@id='submit']")

    # time.sleep(1)
    myScreenshot = pyautogui.screenshot(region=(60, 520, 200, 95)) #region=(horizontal pos, vertical pos, vertical ratio, horizontal ratio) 
    #                   change co-ordinates to ((60, 520, 200, 95) for 125% display in windows)
    #                   change co-ordinates to ((45, 415, 170, 80) for 100% display in windows)
    myScreenshot.save(r'R:\\VTU_web_scrapper\\captcha\\captcha_img.png') #change according to your dir.

    os.chdir('R:\VTU_web_scrapper') 
    img = cv.imread('captcha\captcha_img.png',0)
    ret,thresh = cv.threshold(img,103,150,cv.THRESH_TOZERO_INV)
    os.chdir('R:\VTU_web_scrapper\captcha')
    cv.imwrite("threshold_img.png", thresh)

    # time.sleep(1)
    img2 = cv.imread('threshold_img.png',0)
    #install tesseract from https://github.com/UB-Mannheim/tesseract/wiki choose 64-bit 
    pytesseract.pytesseract.tesseract_cmd = r'C:\\Program Files\\Tesseract-OCR\\tesseract.exe' 
    custom_config = r'--oem 3 --psm 6'
    pre_captcha = pytesseract.image_to_string(img2, config=custom_config)
    pre_captcha.replace(" ", "").strip()
    captcha = re.sub('[^A-Za-z0-9]+', '', pre_captcha)
    print("Printing solved Captcha " +captcha)

    if(len(captcha) != 6 ):
        return -1
    # time.sleep(1)

    try:
        textbox.send_keys(usn)
        captchabox.send_keys(captcha) 
        button.click()
    except:
        return -1

    try:
        obj = browser.switch_to.alert
        msg=obj.text
        obj.accept()
        print(msg)
        if(msg == "Invalid captcha code !!!"):
            return -1
        if(msg == "University Seat Number is not available or Invalid..!"):
            return 1
    except NoAlertPresentException: 

        marks_list = []
        marks_list.append(usn)
        
        try:
            sub_code = 0
            while sub_code < len(subject_codes):
                internal_marks = browser.find_element_by_xpath("//*[@id='dataPrint']//*[contains(text(),'"+subject_codes[sub_code]+"')]//following::div[2]").text
                external_marks = browser.find_element_by_xpath("//*[@id='dataPrint']//*[contains(text(),'"+subject_codes[sub_code]+"')]//following::div[3]").text
                total_marks = browser.find_element_by_xpath("//*[@id='dataPrint']//*[contains(text(),'"+subject_codes[sub_code]+"')]//following::div[4]").text
                remarks = browser.find_element_by_xpath("//*[@id='dataPrint']//*[contains(text(),'"+subject_codes[sub_code]+"')]//following::div[5]").text

                marks_list.append(internal_marks)
                marks_list.append(external_marks)
                marks_list.append(total_marks)
                marks_list.append(remarks)
                sub_code +=1
        
        except NoSuchElementException:
            print("Invalid (USN, Subject Code) combination")
            with open('R:\VTU_web_scrapper\\result\marks.csv', 'a',) as f:
                write = csv.writer(f)
                marks_list.append("NA")
                write.writerow(marks_list)
            csv_read = pd.read_csv('R:\VTU_web_scrapper\\result\marks.csv')
            csv2excel = pd.ExcelWriter('R:\VTU_web_scrapper\\result\student_marks.xlsx')
            csv_read.to_excel(csv2excel)
            csv2excel.save()
            return 1 
        
        print(marks_list)
        with open('R:\VTU_web_scrapper\\result\marks.csv', 'a') as f:
            write = csv.writer(f)
            write.writerow(marks_list)
        csv_read = pd.read_csv('R:\VTU_web_scrapper\\result\marks.csv')
        csv2excel = pd.ExcelWriter('R:\VTU_web_scrapper\\result\student_marks.xlsx')
        csv_read.to_excel(csv2excel)
        csv2excel.save()

def main():
    ite=0
    student_usn = []
    os.chdir("R:\\VTU_web_scrapper\\result")
    file = open('student_usn.csv')
    csvreader = csv.reader(file)
    for usns in csvreader:
        student_usn.append(usns[0])
    
    headerList = ['USN']
    subject_codes = []
    file = open('R:\VTU_web_scrapper\\result\codes.csv')
    csvreader = csv.reader(file)
    for row in csvreader:
        subject_codes.append(row[0])
        headerList.append('Internals')
        headerList.append('Externals')
        headerList.append('Total')
        headerList.append('Remarks')
    print(subject_codes)

    with open('R:\VTU_web_scrapper\\result\marks.csv', 'a') as file:
        dw = csv.DictWriter(file, delimiter=',', fieldnames=headerList)
        dw.writeheader()

    print("START")
    while ite < len(student_usn):
        usn = student_usn[ite]
        print(usn)
        x = fillLoginpage(usn, subject_codes)
        if(x == 1):
            #Iterate with next usn as error handling
            ite = ite+1
            continue
        elif(x == -1):
            #iterate with same usn
            continue
        #Iterate with next usn
        ite = ite+1

if __name__ == "__main__":
    main()
