
from selenium import webdriver
from selenium import *
from selenium.webdriver.common.by import By
import time
from selenium.webdriver.support.ui import Select
from xlwt import Workbook
import xlwt

count=0
sections=""

course=input("enter the course code: ")

wb= xlwt.Workbook()
ws = wb.add_sheet('A Test Sheet')
style1= xlwt.easyxf("pattern: pattern solid, fore_color green; font: color white; align: horiz center")
style2= xlwt.easyxf("pattern: pattern solid, fore_color white; font: color black; align: horiz center")


ws.write(0, 0, "section",style1)
ws.write_merge(0, 0, 1, 4, 'instructor',style1)

url = "https://ucs.uob.edu.bh/"
driver = webdriver.Chrome()
driver.get(url)
#get the url



try:
    #get the url
    year= Select(driver.find_element(By.ID,"year"))
    year.select_by_value("2023")
    #select the year
    semester= Select(driver.find_element(By.ID,"sem"))
    semester.select_by_value("1")
    #select the semester
    radio_button = driver.find_element(By.XPATH,"/html/body/div/div/div/div/div/div/form/div[3]/div[2]/input")
    radio_button.click()
    #click course code button
    code= driver.find_element(By.ID,"code")
    code.send_keys(course)
    #send code
    submit= driver.find_element(By.ID,"btnsubmit")
    submit.click()


    sections =driver.find_element(By.CLASS_NAME,"courseRowClass").text
    driver.quit()
    print(course)
    
except:
    print("unavailable course")
    
    

list = sections.split("Section:")


    
try:
    for i in list[1:]:  
        instructor=i.split("Instructor:")[1].strip().split("\n")[0]
        count+=1
        ws.write(count,0,count,style1)
        ws.write_merge(count, count, 1, 4, instructor,style2)
        wb.save("sections.xls")
except:
    print("file error")









