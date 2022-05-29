from unittest import result
from selenium import webdriver
from openpyxl import load_workbook
import time
username = "na_as_ttytanhson"
password = "Jw3@b@qnPjj7vUT"


driver = webdriver.Chrome()
driver.get("http://hssk.kcb.vn/")
time.sleep(5)

username_input = driver.find_element_by_id("username")
password_input = driver.find_element_by_id("password")
submit_btn = driver.find_element_by_xpath("/html/body/app-root/app-login/div/div[2]/div/form/button")

username_input.send_keys(username)
password_input.send_keys(password)
submit_btn.click()

time.sleep(5)
otp = input("Nhập OTP: ")

otp_input = driver.find_element_by_id("otp")
otp_input.send_keys(otp)
submit_btn = driver.find_element_by_xpath("/html/body/app-root/app-login/div/div[2]/div/div[3]/di/button")
submit_btn.click()
time.sleep(5)

driver.get("http://hssk.kcb.vn/#/ho-so-suc-khoe")
time.sleep(3)
advanced_search = driver.find_element_by_xpath('//*[@id="app-content"]/div[1]/app-info-search/div[3]/app-base-button[2]/button/span')
advanced_search.click()
time.sleep(1)
name_input = driver.find_element_by_xpath('//*[@id="demo"]/div/div[5]/div/input')
birth_input = driver.find_element_by_xpath('//*[@id="demo"]/div/div[17]/div/app-datepicker/div/input')
search_btn = driver.find_element_by_xpath('//*[@id="app-content"]/div[1]/app-info-search/div[3]/app-base-button[1]/button')



dest_filename = 'Thong_ke_xac_minh_doi_tuong.xlsx'
wb = load_workbook(filename = dest_filename)
ws = wb.active
for i in range(9, ws.max_row+1):
    name = ws.cell(i,4).value #Họ và tên
    birth = ws.cell(i,5).value #Ngày sinh

    name_input.clear()
    name_input.send_keys(name)
    birth_input.clear()
    birth_input.send_keys(birth)
    search_btn.click()
    print("Đang tìm kiếm: ", name)
    
    time.sleep(5)

    result_search = driver.find_element_by_xpath('//*[@id="app-content"]/div[1]/app-info-search/span[2]')

    # Nếu tìm thấy kết quả
    if result_search.text == "(1)":
        print("Đã tìm thấy kết quả")
        diachithuongtru_output = driver.find_element_by_xpath('//*[@id="app-content"]/div[1]/app-info-search/search-table/div/div/div[1]/table/tbody/tr/td[9]/div')
        phone_output = driver.find_element_by_xpath('//*[@id="app-content"]/div[1]/app-info-search/search-table/div/div/div[1]/table/tbody/tr/td[7]/div')

        print(diachithuongtru_output.text)
        print(phone_output.text)

        ws.cell(i,2).value = diachithuongtru_output.text
        ws.cell(i,3).value = phone_output.text
        wb.save(dest_filename)
        time.sleep(5)
        print("Cập nhập thành công")
    elif result_search.text == "(0)":
        print("Không tìm thấy")
        ws.cell(i,2).value = "Không tìm thấy"
        wb.save(dest_filename)
        time.sleep(5)
    else:
        print("Tìm thấy nhiều hơn 1 kết quả")
        ws.cell(i,2).value = "nhiều hơn 1 kết quả"
        wb.save(dest_filename)
        time.sleep(5)






driver.quit()
