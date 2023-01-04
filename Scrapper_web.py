from selenium.webdriver.common.by import By
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from openpyxl.workbook import Workbook
import smtplib
from email.message import EmailMessage

driver=webdriver.Chrome(ChromeDriverManager().install())
driver.maximize_window()
driver.get("http://www.amazon.in/")
driver.implicitly_wait(10)
driver.find_element(By.XPATH,"//input[contains(@id,'search')]").send_keys("Samsung phones")
driver.find_element(By.XPATH,"//input[@value='Go']").click()
driver.find_element(By.XPATH,"//span[text()='Samsung']").click()
phonenames=driver.find_elements(By.XPATH,"//span[contains(@class,'a-color-base a-text-normal')]")
prices=driver.find_elements(By.XPATH,"//span[contains(@class,'price-whole')]")

myphone=[]
myprice=[]


for phone in phonenames:
    # print(phone.text)
    myphone.append(phone.text)

print("*"*50)

for price in prices:
    # print(price.text)
    myprice.append(price.text)

finallist=zip(myphone,myprice)

# for data in list(finallist):
#     print(data)

# Data scraped and stored in finallist

wb=Workbook()
wb['Sheet'].title='Samsung Data'
sh1=wb.active

sh1.append(['Name','Price'])

for x in list(finallist):
    sh1.append(x)

wb.save("FinalRec.xlsx")

# Data saved to .xlsx File

msg=EmailMessage()
msg['Subject']='Samsung phone data'
msg['From']='Sujan kumar'
msg['To']='chask96969@gmail.com'
with open('MSGtosend.txt') as myfile:
    data=myfile.read()
    msg.set_content=(data)

with open('FinalRec.xlsx','rb') as f:
    file_data=f.read()
    print("File data in binary",file_data)
    file_name=f.name
    print("File name is",file_name)
    msg.add_attachment(file_data,maintype="application",subtype="xlsx",filename=file_name)

with smtplib.SMTP_SSL('smtp.gmail.com',465) as server:
    server.login("160220a009@gmail.com","NewDemo@1234")
    server.send_message(msg)

print("Email sent")