from selenium import webdriver
from selenium.webdriver.common.by import By
from PIL import Image
import io,time,os,ddddocr,openpyxl,sqlite3
import xml.etree.ElementTree as xet


# wb=openpyxl.load_workbook('D:\\G947\\VScode\\Python_Test\\rr\\RDT_list.xlsx')
# ws=wb.get_sheet_by_name('rdtlist')

######################################################勿動######################################################################
options=webdriver.ChromeOptions()
prefs = {'download.prompt_for_download': False, 'download.default_directory': 'C:\\Users\\user\\Downloads',}
options.add_experimental_option("prefs", prefs)
driver = webdriver.Chrome()
driver.command_executor._commands["send_command"] = ("POST", '/session/$sessionId/chromium/send_command')
params = {'cmd': 'Page.setDownloadBehavior', 'params': {'behavior': 'allow', 'downloadPath': "C:\\Users\\user\Downloads"}}
driver.execute("send_command", params)
######################################################勿動######################################################################


driver.get('https://grm.phison.com/index.php')
driver.find_element(By.ID,'name2').send_keys('COS001_01')
driver.find_element(By.ID,'name3').send_keys('20230828@Ose')
bb=driver.find_element(By.ID,'captcha_img')
pitt=Image.open(io.BytesIO(bb.screenshot_as_png))
pitt.save('Verifiy.png')
ocr=ddddocr.DdddOcr()
with open('Verifiy.png','rb') as f:
    imgg=f.read()
res=ocr.classification(imgg)
print(res[0:4])
driver.find_element(By.NAME,'captcha').send_keys(res[0:4])
btn=driver.find_element(By.NAME,'login')
btn.click()

while True:
    wonumber=input('Enter WO:')
    driver.get('https://grm.phison.com/Q10199.php')
    driver.find_element(By.NAME,'orderID').click()
    driver.find_element(By.NAME,'orderID').send_keys(wonumber)
    driver.find_element(By.XPATH,'//*[@id="form"]/div/div[2]/div[3]/div[5]/input[2]').click()
    time.sleep(1)
    driver.find_element(By.XPATH,'//*[@id="form"]/div/div[2]/div[3]/div[5]/label').click()
    time.sleep(30)

# for cc in range(ws.max_row):
#     try:
#         # wonumber=input('Enter WO:')
#         if(ws.cell(cc+2,2).value==None):
#             wonumber=ws.cell(cc+2,1).value
#             driver.find_element(By.NAME,'orderID').clear()
#             driver.find_element(By.NAME,'orderID').send_keys(wonumber)
#             driver.find_element(By.ID,'query').click()
            
#             driver.find_element(By.XPATH,'/html/body/div[1]/table/tbody/tr/td[7]/select/option[2]').click()
#             driver.find_element(By.CLASS_NAME,'download').click()
#             version=driver.find_element(By.CLASS_NAME,'VER').text
#             time.sleep(1)
#             xmlfile=f'C:/Users/user/Downloads/RunCard_COS001_01-{wonumber}-{version}.xml'
#             tree=xet.parse(xmlfile)
#             root=tree.getroot()

#             for i in root.findall('Basic'):
#                 print(i.find('MO').text,' ',i.find('ITEM_NO').text,' ',i.find('PCBA').text,' ',i.find('total_capacity').text)
#                 ws.cell(cc+2,2).value=i.find('Cust').text
#                 ws.cell(cc+2,3).value=i.find('PCBA').text
#                 ws.cell(cc+2,4).value=i.find('ITEM_NO').text
#                 ws.cell(cc+2,5).value=i.find('total_capacity').text

#                 if(cur.execute('''CREATE TABLE if not exists rdt_table
#                 (
#                     'wo' nvarchar(50) not null PRIMARY KEY,
#                     'cust' nvarchar(50),
#                     'pcba' nvarchar(100),
#                     'item_no' nvarchar(100),
#                     'total_capacity' nvarchar(20),
# 					'date' timestamp DEFAULT(datetime('now','localtime')) not NULL
#                 )''')):
#                     cc=cur.execute(f'''select wo from rdt_table where wo="{i.find('MO').text}"''')
#                     if(len(list(cc))==0):
#                         cur.execute(f"""insert into rdt_table values('{i.find('MO').text}','{i.find('Cust').text}','{i.find('PCBA').text}','{i.find('ITEM_NO').text}','{i.find('total_capacity').text}',datetime('now','localtime'))""")
                
#             os.remove(xmlfile)
#             wb.save('D:\\G947\\VScode\\Python_Test\\rr\\RDT_list.xlsx')
#             con.commit()
#         else:
#             cc+=1

#         # wb.save('D:\\G947\\VScode\\Python_Test\\rr\\wolist.xlsx')
#         # wb.close()
#     except:
#         print()
#         cc+=1