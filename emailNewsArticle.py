import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders
from bs4 import BeautifulSoup
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from openpyxl import Workbook
import time

#Web driver as Chrome
driver = webdriver.Chrome(ChromeDriverManager().install())

#Url for to search the result
#뉴질랜드 is New Zealand in Korean
url = "https://search.naver.com/search.naver?&where=news&query=뉴질랜드"
driver.get(url)

#Pulling data out of HTML and XML files
time.sleep(4)
req = driver.page_source
soup = BeautifulSoup(req, 'html.parser')

# for article in articles:
articles = soup.select('#main_pack > section.sc_new.sp_nnews._prs_nws > div > div.group_news > ul > li')

#Create excel sheet
wb = Workbook()
ws1 = wb.active
#Excel file title
ws1.title = "articles1"
#Fields of Excel file
ws1.append(["title", "url", "comp"])

#Extract headline, url, and company name
for article in articles:
    title = article.select_one('div.news_wrap.api_ani_send > div > a').text
    url = article.select_one('div.news_wrap.api_ani_send > div > a')['href']
    comp = article.select_one('div > div > div.news_info > div.info_group > a.info.press').text
    ws1.append([title, url, comp])

#Save the excel file
wb.save(filename='articles1.xlsx')
#Exit the driver
driver.quit()


#Email senders info.
me = "Sender's email address"
my_password = "password of the email address"

#Login to the email
s = smtplib.SMTP_SSL('smtp.gmail.com')
s.login(me, my_password)

#Receiver's email address
you = "Destined email address"

#Email settings
msg = MIMEMultipart('alternative')
msg['Subject'] = "News article about New Zealand"
msg['From'] = me
msg['To'] = you

#Content of the email
content = "Live full potential, be better than yesterday!"
part2 = MIMEText(content, 'plain')
msg.attach(part2)

#Attaching the file
part = MIMEBase('application', "octet-stream")
with open("articles1.xlsx", 'rb') as file:
    part.set_payload(file.read())
encoders.encode_base64(part)
part.add_header('Content-Disposition', "attachment", filename="recentArticles.xlsx")
msg.attach(part)

#Send email and quit the email server
s.sendmail(me, you, msg.as_string())
s.quit()