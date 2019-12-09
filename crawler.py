from bs4 import BeautifulSoup
from time import sleep
import requests
import smtplib, ssl
import os, json
from sys import argv
from email.mime.text import MIMEText
import openpyxl

def get_urls():
    
    wb = openpyxl.load_workbook('hns.xlsx', data_only=True)
    sheet = wb['Sheet1']

    ret = []

    for i, row in enumerate(sheet.rows):
        if i == 0:
            continue
        if row[3].value in ("Live", "setting"):
            ret.append(row[18].value)

    return ret

def login_info():
    with open('login.json') as key_file:
        js = json.load(key_file)
        return js['ID'], js['password'], js ['to']

def write_mail(soldout_list):

    contents = ''

    for soldout in soldout_list:
        contents += soldout + '\n'

    msg = MIMEText(contents, _charset = 'utf-8')

    msg['Subject'] = '[skynet]구매 불가 알림'
    msg['From'] = 'skynet'
    msg['To'] = 'digitalFirst'
    msg['X-Priority'] = '2'

    return msg


def send(soldout_list):
    context = ssl.SSLContext(ssl.PROTOCOL_TLS)
    mail_server = smtplib.SMTP('smtp.gmail.com', 587)
    mail_server.ehlo()
    mail_server.starttls(context=context)
    mail_server.ehlo()
    ID, password, dest = login_info()
    mail_server.login(ID, password)

    mail = write_mail(soldout_list)
    try:
        mail_server.sendmail(ID + '@gmail.com', dest, mail.as_string())
        print(ID, "sent mail", dest)
    except:
        print("email failed")

    mail_server.quit()

def get_soldout_list(urls, sleep_time):

    ret = []

    for url in urls:
        print('access to ', url + "....")
        req = requests.get(url)
        html = req.text

        soup = BeautifulSoup(html, 'html.parser')

        if req.status_code != 200:
            print('접속 불가')

        buy = soup.find_all("div", class_ = "btn_buy")
        soldout = soup.find_all("div", class_ = "btn_soldout")

        if soldout:
            print('can\'t buy at', url)
            ret.append(url)

        sleep(sleep_time)

    if len(ret) == 0:
        print("all urls are valid.")
    return ret

if __name__ == "__main__":
    urls = get_urls()
    sleep_time = 0 if len(argv) == 1 else int(argv[1])
    while True:
        send(get_soldout_list(urls, sleep_time))
        sleep(7200)