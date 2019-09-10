import os
import time
import urllib
import bs4
import requests
import re
from bs4 import BeautifulSoup
import pandas as pd
import math

data = pd.read_excel('HBP.xlsx')

def get_email(website):
    try:
        with requests.Session() as s:
            base_page = s.get(website)
            soup = BeautifulSoup(base_page.content, 'html.parser')
            email = re.findall(r"[a-z0-9\.\-+_]+@[a-z0-9\.\-+_]+\.[a-z]+",soup.text, re.I)

    except Exception:
        email = 'N/A'
    time.sleep(2)
    return email

for i in range(data.shape[0]):
    c = data['subcat description'][i]
    n = math.ceil(data['Taiwan Trade Supplier #'][i]/20)
    url = data['TW Trade URL'][i]
    links_with_text = []
    vendor_name=[]
    vendor_url=[]
    vendor_phone=[]

    for i in range(n):
        link = str(url+'&page='+str(i+1))
        with requests.Session() as s:

                base_page = s.get(link)
                soup = BeautifulSoup(base_page.content, 'html.parser')
                title = soup.find_all("h3")

                for h3 in title:
                    for link in h3.find_all('a'):
                        links_with_text.append(link.get('href'))

                        time.sleep(2)
    if len(links_with_text)==0:
        for i in range(n):
            link = str(url.replace('search/','search?word=').replace('%20','+').replace('-',str('&type=product&page='+str(i+1)+'&style=')).replace('.html',''))
            with requests.Session() as s:

                base_page = s.get(link)
                soup = BeautifulSoup(base_page.content, 'html.parser')
                title = soup.find_all("h3")

                for h3 in title:
                    for link in h3.find_all('a'):
                        links_with_text.append(link.get('href'))

                        time.sleep(2)
    new_links_with_text = []
    for i in links_with_text:
        if i.startswith('https:'):
            new_links_with_text.append(i)
        else:
            new_links_with_text.append(str('https://www.taiwantrade.com/'+i))

    for i in new_links_with_text:
        with requests.Session() as s:

            url_new=str(i+'/about-us')

            base_page = s.get(url_new)
            soup = BeautifulSoup(base_page.content, 'html.parser')
            txt = soup.find_all("span")
            text = soup.find_all("dd")
            header = soup.find('h3')
            try:
                vendor_url.append((str(txt).split(' <span itemprop="url">')[1]).split('</span')[0])
            except Exception:
                try:
                    base_page = s.get(i)
                    soup = BeautifulSoup(base_page.content, 'html.parser')
                    txt = soup.find_all("span")
                    text = soup.find_all("a", attrs={'class':'link'})
                    vendor_url.append(str(text).split('>')[1].split('<')[0])
                except Exception:
                    vendor_url.append('N/A')

            try:
                base_page = s.get(url_new)
                soup = BeautifulSoup(base_page.content, 'html.parser')
                txt = soup.find_all("span")
                text = soup.find_all("dd")
                header = soup.find('h3')
                vendor_phone.append((str(text).split('"telephone">')[1]).split('<')[0])
            except Exception:
                try:
                    txt = soup.find_all("span")
                    text = soup.find_all("dd")
                    header = soup.find('h3')
                    vendor_phone.append((str(text).split('"telephone">')[1]).split('<')[0])
                except Exception:
                    vendor_phone.append('N/A')

            try:
                base_page = s.get(url_new)
                soup = BeautifulSoup(base_page.content, 'html.parser')
                txt = soup.find_all("span")
                text = soup.find_all("dd")
                header = soup.find('h3')
                vendor_name.append(str(header).split('>')[1].split('<')[0])
            except Exception:
                try:
                    base_page = s.get(i)
                    soup = BeautifulSoup(base_page.content, 'html.parser')
                    txt = soup.find_all("span")
                    text = soup.find_all("a", attrs={'class':'link'})
                    vendor_name.append(str(txt).split(' <span itemprop="name">')[2].split('</span')[0])

                except Exception:
                    try:
                        base_page = s.get(url_new)
                        soup = BeautifulSoup(base_page.content, 'html.parser')
                        txt = soup.find_all("span")
                        text = soup.find_all("dd")
                        header = soup.find('h3')
                        vendor_name.append(str(header).split('>')[1].split('<')[0])
                    except Exception:
                        vendor_name.append('N/A')


    time.sleep(2)
    try:
        new_data = pd.DataFrame(
            {'Company_name': vendor_name,
             'Website': vendor_url,
             'Phone':vendor_phone
            })
        new_data['subcat description'] = c
        new_data['email'] = new_data['Website'].apply(lambda x: get_email(x))
        new_data.to_excel(str('Taitra_'+c.replace('/','-')+'.xlsx'),index=False)
    except Exception  as e:
        print(e)
