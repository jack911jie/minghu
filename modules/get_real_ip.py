import os
import requests
from bs4 import BeautifulSoup
import socket
import json


def get_real_ip(url='https://ip.138.com'):
    headers={
	'User-Agent':'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:80.0) Gecko/20100101 Firefox/80.0',
}
    res=requests.get(url)
    content=res.text
    # soup=BeautifulSoup(content,'html.parser')
    # body=soup.find('body').text 
    js=res.json()
    
    return js['ip']



def get_public_ip():
    ip = socket.gethostbyname('myip.opendns.com')
    return ip

# 示例用法



if __name__=='__main__':
    ip=get_real_ip(url='https://api.ipify.org/?format=json')
    print(ip)

    # ip=get_public_ip()
    # print(ip)