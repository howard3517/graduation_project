from time import sleep
#from selenium.webdriver.common.keys import Keys 
from selenium.webdriver.common.by import By
from selenium import webdriver
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
import requests 
from bs4 import BeautifulSoup as bs
import os
import pandas as pd
import openpyxl
from openpyxl import load_workbook


header = {'User-Agen':'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/16.1 Safari/605.1.15'}
url = 'https://www.global-shiseido.com.tw'

def crawler(url,dir,producttype,filepath,writer):
    try:
        print(f'{producttype} clawing ... ')
        try:
            req = requests.get(url+ dir +'?srule=most-popular&sz=156',headers=header)
            data = bs(req.text,'html.parser')
        except:
            try:
                req = requests.get(url+ dir +'?srule=most-popular&sz=36',headers=header)
                data = bs(req.text,'html.parser')
            except:
                req = requests.get(url+ dir +'?srule=most-popular&sz=24',headers=header)
                data = bs(req.text,'html.parser')
        
        productname = []
        price = []
        pic = []
        brand = []
        
        # get product name
        for product in data.select('h5.product-name'):
            productname.append(''.join(product.text.split('\n')))
            
        # get product price
        for product in data.find_all('h2'):
            temp_text = ''.join(product.text.split('\n'))
            temp_text = ''.join(temp_text.split(','))
            price.append(''.join(temp_text.split('NT$')))
        
        # get product pic
        for product in data.select('img.thumb-image'):
            pic.append(product.get('src'))
            
        #get product brand
        for product in data.select('h4.product-brand'):
            brand.append(''.join(product.string.split('\n')))
            
        # make table
        df = pd.DataFrame()
        df['name'] = productname
        df['price'] = price
        df['brand'] = brand
        df['pic'] = pic
        
        
        # save data to excel
        df.to_excel(writer, sheet_name = producttype ,index=False)

        
        
        #check amount of product
        crawlamount= df.shape[0]
        temp = data.select('#results-hits-top > span')
        temp = str(temp)
        temp = ''.join(temp.split('<span>'))
        temp = ''.join(temp.split('</span>'))
        temp = ''.join(temp.split('['))
        temp = ''.join(temp.split(']'))
        shouldamount = int(temp)
        
        if crawlamount == shouldamount:
            print(f'{producttype} : Done \nTotal amount : {shouldamount}\n')
        else:
            print(f'{producttype} : got wrong \nactual amount : {shouldamount}\ncrawl amount : {crawlamount}\n')
    except Exception as e:
        print(e)
        writer.close()
        
        
        
def get_type(url,filepath,writer):
    res = requests.get(url)
    text = bs(res.text,'html.parser')
    
    producttype = []
    href = []
    #original : navigation > div > ul > li:nth-child(4) > div > div.level-2-wrapper > div.level-2-full-width > div > ul > li:nth-child(4) > ul > li > ul > li:nth-child(3) > a
    for name in text.select(' div > div.level-2-wrapper > div.level-2-full-width > div > ul > li > ul > li > ul > li > a'):
        temp_text = ' '.join(name.text.split('/')) # excel sheet name can't include /
        producttype.append(''.join(temp_text.split('\n'))) # 去除品名內的換行符號
        href.append(name.get('href'))
    
    df = pd.DataFrame()
    df['type'] = producttype
    df['href'] = href
    
    df.drop(index=0,inplace=True)#第一個頁面是空的不用爬 
    #刪除後 index 從 1 開始
    

    
    df.to_excel(writer, sheet_name='Summary' ,index=False)
    
    print(f'Summary finish!\n')
    
    return df


if __name__ == '__main__':
    
    #create a new excel to save crawler's data
    workbook = openpyxl.Workbook()
    workbook.save('Shiseido_v1.xlsx')
    
    #open excel file
    filepath = 'Shiseido_v1.xlsx'
    book = load_workbook(filepath)
    writer = pd.ExcelWriter(filepath, engine = 'openpyxl')
    
    summary = get_type(url,'Shiseido_v1.xlsx',writer)

    
    for i in range(1,summary.shape[0]):
        # index 11 新艷陽．夏 架構不一樣
        if i != 11:
            crawler(url,summary['href'][i],summary['type'][i],'Shiseido_v1.xlsx',writer)
        
    #excel 要一直開著才能在同個檔案保留舊的並新增新的 sheet
    writer.close()
        
    print('Finish !!!')