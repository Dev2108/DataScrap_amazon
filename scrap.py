import requests
from openpyxl.workbook import Workbook
from bs4 import BeautifulSoup
speakers=[]
link="https://www.amazon.in/s?k=top+50+speakers&rh=p_72%3A1318476031&dc&page=2&qid=1612175440&rnid=1318475031&ref=sr_pg_"
for i in range(1,3):
    site=requests.get(link+str(i)).text
    s=BeautifulSoup(site,'lxml')
    main=s.findAll("div",attrs={'class':'s-include-content-margin s-border-bottom s-latency-cf-section'})

    for d in main:
 
        name=d.find('span', attrs={'class':'a-size-medium a-color-base a-text-normal' }).text
  
        if d.find('span', attrs={'class':'a-price-whole'}):
            Listed_price=d.find('span', attrs={'class':'a-price-whole'}).text
        else:
            Listed_price="NA"
        if d.find('span', attrs={'class':'a-price a-text-price'}):
            a=d.find('span', attrs={'class':'a-price a-text-price'})
            if a.find('span', attrs={'class':'a-offscreen'}):
                Actual_price=a.find('span', attrs={'class':'a-offscreen'}).text
            else:
                Actual_price="NA"
        rating=d.find('span', attrs={'class':'a-icon-alt'}).text
        speakers.append([name,Listed_price,Actual_price,rating])
                 

        import pandas as pd
        df= pd.DataFrame()
        df['speakers']=speakers

#print(df)



wb=Workbook()
    # grab the active worksheet
ws = wb.active
    # Data can be assigned directly to cells
ws['A1'] = 'Name'
ws['B1'] = 'Listing Price'

ws['C1'] = 'Actual Price'
ws['D1'] = 'Ratings'
for i in speakers:
    ws.append(i)
  # Save the file
wb.save("amazon_top_speakers.xlsx")
