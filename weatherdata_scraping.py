#importing necessary libraries
import requests,openpyxl
from bs4 import BeautifulSoup

#creating a empty excel file
excel=openpyxl.Workbook()

#identifying the active sheet
sheet=excel.active

#setting the title for sheet
sheet.title="weather status"

#setting column headers
sheet.append(['TEMPERATURE','LOCATION','SKY_STATE','PRECIPITATION','HUMIDITY','WIND'])

# query=input("ENTER THE LOCATION : ")

#list of states used to scrape the weather status
loc=["tamilnadu","andhra","himachel",'haryana','uttar pradesh','rajasthan','kerala','patna','goa','jaipur','manipur','nagaland','karnataka','punjab']

#user agent for permitting the access for webpage
headers={"User-Agent":"Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/106.0.0.0 Safari/537.36"}

try:

 for i in loc:

    #getting the html source from the webpage
    source=requests.get("https://www.google.com/search?q="+i+"+weather",headers=headers)

    #will indicate if any error
    source.raise_for_status()

    #scrapes the entire html context
    soup=BeautifulSoup(source.text,'html.parser')


    # fetching the necessary information needed for our process from their respective individual tags in html content

    num=soup.find('span',class_="wob_t q8U8x").text
    celcius=soup.find('div',class_="vk_bk wob-unit").span.text
    temperature=num+celcius 
    precipitation=soup.find('span',id="wob_pp").text
    humidity=soup.find('span',id="wob_hm").text
    wind=soup.find('span',id="wob_ws").text
    location=soup.find('div',id="wob_loc").text
    day_time=soup.find('div',id="wob_dts").text
    sky=soup.find('span',id="wob_dc").text

    #dispalys the weather report
    print(f"Temperature : ",temperature,"      Location  :  ",location)

    print(f"Sky state : ",sky)

    print(f"Precipitation : ",precipitation,"  Humidity : ",humidity,"  Wind : ",wind)


    #appending the each individual info to a excel file 
    sheet.append([temperature,location,sky,precipitation,humidity,wind])

#indicates if any error occurs
except Exception as e:
    print("e")

#saving locally the entire weather status as a excel file
excel.save("weather report7.xlsx")