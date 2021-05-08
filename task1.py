import requests
from bs4 import BeautifulSoup

from PIL import Image, ImageDraw, ImageFont
import textwrap

from email.message import EmailMessage
import pandas as pd
import imghdr
import smtplib

import os



class topics():
    """Stores details of events"""
    count = 0
    def __init__(self,name,date,time,desc):
        topics.count+=1
        self.id = topics.count
        self.name = name
        self.date = date
        self.time = time
        self.desc = desc



def sendmail(toplist):
    #credentials
    email_address = 'sample-email@gmail.com'
    email_pass = os.environ['GPWD'] #using environment variable to get password

    df = pd.read_excel('mydata.xlsx')

    Name = df['Name'].tolist()
    mailId=df['Email'].tolist()
    event = df['Event Name'].tolist()
    
    newlist = [x.name.strip() for x in toplist]
    
    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
        for i in range(len(df.index)):
            
            try:
                ind = newlist.index(event[i])
                #print(ind)
            except ValueError:
                
                pass

            #send email
            newobj = toplist[ind]
            msg = EmailMessage()
            msg['Subject'] = "TECH ROADIES"
            msg['From'] = email_address
            msg.set_content(f"Hi, {Name[i]}")

                #adding attachements
            print(f"Name : {Name[i]}\nEmail :{mailId[i]}")
            with open(f"images/{newobj.id}.jpg",'rb') as f:
                file_data = f.read() 
                file_type = imghdr.what(f.name)
                file_name = f.name

            msg['To'] = mailId[i]
            msg.add_attachment(file_data, maintype='image', subtype=file_type, filename=file_name)   

            smtp.login(email_address,email_pass)
            smtp.send_message(msg)

            print("Mail successfully sent\n")



def genimage(nameobj,date,time,desc):
    """Generates image based on the given template and data"""
    
    #object to handle file naming
    filname = nameobj.id
    name = nameobj.name
    
    image = Image.open('pic.jpg')
    draw = ImageDraw.Draw(image)

    anwrapper = textwrap.TextWrapper(width=36) 
    anword_list = anwrapper.wrap(text=name) 
    name = "\n".join(anword_list)
    
    wrapper = textwrap.TextWrapper(width=60) 
    word_list = wrapper.wrap(text=desc) 
    desc = "\n".join(word_list)
    string = f"Date: {date}\n\nTime: {time}\n\nDescription:\n\n{desc}"


    newfont = ImageFont.truetype('Roboto-Regular.ttf', size=30)
    newerFont = ImageFont.truetype('Roboto-Regular.ttf', size=18)

    draw.text((82,160),name,fill= (0,0,0), font=newfont)
    draw.text((82,260),string, fill=(0,0,0), font=newerFont)

    image.save(f"images/{filname}.jpg", resolution=100.0)


def scrape():
    """Scrapes the websites to get list of events, which is stored as a list of class-'topics' """
    url = 'http://www.ieeeuvce.in/events/evold'
    headers = {'user-agent':
    'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/84.0.4147.105 Safari/537.36'}
    resp = requests.get(url,headers=headers)
    soup = BeautifulSoup(resp.content,"html.parser")

    events = soup.find_all('div',{"class":"event"})
    print("\n----------------------------------------------------\n")
    
    topiclist = []
    for feve in events:
        lin = feve.find('a')["href"]
        flin = "http://www.ieeeuvce.in" + lin
        #print(flin)
        ereq = requests.get(flin,headers=headers)
        esoup = BeautifulSoup(ereq.content,"html.parser")
        #print(esoup.find('div',{"class":"data"}))


        #getting description
        para = esoup.find('div',{"class":"data"}).text
        desc = para.split('Description: ',1)[1].strip().strip('Event Link')
        if 'Winner' in desc:
            desc = desc.split('Winner')[0]  
        
        
        
        #getting title 
        title = esoup.find('div',{"class":"section-title"}) 
        children = title.findChildren("h2", recursive=False)
        name = children[0].text 
        
        #extracting date and time
        date, time, *_ = esoup.find_all('p',{"class":"info"})

        #formatting date and time
        date = date.text.split(':')[-1].strip()
        time = time.text.split(':')[-1].strip() 
    
        evestr = f"Name:{name}\nDate: {date}\nTime: {time}\nDescription: {desc}"
        print(evestr)
        
        
        nameobj = topics(name,date,time,desc)
        topiclist.append(nameobj)
        print(nameobj.id)
        
        genimage(nameobj,date,time,desc)
        print("\n--------------------------------------------\n")

    print("---------------------Done scraping--------------------")
    return topiclist





if __name__ == '__main__':
    toplist = scrape()


    
    sendmail(toplist)