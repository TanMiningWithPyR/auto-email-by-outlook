# -*- coding: utf-8 -*-
"""
Created on Mon Dec  5 15:40:15 2016

@author: tanf
"""

import win32com.client as win32
import pandas as pd
import time 
import os
import random

os.chdir("C:\\Users\\tanalan\\Desktop\\audit_message\\auto-email-by-outlook")
ATTACHMENT1 = os.getcwd() + '\\attachments\\'
TEMPLETE_HTML = os.path.join(os.getcwd(),"templete.html")

def fill_dealer_text(text,VW_dealer_name):
    html_text = text.replace('VW_dealer_name',VW_dealer_name)
    return html_text;
    
def outlook(olook,code,text,subject,recipient,attachment1):
    mail = olook.CreateItem(win32.constants.olMailItem)
    recipient_list = recipient.split(';')
    for one_recipent in recipient_list:
        mail.Recipients.Add(one_recipent)
    mail.Subject = subject
    mail.HTMLBody = text 
    mail.Attachments.Add(attachment1)
    mail.Send()   

def read_html(html):
    text = ""
    templete = open(html,'r',encoding='utf8')
    for eachLine in templete:
      text = text + eachLine
    
    return text
      
if __name__=='__main__':
    app = 'Outlook'
    olook = win32.gencache.EnsureDispatch('%s.Application' % app)
    mail_data = pd.read_excel("mail_data.xlsx").set_index('Code')    
    code = pd.read_excel("mail_data.xlsx", 1)
    html_text = read_html(TEMPLETE_HTML)
    for each_dealer in code.Code:
        VW_dealer_name = mail_data.ix[each_dealer,'VW_dealer_name']
        each_html_text = fill_dealer_text(html_text,VW_dealer_name)
        
        each_subject = '2017年度软件升级付费通知书——TUV'
        recipient = mail_data.ix[each_dealer,'TO']  
        attachment1 = ATTACHMENT1 + mail_data.ix[each_dealer,'mail_subject'] +'.pdf'
        outlook(olook,each_dealer,each_html_text,each_subject,recipient,attachment1)    
        time.sleep(random.randint(1,10))
   # olook.Quit()