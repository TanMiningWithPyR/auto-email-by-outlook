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

TEXT = """
<html xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns:m="http://schemas.microsoft.com/office/2004/12/omml" xmlns="http://www.w3.org/TR/REC-html40"><head><meta http-equiv=Content-Type content="text/html; charset=utf-8"><meta name=Generator content="Microsoft Word 15 (filtered medium)"><style><!--
/* Font Definitions */
@font-face
	{font-family:宋体;
	panose-1:2 1 6 0 3 1 1 1 1 1;}
@font-face
	{font-family:"Cambria Math";
	panose-1:2 4 5 3 5 4 6 3 2 4;}
@font-face
	{font-family:Calibri;
	panose-1:2 15 5 2 2 2 4 3 2 4;}
@font-face
	{font-family:"\@宋体";
	panose-1:2 1 6 0 3 1 1 1 1 1;}
@font-face
	{font-family:微软雅黑;
	panose-1:2 11 5 3 2 2 4 2 2 4;}
@font-face
	{font-family:"\@微软雅黑";
	panose-1:2 11 5 3 2 2 4 2 2 4;}
/* Style Definitions */
p.MsoNormal, li.MsoNormal, div.MsoNormal
	{margin:0cm;
	margin-bottom:.0001pt;
	font-size:11.0pt;
	font-family:"Calibri",sans-serif;}
a:link, span.MsoHyperlink
	{mso-style-priority:99;
	color:blue;
	text-decoration:underline;}
a:visited, span.MsoHyperlinkFollowed
	{mso-style-priority:99;
	color:purple;
	text-decoration:underline;}
span.EmailStyle17
	{mso-style-type:personal;
	font-family:"Arial",sans-serif;
	color:black;
	font-weight:normal;
	font-style:normal;
	text-decoration:none none;}
span.EmailStyle18
	{mso-style-type:personal;
	font-family:"Arial",sans-serif;
	color:#1F497D;
	font-weight:normal;
	font-style:normal;
	text-decoration:none none;}
span.EmailStyle19
	{mso-style-type:personal-reply;
	font-family:"Arial",sans-serif;
	color:#1F497D;
	font-weight:normal;
	font-style:normal;
	text-decoration:none none;}
.MsoChpDefault
	{mso-style-type:export-only;
	font-size:10.0pt;}
@page WordSection1
	{size:612.0pt 792.0pt;
	margin:72.0pt 90.0pt 72.0pt 90.0pt;}
div.WordSection1
	{page:WordSection1;}
--></style><!--[if gte mso 9]><xml>
<o:shapedefaults v:ext="edit" spidmax="1026" />
</xml><![endif]--><!--[if gte mso 9]><xml>
<o:shapelayout v:ext="edit">
<o:idmap v:ext="edit" data="1" />
</o:shapelayout></xml><![endif]--></head><body lang=ZH-CN link=blue vlink=purple style='text-justify-trim:punctuation'><div class=WordSection1><p class=MsoNormal style='line-height:150%'><span style='font-family:"微软雅黑",sans-serif'>尊敬的VW_dealer_name<span lang=EN-US><o:p></o:p></span></span></p><p class=MsoNormal style='line-height:150%'><span style='font-family:"微软雅黑",sans-serif'>您好！<span lang=EN-US><o:p></o:p></span></span></p><p class=MsoNormal style='line-height:150%'><span style='font-family:"微软雅黑",sans-serif'>随附为贵公司<span lang=EN-US>2016</span>年度上汽大众<span lang=EN-US>VW Audit II</span>现场审核通知书，烦请查收！<span lang=EN-US><o:p></o:p></span></span></p><p class=MsoNormal style='line-height:150%'><span style='font-family:"微软雅黑",sans-serif'>烦请确认后，在通知书的回执栏处签字盖章（公章）后回传至我公司，谢谢！<span lang=EN-US><o:p></o:p></span></span></p><p class=MsoNormal style='line-height:150%'><span style='font-family:"微软雅黑",sans-serif'>传真：<span lang=EN-US>021-61081199</span>转尹圣斐小姐收<span lang=EN-US><o:p></o:p></span></span></p><p class=MsoNormal style='line-height:150%'><span style='font-family:"微软雅黑",sans-serif'>或邮件：<span lang=EN-US><a href="mailto:Cindy.yin@tuv.com">Cindy.yin@tuv.com</a><u><span style='color:blue'><o:p></o:p></span></u></span></span></p><p class=MsoNormal style='line-height:150%'><span style='font-family:"微软雅黑",sans-serif'>如遇审核通知书不清楚或无法识别等情况，可直接电话联系我（<span lang=EN-US>021-60814754</span>）。<span lang=EN-US><o:p></o:p></span></span></p><p class=MsoNormal style='line-height:150%'><span style='font-family:"微软雅黑",sans-serif'>谢谢！<span lang=EN-US><o:p></o:p></span></span></p><p class=MsoNormal style='line-height:150%'><span style='font-family:"微软雅黑",sans-serif'>顺祝商祺！<span lang=EN-US><o:p></o:p></span></span></p><p class=MsoNormal style='line-height:150%'><span lang=EN-US style='font-family:"微软雅黑",sans-serif'><o:p>&nbsp;</o:p></span></p><p class=MsoNormal style='line-height:150%'><span style='font-family:"微软雅黑",sans-serif'>您诚挚的<span lang=EN-US><o:p></o:p></span></span></p><p class=MsoNormal style='line-height:150%'><span style='font-family:"微软雅黑",sans-serif'>莱茵技术（上海）有限公司<span lang=EN-US><o:p></o:p></span></span></p><p class=MsoNormal style='line-height:150%'><span style='font-family:"微软雅黑",sans-serif'>汽车管理体系<span lang=EN-US><o:p></o:p></span></span></p><p class=MsoNormal style='line-height:150%'><span style='font-family:"微软雅黑",sans-serif'>管理体系服务<span lang=EN-US><o:p></o:p></span></span></p><p class=MsoNormal style='line-height:150%'><span style='font-family:"微软雅黑",sans-serif'>尹圣斐<span lang=EN-US><o:p></o:p></span></span></p><p class=MsoNormal><span lang=EN-US style='font-family:"Arial",sans-serif;color:black'><o:p>&nbsp;</o:p></span></p></div></body></html>
"""
os.chdir("C:\\Users\\tanalan\\Desktop\\audit_message")
ATTACHMENT1 = os.getcwd() + '\\attachments\\'


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
    
if __name__=='__main__':
    app = 'Outlook'
    olook = win32.gencache.EnsureDispatch('%s.Application' % app)
    mail_data = pd.read_excel("mail_data.xlsx").set_index('Code')    
    code = pd.read_excel("mail_data.xlsx", 1)
    for each_dealer in code.Code:
        VW_dealer_name = mail_data.ix[each_dealer,'VW_dealer_name']
        each_html_text = fill_dealer_text(TEXT,VW_dealer_name)
        
        each_subject = '2017年度软件升级付费通知书——TUV'
        recipient = mail_data.ix[each_dealer,'TO']  
        attachment1 = ATTACHMENT1 + mail_data.ix[each_dealer,'mail_subject'] +'.pdf'
        outlook(olook,each_dealer,each_html_text,each_subject,recipient,attachment1)    
        time.sleep(random.randint(1,10))
   # olook.Quit()