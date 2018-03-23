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
import argparse

parser = argparse.ArgumentParser(formatter_class=argparse.ArgumentDefaultsHelpFormatter)

# set folder
parser.add_argument('--app_dir', type=str, default='C:\\Users\\tanalan\\Desktop\\alice\\auto-email-by-outlook',
                    help='Path to the application directory.')
parser.add_argument('--year', type=str, default='2018',
                    help='In subject and attachment title')
parser.add_argument('--month', type=str,
                    help='one of 1,2,3,4,5,6,7,8,9,10,11,12')
parser.add_argument('--halfyear', type=str, default='上半年',
                    help='上半年 OR 下半年')
parser.add_argument('--test', type=bool, default=True,
                    help='Try one email before normal sending email.')


def read_html(html):
  text = ""
  f = open(html,'r',encoding='utf8')
  for eachLine in f:
    text = text + eachLine
  
  f.close()
  
  return text
    
def read_table(html):
  table_text = read_html(html)  
  table_start_index = table_text.find("<table")
  table_end_index = table_text.find("</table>") + 8
  table_text = table_text[table_start_index:table_end_index]
  
  return table_text
  
def render_html(html_text, content):
  return_text = html_text
  for key in content:
    old_text = '{{' + key + '}}'
    new_text = content[key]
    return_text = return_text.replace(old_text, new_text)

  return return_text    
 
# fill table for dealer
def fill_table_cell(table_text, dealer_series, re_string="replace_row"):
  replace_num = dealer_series.shape[0]
  new_table = table_text
  for i in range(replace_num):
    old_text = re_string + str(i+1)
    new_text = str(dealer_series[old_text])
    new_table = new_table.replace(old_text, new_text, 1)    

  return new_table 

def parpare_data(flags):
  # path definition
  html_templete_dir = os.path.join(flags.app_dir, 'html_templete')
  table_html_dir = os.path.join(html_templete_dir, 'table.html')
  email_html_dir = os.path.join(html_templete_dir, 'email.html')
  table_dir = os.path.join(flags.app_dir, 'replace_text\\monthly_replace.xlsx') 
  # test
  if flags.test:
    dealer_dir = os.path.join(flags.app_dir, 'contact\\dealer_list_test.xlsx')
  else:
    dealer_dir = os.path.join(flags.app_dir, 'contact\\dealer_list.xlsx')
  attachment_dir = os.path.join(flags.app_dir, 'attachments')

  # load data into memory
  email_html_text = read_html(email_html_dir)
  table_html_text = read_table(table_html_dir)
  df_replace = pd.read_excel(table_dir,'table_data').set_index(keys='replace_row1', drop=False)
  df_dealer = pd.read_excel(dealer_dir, 'E-mail-System').set_index(keys='Title')  
  
  # dealer list
  dealer_list = [dealer_name for dealer_name in df_replace.index] 
  # html body
  html_body = []
  # TO
  TO = []
  # CC
  CC = []
  # attachment
  attachment = []
  for each_dealer in dealer_list: 
      one_attachment = os.path.join(attachment_dir,
                                flags.year + " Porsche Monthly Verify_" + flags.halfyear + "未上传行驶证清单_" + each_dealer + ".xls")
      if not os.path.exists(one_attachment):
        raise ValueError("Attachment is not exist: " + one_attachment)
      else:
        attachment.append(one_attachment)    

      table = fill_table_cell(table_html_text, df_replace.ix[each_dealer])
      content = {'replace_name': df_dealer.ix[each_dealer,'E-mail'].split('@')[0].split('.')[0],
                 'replace_month': flags.month,
                 'replace_table': table,
                 'replace_halfyear': flags.halfyear}
      html_body.append(render_html(email_html_text, content)) 
      
      TO.append(df_dealer.ix[each_dealer,'E-mail'])
      
      all_CC = [df_dealer.ix[each_dealer,each_cc] for each_cc in ['CC 1','CC 2','CC 3','CC 4']]
      useful_CC = [each_cc_mail for each_cc_mail in all_CC if each_cc_mail != '-']
      useful_CC = ';'.join(useful_CC)
      CC.append(useful_CC)
      
  data = {'dealer_list': dealer_list,
          'html_body': html_body,
          'TO': TO,
          'CC': CC,
          'attachment': attachment
          }
          
  df = pd.DataFrame(data).set_index('dealer_list', drop=False)
  
  return df

def outlook(olook,each_dealer,html_body,subject,TO,CC,attachment):
  mail = olook.CreateItem(win32.constants.olMailItem)
  mail.To = TO
  mail.CC = CC
  mail.Subject = subject
  mail.HTMLBody = html_body
  mail.Attachments.Add(attachment)
  mail.Send()    
  print(each_dealer + " has sent!")
  
def run_main_app(flags):
  # run outlook app  
  app = 'Outlook'
  olook = win32.gencache.EnsureDispatch('%s.Application' % app)
  
  each_subject = flags.year + "-" + flags.month + " Porsche OI&VL Verify 资料提交回执" 
  df = parpare_data(flags)

  dealer_list = df['dealer_list']
  for each_dealer in dealer_list:  
    html_body = df.ix[each_dealer,'html_body']
    TO = df.ix[each_dealer,'TO']
    CC = df.ix[each_dealer,'CC']
    attachment = df.ix[each_dealer,'attachment']
    outlook(olook,each_dealer,html_body,each_subject,TO,CC,attachment)    
    time.sleep(random.randint(1,10))
    
  print("Finish!")
 # olook.Quit()  
  
if __name__=='__main__':
  FLAGS = parser.parse_args()
#  class Parameter():
#    def __init__(self):
#      self.app_dir = 'C:\\Users\\tanalan\\Desktop\\alice\\auto-email-by-outlook'
#      self.year = '2018'
#      self.month = '1'
#      self.halfyear = '下半年'
#      self.test = True
      
#  FLAGS = Parameter()
  if FLAGS.month == None:
    raise ValueError("Please input select one month!")
  os.chdir(FLAGS.app_dir)
  run_main_app(FLAGS)
  