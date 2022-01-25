from __future__ import division
import gspread
import datetime
from time import sleep
import sys
import time
import os, io
import matplotlib.pyplot as plt
os.environ['GOOGLE_APPLICATION_CREDENTIALS']='/home/pi/freshexpress_jsonkey/freshex_in.json'
os.environ['GRPC_DNS_RESOLVER'] = 'native'
gc = gspread.service_account(filename='/home/pi/freshexpress_jsonkey/fresh-express-ocr-efe99845e1f0.json')
sh = gc.open_by_key('1ML03e2lmRb5VVYHWdExbTdaXyRPcqkrYZYP1Sir5QA0')

def Reverse(lst):
    return[ele for ele in reversed(lst)]
def control_no_vs_berry(site_code,start_date,end_date,bot,chat_id):
    if int(site_code) == 2:
        worksheet = sh.worksheet('Input Sheet Ideal')
        site = 'Ideal'
    elif int(site_code) == 3:
        worksheet = sh.worksheet('Input Sheet Amaya')
        site = 'Amaya'
    elif int(site_code) == 4:
        worksheet = sh.worksheet('Input Sheet Miraj')
        site = 'Miraj'
    else:
        bot.sendMessage(chat_id,'You have entered wrong Site ID')
    date_list=worksheet.col_values(2)
    date_list2=[]
    for anydate in date_list:
        date_list2.append(str(anydate.decode('utf-8')))
    date_list=date_list2
    if (start_date) in date_list:
        for i in range(len(date_list)):
            if str(date_list[i])==str(start_date):
                break
        
        date_list=Reverse(date_list)
        if (end_date) in date_list:
            for j in range(len(date_list)):
                if str(date_list[j])==str(end_date):
                    break
        else:
            bot.sendMessage(chat_id,'You have entered wrong End Date')
        j = len(date_list)-j
        control_number_list=worksheet.col_values(3)
        berry_percen_list=worksheet.col_values(6)
        control_number_list = control_number_list[i:j]
        berry_percen_list = berry_percen_list[i:j]
        new_control_number_list = []
        new_berry_percen_list = []
        s=1
        for val in control_number_list: 
            if len(val)<12:
                new_control_number_list.append(int(val[-4:].decode('utf-8')))
            else:
                val.replace(" ","")
                new_control_number_list.append(float(str(val[-7:-3].decode('utf-8'))+'.'+str(s)))
                s=s+1
        #print(new_control_number_list)
        for val in berry_percen_list:
            val=val.decode('utf-8')
            if "%" in val:
                new_berry_percen_list.append(float(val[:val.index("%")]))
            else:
                new_berry_percen_list.append(float("0"))
        #print(new_berry_percen_list)
        f = plt.figure()
        plt.bar(new_control_number_list,new_berry_percen_list, color='green')
        
        plt.title('Contron Number vs Berry % from '+str(start_date)+' to '+str(end_date)+' at '+site)
        for x,y in zip(new_control_number_list,new_berry_percen_list):
            label = "{:.2f}".format(y)
            plt.annotate(label, # this is the text
                 (x,y-2), # this is the point to label
                 textcoords="offset points", # how to position the text
                 xytext=(0,1), # distance from text to points (x,y)
                 ha='center',rotation=90)
        plt.xticks(new_control_number_list,new_control_number_list)
        #plt.show()
        plt.setp(plt.gca().get_xticklabels(), rotation=90, horizontalalignment='right')
        plt.savefig('/home/pi/reports/Controlnovsberrypercentage.png', dpi = 1000)
    else:
        bot.sendMessage(chat_id,'Entered Start Date is not in the list.')

#control_no_vs_berry(4,"16.02.2021","19.02.2021",'bot','chat_id')
