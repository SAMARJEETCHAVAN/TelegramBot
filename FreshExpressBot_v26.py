# -*- coding: utf-8 -*-
import gspread
import telepot#imports telepot module, which is installed for handling and creating a bot on telegram.
import datetime
from time import sleep
from difflib import SequenceMatcher
import sys
import time
import os, io
from PIL import Image, ImageOps,ImageEnhance 
from PIL import ImageFilter
from google.cloud import vision
from google.cloud.vision import types
import pandas as pd
from datetime import timedelta
from datetime import datetime
from Control_no_vs_Berry import *
import socket
from pydrive.auth import GoogleAuth
from pydrive.drive import GoogleDrive
#import detectBoxvtwo
import DetectBoxVThree
from os import listdir
from os.path import isfile, join
from leafarea import findleafarea


gauth = GoogleAuth()
drive = GoogleDrive(gauth)

sleep(100)
os.environ['GOOGLE_APPLICATION_CREDENTIALS']='/home/pi/freshexpress_jsonkey/freshex_in.json'
os.environ['GRPC_DNS_RESOLVER'] = 'native'

client = vision.ImageAnnotatorClient()
def detect_labels(path):
    """Detects labels in the file."""
    from google.cloud import vision
    import io
    client = vision.ImageAnnotatorClient()

    with io.open(path, 'rb') as image_file:
        content = image_file.read()

    image = vision.Image(content=content)

    response = client.label_detection(image=image)
    labels = response.label_annotations
    print('Labels:')

    for label in labels:
        print(label.description)

    if response.error.message:
        raise Exception(
            '{}\nFor more info on error messages, check: '
            'https://cloud.google.com/apis/design/errors'.format(
                response.error.message))
def Diff(li1, li2):
    return (list(list(set(li1)-set(li2)) + list(set(li2)-set(li1))))
def translate_text(target, text):
    import six
    from google.cloud import translate_v2 as translate

    translate_client = translate.Client()

    if isinstance(text, six.binary_type):
        text = text.decode("utf-8")

    # Text can also be a sequence of strings, in which case this method
    # will return a sequence of results for each text.
    result = translate_client.translate(text, target_language=target)

    #print(u"Text: {}".format(result["input"]))
    return (u"Translation: {}".format(result["translatedText"]))
    #print(u"Detected source language: {}".format(result["detectedSourceLanguage"]))
def detectText(img):
    with io.open(img,'rb') as image_file:
        content = image_file.read()
    
    image = vision.types.Image(content=content)
    response = client.text_detection(image=image)
    texts = response.text_annotations
    df = pd.DataFrame(columns=['locale','description'])
    
    for text in texts:
        df = df.append(
            dict(locale=text.locale,description=text.description),ignore_index=True)
    return ((df['description'][0]).encode('utf-8'))


def similar(a, b):
    return SequenceMatcher(None, a, b).ratio()

gc = gspread.service_account(filename='/home/pi/freshexpress_jsonkey/fresh-express-ocr-efe99845e1f0.json')
gc1 = gspread.service_account(filename='/home/pi/freshexpress_jsonkey/freshex_in.json')
sh = gc.open_by_key('1xF7ZZZ3CcACiam_6xSytpsYOaiPWNZMgJwUraZQAs8w')
bot = telepot.Bot('894499253:AAF-lnWrFJKdVvI0_KGWYncx0tO9nfszc20')
flag_for_S1_init = 0

def send_S1_intimation():
    global current_name,bot,flag_for_S1_init
    now = datetime.datetime.now()
    dt_string = now.strftime("%H:%M")
    if flag_for_S1_init == 0:
        if (int(str(dt_string)[:2]) >= 12):
            if (int(str(dt_string)[:2]) > 01):
                i = 1
                worksheet = sh.worksheet('Telegram ID')
                for each_value in worksheet.col_values(6):
                    if 'SUP' in each_value:
                        supervisor_id =str(worksheet.cell(i,5).value)
                        bot.sendMessage(supervisor_id,'Hello '+ str(worksheet.cell(i,2).value)+'!\nNone of the supervisors have sent S1 command yet.')
                    i = i + 1
                flag_for_S1_init = 1
                print('Intimation for S1 sent')
    elif flag_for_S1_init == 1:
        if (int(str(dt_string)[:2])==11):
            if (int(str(dt_string)[3:])> 55):
                flag_for_S1_init = 0
def verify_gggn(chat_id,GGGN):
    GGGN = GGGN.replace(" ","")
    GGGN = GGGN.replace("\n","")
    sh1 = gc1.open_by_key('1VnaxfZKonE418oPldgDqcL2_KkaN1DaZpdsTt-MyZEU')
    worksheet = sh1.worksheet('Nasik')
    ggn_list = worksheet.col_values(2)                            
    n = 1
    gggnflg=0
    for ggn in ggn_list:
        if str(ggn)in GGGN:
            gggnflg = 1
            bot.sendMessage(chat_id,"GGN OKAY!")
            if str(worksheet.cell(n,3).value)=="Fairtrade":
                bot.sendMessage(chat_id,"Fairtrade OKAY!")
            else:
                bot.sendMessage(chat_id,"Problem with Fairtrade!\n Please Verify!")    
            break
        n = n+1                        
    if gggnflg == 0:
        bot.sendMessage(chat_id,"GGN not in the list at\nhttps://docs.google.com/spreadsheets/d/1VnaxfZKonE418oPldgDqcL2_KkaN1DaZpdsTt-MyZEU/edit?usp=sharing\nPlease Verify!")
def verify_gggn2(chat_id,GGGN):
    GGGN = GGGN.replace(" ","")
    GGGN = GGGN.replace("\n","")
    sh1 = gc1.open_by_key('1VnaxfZKonE418oPldgDqcL2_KkaN1DaZpdsTt-MyZEU')
    worksheet = sh1.worksheet('Nasik')
    ggn_list = worksheet.col_values(2)                            
    n = 1
    gggnflg=0
    for ggn in ggn_list:
        if str(ggn)in GGGN:
            gggnflg = 1
            bot.sendMessage(chat_id,"GGN OKAY!")
      
            break
        n = n+1                        
    if gggnflg == 0:
        bot.sendMessage(chat_id,"GGN not in the list at\nhttps://docs.google.com/spreadsheets/d/1VnaxfZKonE418oPldgDqcL2_KkaN1DaZpdsTt-MyZEU/edit?usp=sharing\nPlease Verify!")

def check_ocr(chat_id,detected_text,image_name):
    print(detected_text)
    upload_file_list = [image_name]
    for upload_file in upload_file_list:
        gfile = drive.CreateFile({'parents': [{'id': '1kqP50wKpBNC3H6JbhPtPa-LOBAJjnjCQ'}]})
        # Read file and set it as a content of this instance.
        gfile.SetContentFile(upload_file)
        gfile.Upload() # Upload the file.
    now = str(datetime.datetime.now())
    weekno= str(datetime.date(int(now[:4]),int(now[5:7]),int(now[8:10])).isocalendar()[1])
    dayno= str(datetime.date(int(now[:4]),int(now[5:7]),int(now[8:10])).isocalendar()[2])
    #print(weekno,dayno)
    weekno1=''
    weekno2=''
    weekno3=''
    dayno1=''
    dayno2=''
    dayno3=''
    if len(weekno) == 1:
        weekno1 = '0'+str(weekno)
        weekno2 = 'O'+str(weekno)
    elif len(weekno)== 2:
        weekno3 = str(weekno)
    if len(dayno) == 1:
        dayno1 = '0'+str(dayno)
        dayno2 = 'O'+str(dayno)
    elif len(dayno)== 2:
        dayno3 = str(dayno)
    lcodelist=[]
    correctlcode1 = 'L'+weekno1+dayno1
    lcodelist.append(str(correctlcode1))
    correctlcode2 = 'L'+weekno1+dayno2
    lcodelist.append(str(correctlcode2))
    correctlcode3 = 'L'+weekno1+dayno3
    lcodelist.append(str(correctlcode3))
    correctlcode4 = 'L'+weekno2+dayno1
    lcodelist.append(str(correctlcode4))
    correctlcode5 = 'L'+weekno2+dayno2
    lcodelist.append(str(correctlcode5))
    correctlcode6 = 'L'+weekno2+dayno3
    lcodelist.append(str(correctlcode6))
    correctlcode7 = 'L'+weekno3+dayno1
    lcodelist.append(str(correctlcode7))
    correctlcode8 = 'L'+weekno3+dayno2
    lcodelist.append(str(correctlcode8))
    correctlcode9 = 'L'+weekno3+dayno3
    lcodelist.append(str(correctlcode9))

    if(("Peragri Alliance Daniken CH" in detected_text)and("Trauben weiss kernlos" in detected_text)):
        if("Produkt/Produit Trauben weiss kernlos" in detected_text[:detected_text.index("GGN")]):
            bot.sendMessage(chat_id,"CUSTOMER = PPO-Fairtrade White 16mm")
            lcode= str(detected_text[detected_text.index("Lot Nr.")+7:detected_text.index("G.GGN")])
            bot.sendMessage(chat_id,lcode)
            lcodeflag=0
            for code in lcodelist:
                if lcode[:5]==code:
                    lcodeflag=1
                    break
                else:
                    lcodeflag=0
            if lcodeflag==1:
                bot.sendMessage(chat_id,"L-code OKAY!")
            else:
                bot.sendMessage(chat_id,'Today\'s Week number :'+weekno)
                bot.sendMessage(chat_id,'Today\'s Day number :'+dayno)                        
                bot.sendMessage(chat_id,"There is a problem with L-code,\nPlease verify!")
            GGNcode = str(detected_text[detected_text.index("India\nGGN."):detected_text.index("Producteur")])
            #print(GGNcode)
            if "4049928985569" in GGNcode:
                bot.sendMessage(chat_id,"GGN OKAY!")
            else:
                bot.sendMessage(chat_id,"GGN is not correct.")
            if "FLO ID 41473\nPeragri Alliance Daniken CH" in detected_text:
                bot.sendMessage(chat_id,"Importeur/Importateur OKAY!")
            else:
                bot.sendMessage(chat_id,"Problem with Importeur/Importateur")
            if "Thompson Seedless" in detected_text[:detected_text.index("Indien")]:
                bot.sendMessage(chat_id,"Sorte/Variete OKAY!")
            else:
                bot.sendMessage(chat_id,"Problem with Sorte/Variete")
            if "Indien" in detected_text:
                bot.sendMessage(chat_id,"Ursprung/Origine OKAY!")
            else:
                bot.sendMessage(chat_id,"Problem with Ursprung/Origine")
            GGGN = detected_text[detected_text.index("G.GGN")+5:detected_text.index("Exporteur/Exportateur:")]
            verify_gggn(chat_id,GGGN)
            if "Exporteur/Exportateur: FLO ID 37975" in detected_text:
                bot.sendMessage(chat_id,"Exporteur/Exportateur OKAY!")
            else:
                bot.sendMessage(chat_id,"Problem with Exporteur/Exportateur")
            detected_text = detected_text.replace(" ","")
            detected_text = detected_text.replace("\n","")
            barcode = detected_text[-13:]
            #print(barcode)
            barcodepercentage = similar(barcode,'7613356150393')
            #bot.sendMessage(chat_id,barcode)
            if barcodepercentage>0.8:
                bot.sendMessage(chat_id,"BARCODE OKAY!")
            else:
                bot.sendMessage(chat_id,"BARCODE not OKAY!\nPlease verify")
    elif(("GENERIC" in detected_text)and("5kg" in detected_text)and("AMF" in detected_text)):
        if("AMF" in detected_text[detected_text.index("Seedless"):detected_text.index("GGN")]):
            if("5kg" in detected_text[detected_text.index("Seedless"):detected_text.index("GGN")]):
                bot.sendMessage(chat_id,"CUSTOMER = AMF 17mm box label")
                barcode = detected_text[:detected_text.index("Co")]
                barcode = barcode.replace(" ","")
                lcode = barcode[-6:]
                barcode = barcode[:barcode.index(lcode)]
                barcodepercentage = similar("(01)08906026030170(10)",barcode)
                if (barcodepercentage>0.9):
                    bot.sendMessage(chat_id,"Barcode OKAY!")
                else:
                    bot.sendMessage(chat_id,"Problem with Barcode")
                if ("Thompson Seedless" in detected_text[:detected_text.index("GENERIC")]):
                    bot.sendMessage(chat_id,"Variety OKAY!")
                else:
                    bot.sendMessage(chat_id,"Problem with Variety")
                GGGN = detected_text[detected_text.index("5kg"):detected_text.index("UK")]
                verify_gggn(chat_id,GGGN)
                lcodeflag=0
                for code in lcodelist:
                    if lcode[:5] in code:
                        lcodeflag=1
                        break
                    else:
                        lcodeflag=0
                if lcodeflag==1:
                    bot.sendMessage(chat_id,"L-code OKAY!")
                else:
                    bot.sendMessage(chat_id,'Today\'s Week number :'+weekno)
                    bot.sendMessage(chat_id,'Today\'s Day number :'+dayno)                        
                    bot.sendMessage(chat_id,"There is a problem with L-code,\nPlease verify!")
                if("UK" in detected_text[detected_text.index("5kg"):detected_text.index("Batch")]):
                   bot.sendMessage(chat_id,"Target Market Okay!")
                else:
                   bot.sendMessage(chat_id,"Problem with Target Market!")
                if ("8906026039982" in detected_text[detected_text.index("PH GGN"):detected_text.index("4049928985569")]):
                    bot.sendMessage(chat_id,"PH GLN OKAY for MIRAJ!")
                elif ("8906026039982" in detected_text[detected_text.index("PH GGN"):detected_text.index("4049928985569")]):
                    bot.sendMessage(chat_id,"PH GLN OKAY for NASHIK!")
                else:
                    bot.sendMessage(chat_id,"Problem with PH GLN")
                if ("4049928985569" in detected_text[detected_text.index("PH GGN"):detected_text.index("Fresh Express")]):
                    bot.sendMessage(chat_id,"PH GGN OKAY for NASHIK!")
                else:
                    bot.sendMessage(chat_id,"Problem with PH GGN")                        
    elif(("UNIGRAPE" in detected_text)and("5kg" in detected_text)and("FRU" in detected_text)):
    #elif("UNIGRAPE" in detected_text[detected_text.index("Seedless"):detected_text.index("Grower GGN")]):
        if("FRU" in detected_text[detected_text.index("Seedless"):detected_text.index("Grower GGN")]):
            if("5kg" in detected_text[detected_text.index("Seedless"):detected_text.index("Grower GGN")]):
                bot.sendMessage(chat_id,"CUSTOMER = FRU 16mm box label")
                #print(detected_text)
                #bot.sendMessage(chat_id,detected_text)
                barcode = detected_text[:detected_text.index("Co")]
                barcode = barcode.replace(" ","")
                barcode = barcode.replace("\n","")
                lcode = barcode[-6:]
                print('Detected Barcode : ')
                print(barcode)
                print('Detected lcode : ')
                print(lcode)
                barcode = barcode[:barcode.index(lcode)]
                barcodepercentage = similar("(01)08906026030170(10)",barcode)
                if (barcodepercentage>0.7):
                    bot.sendMessage(chat_id,"Barcode OKAY!")
                else:
                    bot.sendMessage(chat_id,"Problem with Barcode")
                if ("Thompson Seedless" in detected_text[detected_text.index("Variety"):detected_text.index("FRU")]):
                    bot.sendMessage(chat_id,"Variety OKAY!")
                else:
                    bot.sendMessage(chat_id,"Problem with Variety")
                GGGN = detected_text[detected_text.index("5kg"):detected_text.index("Batch")]
                verify_gggn2(chat_id,GGGN)
                lcodeflag=0
                for code in lcodelist:
                    if lcode[:5] in code:
                        lcodeflag=1
                        break
                    else:
                        lcodeflag=0
                if lcodeflag==1:
                    bot.sendMessage(chat_id,"L-code OKAY!")
                else:
                    bot.sendMessage(chat_id,'Today\'s Week number :'+weekno)
                    bot.sendMessage(chat_id,'Today\'s Day number :'+dayno)                        
                    bot.sendMessage(chat_id,"There is a problem with L-code,\nPlease verify!")
                if "EU" in detected_text[detected_text.index("Target Market"):detected_text.index("Batch")]:
                    bot.sendMessage(chat_id,"Market OKAY!")
                else:
                    bot.sendMessage(chat_id,"Problem with Market")
                if ("8906026030019" in detected_text):
                    bot.sendMessage(chat_id,"PH GLN OKAY for MIRAJ!")
                elif ("8906026039982" in detected_text):
                    bot.sendMessage(chat_id,"PH GLN OKAY for NASHIK!")
                else:
                    bot.sendMessage(chat_id,"Problem with PH GLN")
                if ("4049928985569" in detected_text):
                    bot.sendMessage(chat_id,"PH GGN OKAY!")
                else:
                    bot.sendMessage(chat_id,"Problem with PH GGN")

    elif(("FRU" in detected_text)and("UNI" in detected_text)and("PREMIUM" in detected_text)and("4kg" in detected_text)):
    #("UNI PREMIUM" in detected_text[detected_text.index("Seedless"):detected_text.index("Grower GGN")]):
        if("FRU" in detected_text[detected_text.index("Seedless"):]):
            if("4kg" in detected_text[detected_text.index("Seedless"):]):
                bot.sendMessage(chat_id, "CUSTOMER = FRU-Premium 18mm")
                #print(detected_text)
                #bot.sendMessage(chat_id,detected_text)
                barcode = detected_text[:detected_text.index("Co")]
                barcode = barcode.replace(" ", "")
                barcode = barcode.replace("\n", "")
                lcode = barcode[-6:]
                barcode = barcode[:barcode.index(lcode)]
                barcodepercentage = similar("(01)08906026030170(10)",barcode)
                if (barcodepercentage>0.8):
                    bot.sendMessage(chat_id,"Barcode OKAY!")
                else:
                    bot.sendMessage(chat_id,"Problem with Barcode")
                if("Thompson Seedless" in detected_text[detected_text.index("GR"):detected_text.index("PREMIUM")]):
                    bot.sendMessage(chat_id,"Variety OKAY!") 
                else:
                    bot.sendMessage(chat_id,"Problem with Variety")
                GGGN = detected_text[detected_text.index("4kg"):detected_text.index("EU")]
                verify_gggn(chat_id,GGGN)
                lcodeflag=0
                for code in lcodelist:
                    if lcode[:5] in code:
                        lcodeflag=1
                        break
                    else:
                        lcodeflag=0
                if lcodeflag==1:
                    bot.sendMessage(chat_id,"L-code OKAY!")
                else:
                    bot.sendMessage(chat_id,'Today\'s Week number :'+weekno)
                    bot.sendMessage(chat_id,'Today\'s Day number :'+dayno)                        
                    bot.sendMessage(chat_id,"There is a problem with L-code,\nPlease verify!")
                if "EU" in detected_text[detected_text.index("Target Market"):detected_text.index("Batch")]:
                    bot.sendMessage(chat_id, "Market OKAY!")
                else:
                    bot.sendMessage(chat_id,"Problem with Market")
                if ("8906026030019" in detected_text[detected_text.index("PH GLN"):detected_text.index("Fresh Express Logistics Pvt Ltd,Miraj 416 410")]):
                    bot.sendMessage(chat_id,"PH GLN OKAY for MIRAJ!")
                elif ("8906026039982" in detected_text[detected_text.index("PH GLN"):detected_text.index("Fresh Express Logistics Pvt Ltd,Miraj 416 410")]):
                    bot.sendMessage(chat_id,"PH GLN OKAY for NASHIK!")
                else:
                    bot.sendMessage(chat_id,"Problem with PH GLN")
                if ("4049928985569" in detected_text[detected_text.index("PH GGN"):detected_text.index("Fresh Express Logistics Pvt Ltd,Miraj 416 410")]):
                    bot.sendMessage(chat_id,"PH GGN OKAY!")
                else:
                    bot.sendMessage(chat_id,"Problem with PH GGN")
    elif(("SAF" in detected_text)and("SAFPRO" in detected_text)and("8.2kg" in detected_text)):
    #("SAFPRO" in detected_text[detected_text.index("Seedless"):detected_text.index("Class")]):
        if("SAF" in detected_text[detected_text.index("SAFPRO"):detected_text.index("8.2kg")]):
            if("8.2kg" in detected_text[detected_text.index("SAF"):detected_text.index("Class")]):
                bot.sendMessage(chat_id, "Customer = SAFPRO 16mm")
                #print(detected_text)
                #bot.sendMessage(chat_id,detected_text)
                barcode = detected_text[:detected_text.index("Co")]
                barcode = barcode.replace(" ", "")
                barcode = barcode.replace("\n", "")
                lcode = barcode[-6:]
                barcode = barcode[:barcode.index(lcode)]
                barcodepercentage = similar("(01)08906026030170(10)",barcode)
                if (barcodepercentage>0.8):
                    bot.sendMessage(chat_id,"Barcode OKAY!")
                else:
                    bot.sendMessage(chat_id,"Problem with Barcode")
                if("Thompson Seedless" in detected_text[detected_text.index("Variety"):detected_text.index("SAFPRO")]):
                    bot.sendMessage(chat_id,"Variety OKAY!") 
                else:
                    bot.sendMessage(chat_id,"Problem with Variety")
                if "CA" in detected_text[detected_text.index("Market"):detected_text.index("PH GLN")]:
                    bot.sendMessage(chat_id, "Market OKAY!")
                else:
                    bot.sendMessage(chat_id,"Problem with Market")
                if ("8906026030019" in detected_text[detected_text.index("PH GLN"):detected_text.index("Fresh Express")]):                                                                                                                                                                                                                        
                    bot.sendMessage(chat_id,"PH GLN OKAY for MIRAJ!")
                elif ("8906026039982" in detected_text[detected_text.index("PH GLN"):detected_text.index("Fresh Express")]):
                    bot.sendMessage(chat_id,"PH GLN OKAY for NASHIK!")
                else:
                    bot.sendMessage(chat_id,"Problem with PH GLN")
                if ("4049928985569" in detected_text[detected_text.index("GGN"):detected_text.index("Fresh Express Logistics")]):
                    bot.sendMessage(chat_id,"GGN OKAY!")
                else:
                    bot.sendMessage(chat_id,"Problem with GGN")

    elif(("SFC" in detected_text)and("GENERIC" in detected_text)and("8.2kg" in detected_text)):
    #("GENERIC" in detected_text[detected_text.index("Seedless"):detected_text.index("Class")]):
        if("SFC" in detected_text[detected_text.index("GENERIC"):detected_text.index("8.2kg")]):
            if("8.2kg" in detected_text[detected_text.index("SFC"):detected_text.index("Class")]):
                bot.sendMessage(chat_id, "Customer = SFC 16mm")
                #print(detected_text)
                #bot.sendMessage(chat_id,detected_text)
                if "Comm" in detected_text:
                    barcode = detected_text[:detected_text.index("Comm")]
                else:
                    barcode = detected_text[:detected_text.index("Variety")]
                barcode = barcode.replace(" ","")
                lcode = barcode[-6:]
                barcode = barcode[:barcode.index(lcode)]
                if (barcode=="(01)08906026030286(10)"):
                    bot.sendMessage(chat_id,"Barcode OKAY!")
                else:
                    bot.sendMessage(chat_id,"Problem with Barcode")
                if ("Thompson Seedless" in detected_text[detected_text.index("Variety"):detected_text.index("GENERIC")]):
                    bot.sendMessage(chat_id,"Variety OKAY!")
                else:
                    bot.sendMessage(chat_id,"Problem with Variety")
                if "CA" in detected_text[detected_text.index("Target Market"):detected_text.index("Batch")]:
                    bot.sendMessage(chat_id,"Market OKAY!")
                else:
                    bot.sendMessage(chat_id,"Problem with Market")
                if ("8906026030019" in detected_text[detected_text.index("PH GLN"):detected_text.index("Express")]):
                    bot.sendMessage(chat_id,"PH GLN OKAY for MIRAJ!")
                elif ("8906026039982" in detected_text[detected_text.index("PH GLN"):detected_text.index("Express")]):
                    bot.sendMessage(chat_id,"PH GLN OKAY for NASHIK!")
                else:
                    bot.sendMessage(chat_id,"Problem with PH GLN")
                if "GGN" in detected_text:
                    if ("4049928985569" in detected_text[detected_text.index("GGN"):detected_text.index("Express")]):
                        bot.sendMessage(chat_id,"GGN OKAY!")
                elif("PH GLN") in detected_text:
                    if ("4049928985569" in detected_text[detected_text.index("PH GLN"):detected_text.index("Express")]):
                        bot.sendMessage(chat_id,"GGN OKAY!")
                else:
                    bot.sendMessage(chat_id,"Problem with GGN")
    elif(("Abgepackt"in detected_text)and("pour"in detected_text)and("Peragri Alliance Daniken CH" in detected_text)and("Trauben mix kernarm" in detected_text)):
        bot.sendMessage(chat_id,"CUSTOMER = PPO-Fairtrade Mix 16mm Label")
        lcode= str(detected_text[detected_text.index("Lot Nr.")+7:detected_text.index("Gewicht")])
        lcode.replace(" ","")
        lcode.replace("\n","")
        bot.sendMessage(chat_id,lcode)
        lcodeflag=0
        for code in lcodelist:
            if lcode[:5]==code:
                lcodeflag=1
                break
            else:
                lcodeflag=0
        if lcodeflag==1:
            bot.sendMessage(chat_id,"L-code OKAY!")
        else:
            bot.sendMessage(chat_id,'Today\'s Week number :'+weekno)
            bot.sendMessage(chat_id,'Today\'s Day number :'+dayno)                        
            bot.sendMessage(chat_id,"There is a problem with L-code,\nPlease verify!")
        GGNcode = str(detected_text[detected_text.index("FLO ID 37975\nGGN"):detected_text.index("Produzent")])
        if "4049928985569" in GGNcode:
            bot.sendMessage(chat_id,"GGN OKAY!")
        else:
            bot.sendMessage(chat_id,"GGN is not correct.")
        if (("Peragri Alliance Daniken CH" in detected_text)and("FLO ID 41473" in detected_text)):
            bot.sendMessage(chat_id,"Importeur/Importateur OKAY!")
        else:
            bot.sendMessage(chat_id,"Problem with Importeur/Importateur")
        if(("Thompson seedless" in detected_text)and("Sharad Seedless" in detected_text)):
            bot.sendMessage(chat_id,"Sorte/Variete OKAY!")
        else:
            bot.sendMessage(chat_id,"Problem with Sorte/Variete")
        if (("Indien" in detected_text)or("Inde" in detected_text)):
            bot.sendMessage(chat_id,"Ursprung/Origine OKAY!")
        else:
            bot.sendMessage(chat_id,"Problem with Ursprung/Origine")
        GGGN = detected_text[detected_text.index("G.GGN")+5:detected_text.index("Ursprung")]
        print(GGGN)
        GGGN.replace(" ","")
        GGGN.replace("\n","")
        GGGN.replace(":","")
        GGGN1=GGGN[:GGGN.index("+")]
        print(GGGN1)
        bot.sendMessage(chat_id,"First")
        verify_gggn(chat_id,GGGN1)
        GGGN2=GGGN[GGGN.index("+")+1:]
        print(GGGN2)
        bot.sendMessage(chat_id,"Second")
        verify_gggn(chat_id,GGGN2)
        if "FLO ID 37975" in detected_text:
            bot.sendMessage(chat_id,"Exporteur/Exportateur OKAY!")
        else:
            bot.sendMessage(chat_id,"Problem with Exporteur/Exportateur")
        detected_text = detected_text.replace(" ","")
        detected_text = detected_text.replace("\n","")
        barcode = detected_text[-13:]
        print(barcode)
        barcodepercentage = similar(barcode,'7613379246882')
        #bot.sendMessage(chat_id,barcode)
        if barcodepercentage>0.8:
            bot.sendMessage(chat_id,"BARCODE OKAY!")
        else:
            bot.sendMessage(chat_id,"BARCODE not OKAY!\nPlease verify")
        if(("Fresh Express Logistics Pvt Ltd" in detected_text)and("FLO ID 37975" in detected_text)):
            bot.sendMessage(chat_id,"Exporteur/Exportateur OKAY!")
        else:
            bot.sendMessage(chat_id,"Exporteur/Exportateur not OKAY!\nPlease verify")
        if(("Vinifex Agro Producer Co Ltd" in detected_text)and("FLO ID 37976" in detected_text)):
            bot.sendMessage(chat_id,"Produzent/Producteur OKAY!")
        else:
            bot.sendMessage(chat_id,"Produzent/Producteur not OKAY!\nPlease verify")            
    elif("MATERIAL INWARD SLIP" in detected_text):
        print(detected_text)
        bot.sendMessage(chat_id,'Material Inward Detected')
        sh1 = gc.open_by_key('1UwRf-1FThSNB1u9MZc4r8wo9tXh9BxRevVHoI4vDXE4')
        worksheet1 = sh1.worksheet('Sheet1')
        next_row = len(worksheet1.col_values(27))+1
        DetectedText = detected_text
        materialin = {'DATE':'','SR. NO.':'','FARM OUT TIME':'','P.H. IN TIME':'','GROWER CODE':'','GROWER NAME':'','ADDRESS':'','MOBILE':'','MH CODE':'','VEHICLE NO.':'','ITEM TYPE':'','TOTAL CRATES':'','5 KG QTY':'','8 KG QTY':'','TOTAL WT INCLUDING CRATES':'','WEIGHT OF EMPTY CRATES':'','GOODS WEIGHT':'','BERRY WEIGHT':'','BUNCHES WEIGHT':'','ACCEPTED QTY':'','BERRY WASTAGE':'','AVERAGE WT':'','SUPERVISOR NAME':'','DRIVER':'','NOTES':''}
        def update_data(n,data_val):
            worksheet1.update_cell(next_row,n,data_val)
            sleep(1)
        def findinbetweenvalue(start,startcorrect,stop,stopcorrect):
            value = DetectedText[DetectedText.index(start)+startcorrect:DetectedText.index(stop)+stopcorrect]
            value = value.replace("\n","")
            return value
        update_data(27,DetectedText)  
        DetectBoxVThree.box_extraction(image_name)
        mypath = '/home/pi/Output/'
        onlyfiles = [f for f in listdir(mypath) if isfile(join(mypath, f))]
        alltexts= []
        for files in onlyfiles:
            try:
                text_here = str(detectText(mypath+files).replace("\n",""))
                alltexts.append(text_here)
                print(files,text_here)
            except:
                temp = 'temp'
            #os.remove(mypath+files)
        #print('alltexts')
        for i in alltexts:
#            print('i',i)     
            for j in materialin.keys():
#                print('j',j)
                if j in i:
                    #i = i[i.index(j):]
                    print(j,i)
                    if i[-1]==".":
                        i = i.replace(".","")
                    if i[0]==".":
                        i = i.replace(".","")
                    i = i.replace(j,"")
                    i = i.replace("=","")
                    i = i.replace("IN % ","")
                    i = i.replace("%3D","")
                    i = i.replace("%3","")
                    i = i.replace("%","")
                    i = i.replace(". =","")
                    i = i.replace("'S NAME","")
                    i = i.replace("!","")
                    i = i.replace("/CRATE IN Kg","")
                    i = i.replace("-",".")
                    materialin[j] = i
                    #print(i)
                else:
                    temp = 'temp'
#                    print(j,'not in ',i)
        for k in materialin:
            print(k,materialin[k])
        for valu in materialin:
            try:
                materialin[valu] = materialin[valu].replace(" ","")
                if valu == 'DATE':
                    materialin[valu]=materialin[valu].replace(" ","")
                    materialin[valu] = materialin[valu].replace("Q","2")
                    materialin[valu] = materialin[valu][:materialin[valu].index('2021')+4]
                    if len(materialin[valu])>10:
                        (materialin[valu])=(materialin[valu])[:10]
                    if ((materialin[valu] !="")and len(materialin[valu])>=4):
                        update_data(2,str(materialin[valu]))
                        print(valu,'from segmentation')
                    else:
                        that_date =DetectedText[DetectedText.index('DATE')+len('DATE'):DetectedText.index('SR. NO.')+0]
                        that_date = that_date.replace("\n","")
                        that_date = that_date.replace("=","")
                        that_date = that_date.replace("F","")
                        that_date = that_date.replace("Q","2")
                        update_data(2,that_date)
                        print(valu,'from text')        
                elif valu == 'SR. NO.':
                    materialin[valu] = materialin[valu].replace(" ","")
                    if ((materialin[valu] !="")and len(materialin[valu])==4):
                        update_data(1,str(materialin[valu]))
                        print(valu,'from segmentation')
                    else:
                        if((DetectedText.index('FARM OUT TIME'))>(DetectedText.index('SR. NO.'))):
                            SR_No = findinbetweenvalue('SR. NO.',len('SR. NO.'),'FARM OUT TIME',0)
                        else:
                            indexoSR_No = DetectedText.index('SR. NO.')+len('SR. NO. =')
                            nextequalstosign = indexoSR_No + DetectedText[indexoSR_No:].index('=')
                            SR_No = DetectedText[indexoSR_No:nextequalstosign]
                        SR_No = SR_No.replace("\n","")
                        SR_No = SR_No.replace("=","")
                        update_data(1,SR_No)
                        print(valu,'from text')
                elif valu == 'FARM OUT TIME':
                    materialin[valu] = materialin[valu][materialin[valu].index(":")-2:materialin[valu].index(":")+2]
                    if ((materialin[valu] !="")and len(materialin[valu])>=4):
                        update_data(3,str(materialin[valu]))
                        print(valu,'from segmentation')
                    else:
                        FARM_OUT = findinbetweenvalue('FARM OUT TIME',len('FARM OUT TIME'),'P.H. IN TIME',0)
                        FARM_OUT = FARM_OUT.replace("\n","")
                        FARM_OUT = FARM_OUT.replace("=","")
                        FARM_OUT = FARM_OUT[FARM_OUT.index(":")-2:FARM_OUT.index(":")+2]
                        update_data(3,FARM_OUT)
                        print(valu,'from text')
                elif valu =='P.H. IN TIME':
                    if '|' in materialin[valu]:
                        materialin[valu] = materialin[valu].replace('|','1')
                    if ((materialin[valu] !="")and len(materialin[valu])>=4):
                        update_data(4,str(materialin[valu]))
                        print(valu,'from segmentation')
                    else:
                        PH_IN = findinbetweenvalue('P.H. IN TIME',len('P.H. IN TIME'),'GROWER CODE',0)
                        PH_IN = PH_IN.replace("=","")
                        PH_IN = PH_IN.replace("|","1")
                        update_data(4,PH_IN)
                        print(valu,'from text')
                elif valu =='GROWER CODE':
                    materialin[valu] = materialin[valu].replace("=","")
                    materialin[valu] = materialin[valu].replace("\n","")
                    materialin[valu] = materialin[valu].replace(" ","")
                    if ((materialin[valu] !="")and len(materialin[valu])>=15):
                        update_data(5,str(materialin[valu]))
                        print(valu,'from segmentation')
                    else:
                        grower_code = findinbetweenvalue('GROWER CODE',len('GROWER CODE'),'GROWER NAME',0)
                        grower_code = grower_code.replace("=","")
                        grower_code = grower_code.replace("\n","")
                        grower_code = grower_code.replace(" ","")
                        grower_code = grower_code[:15]
                        update_data(5,grower_code)
                        print(valu,'from text')
                elif valu =='GROWER NAME':
                    print(valu,materialin[valu])
                    materialin[valu] = materialin[valu][:materialin[valu].index("\n")]
                    print(valu,materialin[valu])
                    if ((materialin[valu] !="")):
                        update_data(6,str(materialin[valu]))
                        print(valu,'from segmentation')
                    else:
                        grower_name =findinbetweenvalue('GROWER NAME',len('GROWER NAME'),'ADDRESS',0)
                        grower_name = grower_name.replace("=","")
                        update_data(6,grower_name)
                        print(valu,'from text')
                elif valu =='ADDRESS':
                    if ((materialin[valu] !="")):
                        update_data(7,str(materialin[valu]))
                        print(valu,'from segmentation')
                    else:
                        if((DetectedText.index('MOBILE'))>(DetectedText.index('ADDRESS'))):
                            address =findinbetweenvalue('ADDRESS',len('ADDRESS'),'MOBILE',0)
                        else:
                            indexodaddr = DetectedText.index('ADDRESS = ')+len('ADDRESS = ')
                            nextequalstosign = indexodaddr + DetectedText[indexodaddr:].index('=')
                            address = DetectedText[indexodaddr:nextequalstosign]
                        address = address.replace("=","")
                        update_data(7,address)
                        print(valu,'from text')
                elif valu =='MOBILE':
                    if ((materialin[valu] !="")and len(materialin[valu])>=10):
                        update_data(8,str(materialin[valu]))
                        print(valu,'from segmentation')
                    else:
                        mobile = findinbetweenvalue('MOBILE',len('MOBILE ='),'MH CODE',0)
                        mobile = mobile.replace(" ","")
                        mobile = mobile.replace("=","")
                        if len(mobile)>10:
                            mobile = mobile[:10]
                        update_data(8,mobile)
                        print(valu,'from text')
                elif valu =='MH CODE':
                    if ((materialin[valu] !="")and len(materialin[valu])>=12):
                        update_data(9,str(materialin[valu]))
                        print(valu,'from segmentation')
                    else:
                        if((DetectedText.index('VEHICLE NO'))>(DetectedText.index('MH CODE'))):
                            mh_code =findinbetweenvalue('MH CODE',len('MH CODE'),'VEHICLE NO',0)
                        else:
                            indexoftotal_crates = DetectedText.index('MH CODE')+len('MH CODE =')
                            nextequalstosign = indexoftotal_crates + DetectedText[indexoftotal_crates:].index('=')
                            print(indexoftotal_crates,nextequalstosign)
                            mh_code = DetectedText[indexoftotal_crates:nextequalstosign]
                            print("mh_code pre",mh_code)
                            mh_code = mh_code.replace("\n","")
                            for smval in materialin.keys():
                                print(smval)
                                if smval in mh_code:
                                    mh_code = mh_code.replace(smval,"")
                                else:
                                    print(smval," not in ",mh_code)
                        print("mh_code final",mh_code)
                        mh_code = mh_code.replace("=","")
                        update_data(9,mh_code)
                        print(valu,'from text')
                        
                elif valu =='VEHICLE NO.':
                    if ((materialin[valu] !="")and len(materialin[valu])>=8):
                        update_data(10,str(materialin[valu]))
                        print(valu,'from segmentation')
                    else:
                        vehicle_no =findinbetweenvalue('VEHICLE NO',len('VEHICLE NO'),'ITEM TYPE',0)
                        vehicle_no = vehicle_no.replace("=","")
                        update_data(10,vehicle_no)
                        print(valu,'from text')
                elif valu =='ITEM TYPE':
                    if ((materialin[valu] !="")):
                        update_data(11,str(materialin[valu]))
                        print(valu,'from segmentation')
                    else:
                        itm_type =findinbetweenvalue('ITEM TYPE',len('ITEM TYPE'),'TOTAL CRATES',0)
                        itm_type = itm_type.replace("=","")
                        update_data(11,itm_type)
                        print(valu,'from text')
                elif valu =='TOTAL CRATES':
                    materialin[valu] = materialin[valu].replace("G","4")
                    materialin[valu] = materialin[valu].replace(")","1")
                    if(materialin[valu] !=""):
                        if((materialin[valu] !="")and ((type(float((materialin[valu]))) == float))):
                            update_data(12,str(materialin[valu]))
                            print(valu,'from segmentation')
                    else:
                        if((DetectedText.index('5 KG QTY'))>(DetectedText.index('TOTAL CRATES'))):
                            total_crates =findinbetweenvalue('TOTAL CRATES',len('TOTAL CRATES'),'5 KG QTY',0)
                        else:
                            indexoftotal_crates = DetectedText.index('TOTAL CRATES')+len('TOTAL CRATES = ')
                            nextequalstosign = indexoftotal_crates + DetectedText[indexoftotal_crates:].index('=')
                            total_crates = DetectedText[indexoftotal_crates:nextequalstosign]                        
                        total_crates = total_crates.replace("=","")
                        total_crates = total_crates.replace("G","4")
                        update_data(12,total_crates)
                        print(valu,'from text')
                elif valu =='5 KG QTY':
                    materialin[valu] = materialin[valu].replace(".","")
                    materialin[valu] = materialin[valu].replace(" ","")
                    materialin[valu] = materialin[valu].replace(")","1")
                    if ((materialin[valu] !="")and type(float(materialin[valu])) == float):
                        update_data(13,str(materialin[valu]))
                        print(valu,'from segmentation')
                    else:
                        kg5 =findinbetweenvalue('5 KG QTY',len('5 KG QTY'),'8 KG QTY',0)
                        kg5 = kg5.replace("=","")
                        update_data(13,kg5)
                        print(valu,'from text')
                elif valu =='8 KG QTY':
                    materialin[valu] = materialin[valu].replace(".","")
                    materialin[valu] = materialin[valu].replace(" ","")
                    materialin[valu] = materialin[valu].replace(")","1")
                    if ((materialin[valu] !="")and type(int(materialin[valu])) == int):
                        update_data(14,str(materialin[valu]))
                        print(valu,'from segmentation')
                    else:
                        kg8 =findinbetweenvalue('8 KG QTY',len('8 KG QTY'),'GROSS WT. IN Kg\'s',0)
                        kg8 = kg8.replace("=","")
                        update_data(14,kg8)
                        print(valu,'from text')
                elif valu =='TOTAL WT INCLUDING CRATES':
                    materialin[valu] = materialin[valu].replace(" ","")
                    if ((materialin[valu] !="")and type(float(materialin[valu])) == float):
                        update_data(16,str(materialin[valu]))
                        print(valu,'from segmentation')
                    else:
                        totalwt =findinbetweenvalue('TOTAL WT INCLUDING CRATES',len('TOTAL WT INCLUDING CRATES'),'WEIGHT OF EMPTY CRATES',0)
                        totalwt = totalwt.replace("=","")
                        update_data(16,totalwt)
                        print(valu,'from text')
                elif valu =='WEIGHT OF EMPTY CRATES':
                    materialin[valu] = materialin[valu].replace(" ","")
                    if ((materialin[valu] !="")and type(float(materialin[valu])) == float):
                        update_data(17,str(materialin[valu]))
                        print(valu,'from segmentation')
                    else:
                        emptycrates =findinbetweenvalue('WEIGHT OF EMPTY CRATES',len('WEIGHT OF EMPTY CRATES'),'GOODS WEIGHT',0)
                        emptycrates = emptycrates.replace("=","")
                        update_data(17,emptycrates)
                        print(valu,'from text')
                elif valu =='GOODS WEIGHT':
                    materialin[valu] = materialin[valu].replace(" ","")
                    if ((materialin[valu] !="")and type(float(materialin[valu])) == float):
                        update_data(18,str(materialin[valu]))
                        print(valu,'from segmentation')
                    else:
                        goodswt =findinbetweenvalue('GOODS WEIGHT',len('GOODS WEIGHT'),'BERRY WEIGHT',0)
                        goodswt = goodswt.replace("=","")
                        update_data(18,goodswt)
                        print(valu,'from text')
                elif valu =='BERRY WEIGHT':
                    materialin[valu] = materialin[valu].replace(" ","")
                    if ((materialin[valu] !="")and type(float(materialin[valu])) == float):
                        update_data(19,str(materialin[valu]))
                        print(valu,'from segmentation')
                    else:
                        berrywt =findinbetweenvalue('BERRY WEIGHT',len('BERRY WEIGHT'),'BUNCHES WEIGHT',0)
                        berrywt = berrywt.replace("=","")
                        update_data(19,berrywt)
                        print(valu,'from text')
                elif valu == 'BUNCHES WEIGHT':
                    while (" " in materialin[valu]):
                        materialin[valu] = materialin[valu].replace(" ","")
                    if ((materialin[valu] !="")and type(float(materialin[valu])) == float):
                        print('segment for bunches detected')
                        update_data(20,str(materialin[valu]))
                        print(valu,'from segmentation')
                    else:
                        bunchwt =findinbetweenvalue('BUNCHES WEIGHT',len('BUNCHES WEIGHT'),'ACCEPTED QTY',0)
                        bunchwt = bunchwt.replace("=","")
                        print('text for bunches detected')
                        print(bunchwt)
                        update_data(20,bunchwt)
                        print(valu,'from text')
                elif valu =='ACCEPTED QTY':
                    materialin[valu] = materialin[valu].replace(" ","")
                    if materialin[valu][0]=='.':
                        materialin[valu] = materialin[valu][1:]
                    if ((materialin[valu] !="")and type(float(materialin[valu])) == float):
                        update_data(21,str(materialin[valu]))
                        print(valu,'from segmentation')
                    else:
                        accepted =findinbetweenvalue('ACCEPTED QTY',len('ACCEPTED QTY'),'BERRY WASTAGE IN',0)
                        accepted = accepted.replace("=","")
                        update_data(21,accepted)
                        print(valu,'from text')
                elif valu =='BERRY WASTAGE':
                    materialin[valu] = materialin[valu].replace(" ","")
                    if ((materialin[valu] !="")and type(float(materialin[valu])) == float):
                        update_data(22,str(materialin[valu]))
                        print(valu,'from segmentation')
                    else:
                        berrywastage =findinbetweenvalue('BERRY WASTAGE IN',len('BERRY WASTAGE IN')+2,'AVERAGE WT/CRATE IN Kg',0)
                        berrywastage = berrywastage.replace("=","")
                        update_data(22,berrywastage)
                        print(valu,'from text')
                elif valu =='AVERAGE WT':
                    materialin[valu] = materialin[valu].replace("%","")
                    materialin[valu] = materialin[valu].replace(" ","")
                    if materialin[valu][0]=='.':
                        materialin[valu] = materialin[valu][1:]
                    if ((materialin[valu] !="")and type(float(materialin[valu])) == float):
                        update_data(23,str(materialin[valu]))
                        print(valu,'from segmentation')
                    else:
                        if((DetectedText.index('SUPERVISOR NAME'))>(DetectedText.index('AVERAGE WT/CRATE IN Kg'))):
                            avgwtpercrate =findinbetweenvalue('AVERAGE WT/CRATE IN Kg',len('AVERAGE WT/CRATE IN Kg'),'SUPERVISOR NAME',0)
                        else:
                            indexoSR_No = DetectedText.index('AVERAGE WT/CRATE IN Kg')+len('AVERAGE WT/CRATE IN Kg. =')
                            nextequalstosign = indexoSR_No + DetectedText[indexoSR_No:].index('=')
                            SR_No = DetectedText[indexoSR_No:nextequalstosign]
                        avgwtpercrate = avgwtpercrate.replace("=","")
                        update_data(23,avgwtpercrate)
                        print(valu,'from text')
                elif valu =='SUPERVISOR NAME':             
                    materialin[valu] = materialin[valu].replace("DRIVER","")
                    materialin[valu] = materialin[valu].replace(".","")
                    materialin[valu] = materialin[valu].replace("PRAKASHPAWAR","")
                    materialin[valu] = materialin[valu].replace(" ","")
                    if ((materialin[valu] !="")):
                        update_data(24,str(materialin[valu]))
                        print(valu,'from segmentation')
                    else:
                        supervisor =findinbetweenvalue('SUPERVISOR NAME',len('SUPERVISOR NAME'),'DRIVER\'S NAME',0)
                        supervisor = supervisor.replace("=","")
                        update_data(24,supervisor)
                        print(valu,'from text')
                elif valu =='DRIVER':
                    materialin[valu] = materialin[valu].replace(" ","")
                    if ((materialin[valu] !="")):
                        update_data(25,str(materialin[valu]))
                        print(valu,'from segmentation')
                    else:
                        driver =findinbetweenvalue('DRIVER\'S NAME',len('DRIVER\'S NAME'),'NOTES',0)
                        driver = driver.replace("=","")
                        update_data(25,driver)
                        print(valu,'from text')
                elif valu =='NOTES':
                    materialin[valu] = materialin[valu].replace(" ","")
                    if ((materialin[valu] !="")):
                        update_data(26,str(materialin[valu]))
                        print(valu,'from segmentation')
                    else:
                        notes =DetectedText[DetectedText.index('NOTES')+len('NOTES'):]
                        notes = notes.replace("=","")
                        update_data(26,notes)
                        print(valu,'from text')
            except:
                print('There was a error at', valu)
            print(valu,materialin[valu])
    else:
        bot.sendMessage(chat_id,detected_text)
        translated_text = translate_text("en",detected_text)
        bot.sendMessage(chat_id,"Text translated to english is:")
        bot.sendMessage(chat_id,translated_text)
    os.remove(image_name)
 

def handle(msg):
    #try:
    chat_id = msg['chat']['id']
    #print(msg)
    if 'text' in msg:
        command = msg['text']
        try:
            command = str(command)
        except:
            command = str(command.decode('utf-8'))
        worksheet = sh.worksheet('Chat commands History')
        try:
            if str(msg['from']['last_name']).decode('utf-8'):
                current_name = str(msg['from']['first_name'])+' '+str(msg['from']['last_name'])
            else:
                current_name = str(msg['from']['first_name'])
        except:
            if (msg['from']['last_name']).decode('utf-8')!='':
                current_name = str((msg['from']['first_name']).decode('utf-8'))+' '+str((msg['from']['last_name'].decode('utf-8')))
            else:
                current_name = str((msg['from']['first_name']).decode('utf-8'))
            
        try:
            now = datetime.datetime.now()
            dt_string = now.strftime("%d/%m/%Y %H:%M:%S")
            next_row = len(worksheet.col_values(2))+1
            worksheet.update_cell(next_row,1, str(dt_string))
            worksheet.update_cell(next_row,2, str(command))
            worksheet.update_cell(next_row,3, str(current_name))
            command = command.replace(" ","")
        except:
            temp = 'temp'
        print(command)
        if (('hi' == command.lower())or( command == '/start')):
            worksheet = sh.worksheet('Telegram ID')
            telegram_id_list = worksheet.col_values(5)
            names_list = worksheet.col_values(2)
            total_rows = worksheet.col_values(1)
            if msg['from']['last_name']:
                current_name = str((msg['from']['first_name']).decode('utf-8'))+' '+str((msg['from']['last_name'].decode('utf-8')))
            else:
                current_name = str((msg['from']['first_name']).decode('utf-8'))
            print(current_name)
            flg=0
            for telegram_id in telegram_id_list:
                if (str(chat_id) == str(telegram_id)):
                    flg=1
                    row_no = int(telegram_id_list.index(str(chat_id)))+1
                    namepercentage = similar((str(worksheet.cell(row_no,2).value)).lower(),current_name.lower())
                    if  namepercentage > 0.8:
                        bot.sendMessage(chat_id,'Hello ' + current_name)
                    else:
                        bot.sendMessage(chat_id,'Sorry!\nThis Telegram Account belongs to '+str(worksheet.cell(row_no,2).value))
            if flg==0:
                i=1
                j=[]
                for name in names_list:
                    namepercentage = similar(name.lower(),current_name.lower())
                    if  namepercentage > 0.8:
                        j.append(i)
                    i=i+1
                if j:
                    worksheet.update_cell(j[0],5, str(chat_id))
                else:
                    new_row=len(total_rows)+1
                    worksheet.update_cell(new_row,1, str(new_row-1))
                    worksheet.update_cell(new_row,2, current_name)
                    worksheet.update_cell(new_row,5, str(chat_id))
                    bot.sendMessage(chat_id,'Hello ' + current_name)
                    bot.sendMessage(chat_id,'Welcome to Fresh express!')
                    sleep(2)
                    bot.sendMessage(chat_id,'Please enter your 10 digit mobile number')
        elif command.lower() == 'pb':
            #os.system('python /home/pi/code2getimagefromrtsp_v3.py')
            os.system('python /home/pi/code2getimagefromrtsp_v3.py')
            sleep(5)
            bot.sendMessage(chat_id,'Image Clicked Thank you!')
            
        elif (len(command)==10)and(str(command[0])=='6'or str(command[0])=='7'or str(command[0])=='8'or str(command[0])=='9'):
            flg=0
            worksheet = sh.worksheet('Telegram ID')
            telegram_id_list = worksheet.col_values(5)
            names_list = worksheet.col_values(2)
            if msg['from']['last_name']:
                current_name = str(msg['from']['first_name'])+' '+str(msg['from']['last_name'])
            else:
                current_name = str(msg['from']['first_name'])
            for telegram_id in telegram_id_list:
                if (str(chat_id) == str(telegram_id)):
                    flg=1
                    row_no = int(telegram_id_list.index(str(chat_id)))+1
                    if((str(worksheet.cell(row_no,2).value)).lower() == current_name.lower()):
                        worksheet.update_cell(row_no,4, str(command))
                        bot.sendMessage(chat_id,str(command)+' added as your commumnication mobile number!\nThank You!')
                    else:
                        bot.sendMessage(chat_id,' We couldnt add you number as your register name is different!\nThank You!')
            if flg ==0:
                bot.sendMessage(chat_id,' We couldnt add you number.\nThank You!')
        elif command.lower() == "a1":
            worksheet = sh.worksheet('A1. Real time KYC Details')
            telegram_id_list = worksheet.col_values(6)
            names_list = worksheet.col_values(4)
            if str(chat_id) in telegram_id_list:
                if str(current_name) == str(worksheet.cell(telegram_id_list.index(str(chat_id))+1,4).value):
                    text_msg = 'KYC Details:\nEmployee Name :'+str(worksheet.cell(telegram_id_list.index(str(chat_id))+1,4).value)+'\nMobile No. :'+str(worksheet.cell(telegram_id_list.index(str(chat_id))+1,5).value)+'\nBank A/c Name:'+str(worksheet.cell(telegram_id_list.index(str(chat_id))+1,10).value)+'\nAccount No. :'+str(worksheet.cell(telegram_id_list.index(str(chat_id))+1,8).value)+'\nIFSC:'+str(worksheet.cell(telegram_id_list.index(str(chat_id))+1,9).value)
                    bot.sendMessage(chat_id,text_msg)
                else:
                    bot.sendMessage(chat_id,'Sorry!\nYour name is not in the list')
            else:
                bot.sendMessage(chat_id,'Sorry!\nYou are not a registered user.')
                    
        elif command.lower()[:2]=="a2":
            if (len(command) == 5 or len(command) == 10):
                month = command.lower()[3:5]
                month_list=['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']
                current_month = month_list[int(month)-1]
                
            if(len(command.lower())==5):
                worksheet = sh.worksheet('A2 Ladies Sallary data')
                telegram_id_list = worksheet.col_values(10)
                i = 1
                j = []
                for telegram_id in telegram_id_list:
                    if str(chat_id) == str(telegram_id):
                        #print(str(current_name.lower()),str(worksheet.cell(i,2).value.lower()))
                        #if str(current_name.lower()) == str(worksheet.cell(i,2).value.lower()):
                        j.append(i)
                    i=i+1
                print(j)
                if j:
                    for row_value in j:
                        #print(current_month == str(worksheet.cell(row_value,4).value),current_month,str(worksheet.cell(row_value,4).value))
                        if current_month == str(worksheet.cell(row_value,4).value):
                            text_msg = "Name : "+str(worksheet.cell(row_value,2).value)+"\nMonth : "+str(worksheet.cell(row_value,4).value)+now.strftime("%Y")+"\nPresent Days : "+str(worksheet.cell(row_value,5).value)+"\nAdvance : "+str(worksheet.cell(row_value,8).value)+"\nAmount : "+str(worksheet.cell(row_value,9).value)
                            break
                        else:
                            text_msg = "There is no data for this month"
                    bot.sendMessage(chat_id,text_msg)
                else:
                    bot.sendMessage(chat_id,'You dont have access to this command')
                    
            elif(len(command.lower())==10):
                worksheet = sh.worksheet('Telegram ID')
                telegram_id_list = worksheet.col_values(5)
                if str(chat_id) in telegram_id_list:
                    row_value = int(telegram_id_list.index(str(chat_id)))+1
                    if 'SUP' in str(worksheet.cell(row_value,6).value):
                        worksheet2 = sh.worksheet('A2 Ladies Sallary data')
                        sifa_list = worksheet2.col_values(1)
                        i = 1
                        j = []
                        for sifa in sifa_list:
                            if str(sifa[-4:]) == str(command[6:]):
                                j.append(i)
                            i=i+1
                        if j:
                            for row_values in j:
                                if current_month == str(worksheet2.cell(row_values,4).value):
                                    row_value2 = row_values
                                    text_msg = "Name : "+str(worksheet2.cell(row_value2,2).value)+"\nMonth : "+str(worksheet2.cell(row_value2,4).value)+now.strftime("%Y")+"\nPresent Days : "+str(worksheet2.cell(row_value2,5).value)+"\nAdvance : "+str(worksheet2.cell(row_value2,8).value)+"\nAmount : "+str(worksheet2.cell(row_value2,9).value)
                                    break
                                    
                                else:
                                    text_msg = 'There is no data for this month'                                                           
                            bot.sendMessage(chat_id,text_msg)
                        else:
                            bot.sendMessage(chat_id,'Sorry!\nThere is no such Sifa number.\nPlease Verify!')    
                    else:
                        bot.sendMessage(chat_id,'Sorry!\nYou dont have this access!')
                else:
                    bot.sendMessage(chat_id,'Sorry!\nYou are not registered!')
            else:
                bot.sendMessage(chat_id,'Sorry!\nYou have entered code incorectly.\nPlease Verify!')
        elif command.lower()=='a3':
            worksheet = sh.worksheet('A1. Real time KYC Details')
            telegram_id_list = worksheet.col_values(6)
            if str(chat_id) in telegram_id_list:
                row_no = int(telegram_id_list.index(str(chat_id)))+1
                current_bank_name = str(worksheet.cell(row_no,10).value)
                current_sifa = str(worksheet.cell(row_no,2).value)
                worksheet2 = sh.worksheet('A3 Ladies Payment data')
                sifa_list = worksheet2.col_values(4)
                telegram_id_list2 = worksheet2.col_values(9)
                i=1
                j=[]
                
                for each_id in telegram_id_list2:
                    if str(each_id) == str(chat_id):
                        j.append(i)
                        
                    i=i+1
                    k=0
                
                for each_value in j:
                    text_message = '\n'+str(k+1)+':\n'+str(worksheet2.cell(each_value,5).value)+'\n'+str(worksheet2.cell(each_value,1).value)+" to "+str(worksheet2.cell(each_value,6).value)+', \nAmount'+str(worksheet2.cell(each_value,2).value)+',\n on Date:'+str(worksheet2.cell(each_value,3).value)
                    bot.sendMessage(chat_id,text_message)
                    k=k+1
                if not j:
                    bot.sendMessage(chat_id,'Sorry!\nYour name is not in A3 sheet!')
        
            else:
                bot.sendMessage(chat_id,'Sorry!\nYou have not registered yet.')
        elif command.lower()=='a4':
            bot.sendMessage(chat_id,'Please fill in the jotform on this link for leave application.\n https://bit.ly/2W1tscd')
        elif command.lower()=='a5':
            bot.sendMessage(chat_id,'PLEASE USE FOLLOWING LINK TO SUBMIT REQUEST.\nhttps://bit.ly/2JNzCKH')
        elif command[:2].lower() == 'e1':
            worksheet = sh.worksheet('E. EXPORT REGISTER')
            invoice_no_list = worksheet.col_values(4)
            invoice_no = str(command[3:])
            invoice_no = invoice_no.upper()
            if invoice_no in invoice_no_list:
                container_row = (invoice_no_list.index(invoice_no))
                bot.sendMessage(chat_id,'Container number for invoice number '+invoice_no+' is\n'+str(worksheet.cell(container_row+1,7).value))
                bot.sendMessage(chat_id,'Shipping Line : '+str(worksheet.cell(container_row+1,23).value))
                bot.sendMessage(chat_id,'ETD : '+str(worksheet.cell(container_row+1,26).value))
                bot.sendMessage(chat_id,'ETA : '+str(worksheet.cell(container_row+1,27).value))
            else:
                bot.sendMessage(chat_id,'Invalid Invoice Number!')
        elif command[:2].lower() == 'e2':
            worksheet = sh.worksheet('E. EXPORT REGISTER')
            invoice_no_list = worksheet.col_values(4)
            invoice_no = str(command[3:])
            invoice_no = invoice_no.upper()
            if invoice_no in invoice_no_list:
                msg_box=[]
                container_row = (invoice_no_list.index(invoice_no))
                std_msg='Invoice number '+invoice_no+'\nVariety Type : '+str(worksheet.cell(container_row+1,9).value)+'\nPack Type : '+str(worksheet.cell(container_row+1,10).value)+'\nBox Size : '+str(worksheet.cell(container_row+1,11).value+'\n')
                msg_box.append(std_msg)
                if str(worksheet.cell(container_row+1,12).value)!='0':
                    msg_box.append('Box Qty 4.5kg : ')
                    msg_box.append(str(worksheet.cell(container_row+1,12).value))
                    msg_box.append('\n')
                if str(worksheet.cell(container_row+1,13).value)!='0':
                    msg_box.append('Box Qty 5 kg WHITE : ')
                    msg_box.append(str(worksheet.cell(container_row+1,13).value))
                    msg_box.append('\n')
                if str(worksheet.cell(container_row+1,14).value)!='0':
                    msg_box.append('Box Qty 5KG BiCOLOR : ')
                    msg_box.append(str(worksheet.cell(container_row+1,14).value))
                    msg_box.append('\n')
                if str(worksheet.cell(container_row+1,15).value)!='0':
                    msg_box.append('Box Qty 5KG SHARAD : ')
                    msg_box.append(str(worksheet.cell(container_row+1,15).value))
                    msg_box.append('\n')
                if str(worksheet.cell(container_row+1,16).value)!='0':
                    msg_box.append('Box Quantity  8.00 kg : ')
                    msg_box.append(str(worksheet.cell(container_row+1,16).value))
                    msg_box.append('\n')
                if str(worksheet.cell(container_row+1,17).value)!='0':
                    msg_box.append('Box Quantity 8.2 kg : ')
                    msg_box.append(str(worksheet.cell(container_row+1,17).value))
                    msg_box.append('\n')
                msg_box.append('Total Box Quantity : ')
                msg_box.append(str(worksheet.cell(container_row+1,18).value))
                txt_msg=''
                for each_msg in msg_box:
                    txt_msg = txt_msg+each_msg
                bot.sendMessage(chat_id,txt_msg)
        elif command[:2].lower()=='e3':
            worksheet = sh.worksheet('E. EXPORT REGISTER')
            invoice_no_list = worksheet.col_values(4)
            invoice_no = str(command[3:])
            invoice_no = invoice_no.upper()
            if invoice_no in invoice_no_list:
                container_row = (invoice_no_list.index(invoice_no))
                txt_msg='Invoice number :'+invoice_no+'\nEstimated Date and time of Arrival : \n'+str(worksheet.cell(container_row+1,27).value)
                bot.sendMessage(chat_id,txt_msg)
            else:
                bot.sendMessage(chat_id,"Invalid Invoice")
        elif command[:2].lower()=='e4':
            worksheet = sh.worksheet('E. EXPORT REGISTER')
            invoice_no_list = worksheet.col_values(4)
            invoice_no = str(command[3:])
            invoice_no = invoice_no.upper()
            if invoice_no in invoice_no_list:
                container_row = (invoice_no_list.index(invoice_no))
                txt_msg='Invoice number :'+invoice_no+'\nReceived Advance Amount: \n'+str(worksheet.cell(container_row+1,38).value)+'\nAdvance Amount received Date: \n'+str(worksheet.cell(container_row+1,39).value)
                bot.sendMessage(chat_id,txt_msg)
            else:
                bot.sendMessage(chat_id,"Invalid Invoice")
        elif command[:2].lower()=='e5':
            worksheet = sh.worksheet('E. EXPORT REGISTER')
            invoice_no_list = worksheet.col_values(4)
            invoice_no = str(command[3:])
            invoice_no = invoice_no.upper()
            if invoice_no in invoice_no_list:
                container_row = (invoice_no_list.index(invoice_no))
                txt_msg='Invoice number :'+invoice_no+'\nFinal Balance: \n'+str(worksheet.cell(container_row+1,40).value)
                bot.sendMessage(chat_id,txt_msg)
            else:
                bot.sendMessage(chat_id,"Invalid Invoice")            
                
        
        elif command[:2].lower()=='f1':
            worksheet = sh.worksheet('Telegram ID')
            telegram_id_list = worksheet.col_values(5)
            if str(chat_id) in telegram_id_list:
                row_value = int(telegram_id_list.index(str(chat_id)))+1
                if 'SUP' in str(worksheet.cell(row_value,6).value):
                    worksheet2 = sh.worksheet('11.01 Farmer Codes')
                    query = (command[3:]).lower()
                    
                    if not query[3:].isdigit():  
                        names_list = worksheet2.col_values(2)
                        i=1
                        j=[]
                        for name in names_list:
                            name = name.replace(" ","")
                            if (query in name.lower()):
                                j.append(i)
                            i = i + 1                    
                        if j:
                            if len(j)>10:
                                bot.sendMessage(chat_id,'Sorry!\nThere are plenty of names containing this name.\n Please try a specific name.')
                            else:
                                for r_value in j:
                                    bot.sendMessage(chat_id,'Grower Name :\n'+str(worksheet2.cell(r_value,2).value)+'\nGrower Code :\n'+str(worksheet2.cell(r_value,1).value))                    
                        else:
                            bot.sendMessage(chat_id,'Sorry!\nThere is no such name.')
                    else:
                        farmercode_list = worksheet2.col_values(1)
                        s=1
                        for val in farmercode_list:
                            if query in str(val.decode('utf-8')).lower():
                                bot.sendMessage(chat_id,str(worksheet2.cell(s,2).value))                 
                            s=s+1
                                
                else:
                    bot.sendMessage(chat_id,'Sorry!\nYou are not authorised to use this command.')
        elif command[:2].lower()=='f2':
            worksheet = sh.worksheet('Telegram ID')
            telegram_id_list = worksheet.col_values(5)
            if str(chat_id) in telegram_id_list:
                row_value = int(telegram_id_list.index(str(chat_id)))+1
                if 'SUP' in str(worksheet.cell(row_value,6).value):
                    worksheet2 = sh.worksheet('F2. MRL for Telegram')
                    growercode_list = worksheet2.col_values(1)
                    i=1
                    j=[]
                    for growercode in growercode_list:
                        growercode = growercode.replace(" ","")
                        if (command[3:]).lower() in growercode.lower():
                            j.append(i)
                        i = i + 1                    
                    if j:
                        if len(j)>10:
                            bot.sendMessage(chat_id,'Sorry!\nThere are plenty of reports containing this Grower Code.')
                        else:
                            for r_value in j:
                                bot.sendMessage(chat_id,'Grower Name:\n'+str(worksheet2.cell(r_value,2).value)+'\nGrower Code :'+str(worksheet2.cell(r_value,1).value)+'\nLab Report No. :'+str(worksheet2.cell(r_value,3).value)+'\nMH Code: '+str(worksheet2.cell(r_value,4).value)+'\nTotal Elements :'+str(worksheet2.cell(r_value,5).value)+'\nTotal % :'+str(worksheet2.cell(r_value,6).value)+'\nHighest Single % :'+str(worksheet2.cell(r_value,7).value)+'\nHighest Single % Name:\n'+str(worksheet2.cell(r_value,8).value))
                    else:
                        bot.sendMessage(chat_id,'Sorry!\nThere is no data with this Grower Code.')
                            
                else:
                    bot.sendMessage(chat_id,'Sorry!\nYou are not authorised to use this command.')
            
        elif command[:2].lower()=='f3':
            query_field = str(command[3:].decode('utf-8'))
            worksheet = sh.worksheet('F3.GGN LIST')
            if not query_field.isdigit():
                firstname_list = worksheet.col_values(7)
                name = query_field
                print(name)
                name_list={}
                s=1
                for val in firstname_list:
                    if name.lower() in val.lower():
                        bot.sendMessage(chat_id,str(val.decode('utf-8'))+" : "+str((worksheet.cell(s,2).value).decode('utf-8')))
                        if "fairtrade" in str((worksheet.cell(s,3).value).decode('utf-8')).lower():
                            bot.sendMessage(chat_id,str(val)+' is is Fairtrade supllier.')
                    s = s + 1
            else:
                ggn = query_field
                ggn_list = worksheet.col_values(2)
                s=1
                for val in ggn_list:
                    if ggn in val:
                        for a in range(7,9):
                            bot.sendMessage(chat_id,str((worksheet.cell(s,a).value).decode('utf-8')))
                            sleep(2)
                        if "fairtrade" in str((worksheet.cell(s,3).value).decode('utf-8')).lower():
                            bot.sendMessage(chat_id,str((worksheet.cell(s,7).value).decode('utf-8'))+' is is Fairtrade supllier.')
                    s=s+1   
            
        elif command[1].lower()=='p':
            worksheet = sh.worksheet('Telegram ID')
            telegram_id_list = worksheet.col_values(5)
            if str(chat_id) in telegram_id_list:
                row_value = int(telegram_id_list.index(str(chat_id)))+1
                if 'SUP' in str(worksheet.cell(row_value,6).value):
                    if command[0]=='2':
                        unit = 'IDEAL'
                    elif command[0]=='3':
                        unit = 'AMAYA'
                    elif command[0]=='4':
                        unit = 'MIRAJ'
                    elif command[0]=='5':
                        unit = 'ALL'
                    else:
                        unit = "invalid"
                    if command[2]=='1':
                        if(len(command)==3):
                            worksheet2 = sh.worksheet('P: material dispatch intimation JF05')
                            location_list = worksheet2.col_values(7)
                            i=1
                            j=[]
                            for location in location_list:
                                location = location.replace(" ","")
                                if unit.lower() in location.lower():
                                   j.append(i) 
                                i=i+1
                            if not(j):
                                bot.sendMessage(chat_id,'There is no data for this location')
                            elif len(j)<4:
                                for r_value in j:
                                    txt_msg='Company Name: '+str(worksheet2.cell(r_value,4).value)+'\nItem Name: '+str(worksheet2.cell(r_value,9).value)+'\nQTY UOM.: '+str(worksheet2.cell(r_value,10).value)+str(worksheet2.cell(r_value,11).value)
                                    if (worksheet2.cell(r_value,13).value)!="":
                                        txt_msg=txt_msg + '\nItem Name 2: '+str(worksheet2.cell(r_value,13).value)+'\nQTY UOM.: :'+str(worksheet2.cell(r_value,14).value)+str(worksheet2.cell(r_value,15).value)
                                    txt_msg = txt_msg +'\nVEHICAL NO: '+str(worksheet2.cell(r_value,23).value)+'\nDRIVER NAME: '+str(worksheet2.cell(r_value,24).value)+'\nDRIVER MOBILE NO: '+str(worksheet2.cell(r_value,25).value)+'\nExpected date and time of delivery:\n'+str(worksheet2.cell(r_value,8).value)+'\nTo: '+str(worksheet2.cell(r_value,7).value)
                                    bot.sendMessage(chat_id,txt_msg)
                            else:
                                txt_msg =""
                                k=1
                                for r_value in j:                                 
                                    txt_msg =txt_msg+'\n'+'\n'+str(k)+']Company Name:\n'+str(worksheet2.cell(r_value,4).value)+'\nItem Name:\n'+str(worksheet2.cell(r_value,9).value)+'\nQTY UOM.\n'+str(worksheet2.cell(r_value,10).value)+str(worksheet2.cell(r_value,11).value)
                                    if (worksheet2.cell(r_value,13).value)!="":
                                        txt_msg=txt_msg + '\nItem Name 2:\n'+str(worksheet2.cell(r_value,13).value)+'\nQTY UOM.\n'+str(worksheet2.cell(r_value,14).value)+str(worksheet2.cell(r_value,15).value)
                                    txt_msg = txt_msg +'\nVEHICAL NO:\n'+str(worksheet2.cell(r_value,23).value)+'\nDRIVER NAME'+str(worksheet2.cell(r_value,24).value)+'\nDRIVER MOBILE NO'+str(worksheet2.cell(r_value,25).value)+'\nExpected date and time of delivery:'+str(worksheet2.cell(r_value,8).value)+'\nTo:'+str(worksheet2.cell(r_value,7).value)+'\n+|+|+|+|  +|+|+|+|  +|+|+|+|  +|+|+|+|'
                                    sleep(2)
                                    k=k+1
                                with open("test.doc",'w') as f:
                                    f.write(txt_msg)
                                    f.close()
                                bot.sendDocument(chat_id,document=open('test.doc'))
                                        
                        elif(len(command)>3):
                            no_of_entries = int(command[3:])
                            worksheet2 = sh.worksheet('P: material dispatch intimation JF05')
                            location_list = worksheet2.col_values(7)
                            i=1
                            j=[]
                            for location in location_list:
                                location = location.replace(" ","")
                                if unit.lower() in location.lower():
                                   j.append(i) 
                                i=i+1                        
                            j.reverse()
                            if (no_of_entries)<4:
                                while (no_of_entries):
                                    for r_value in j:
                                        txt_msg='Company Name:\n'+str(worksheet2.cell(r_value,4).value)+'\nItem Name:\n'+str(worksheet2.cell(r_value,9).value)+'\nQTY UOM.\n'+str(worksheet2.cell(r_value,10).value)+str(worksheet2.cell(r_value,11).value)
                                        if (worksheet2.cell(r_value,13).value)!="":
                                            txt_msg=txt_msg + '\nItem Name 2:\n'+str(worksheet2.cell(r_value,13).value)+'\nQTY UOM.\n'+str(worksheet2.cell(r_value,14).value)+str(worksheet2.cell(r_value,15).value)
                                        txt_msg = txt_msg +'\nVEHICAL NO:\n'+str(worksheet2.cell(r_value,23).value)+'\nDRIVER NAME'+str(worksheet2.cell(r_value,24).value)+'\nDRIVER MOBILE NO'+str(worksheet2.cell(r_value,25).value)+'\nExpected date and time of delivery:'+str(worksheet2.cell(r_value,8).value)+'\nTo:'+str(worksheet2.cell(r_value,7).value)
                                        bot.sendMessage(chat_id,txt_msg)
                                        (no_of_entries)=(no_of_entries)-1
                                        if no_of_entries == 0:
                                            break
                            else:
                                txt_msg =""
                                k=1
                                while (no_of_entries): 
                                    for r_value in j:
                                        txt_msg =txt_msg+'\n'+'\n'+str(k)+']Company Name:\n'+str(worksheet2.cell(r_value,4).value)+'\nItem Name:\n'+str(worksheet2.cell(r_value,9).value)+'\nQTY UOM.\n'+str(worksheet2.cell(r_value,10).value)+str(worksheet2.cell(r_value,11).value)
                                        if (worksheet2.cell(r_value,13).value)!="":
                                            txt_msg=txt_msg + '\nItem Name 2:\n'+str(worksheet2.cell(r_value,13).value)+'\nQTY UOM.\n'+str(worksheet2.cell(r_value,14).value)+str(worksheet2.cell(r_value,15).value)
                                        txt_msg = txt_msg +'\nVEHICAL NO:\n'+str(worksheet2.cell(r_value,23).value)+'\nDRIVER NAME'+str(worksheet2.cell(r_value,24).value)+'\nDRIVER MOBILE NO'+str(worksheet2.cell(r_value,25).value)+'\nExpected date and time of delivery:'+str(worksheet2.cell(r_value,8).value)+'\nTo:'+str(worksheet2.cell(r_value,7).value)+'\n+|+|+|+|  +|+|+|+|  +|+|+|+|  +|+|+|+|'
                                        sleep(1)
                                        bot.sendMessage(chat_id,"Generating pdf.\nPlease Wait")
                                        k=k+1
                                        (no_of_entries)=(no_of_entries)-1
                                        if no_of_entries == 0:
                                            break
                                    
                                with open("test.txt",'w') as f:
                                    f.write(txt_msg)
                                    f.close()
                                bot.sendDocument(chat_id,document=open('test.txt'))                    
                    elif command[2]=='2':
                        worksheet2 = sh.worksheet('P2:01.Material Inward cum QC (JF31)')
                        location_list = worksheet2.col_values(4)
                        i=1
                        j=[]
                        for location in location_list:
                            location = location.replace(" ","")
                            if unit.lower() in location.lower():
                               j.append(i) 
                            i=i+1
                        
                        if not(j):
                            bot.sendMessage(chat_id,'There is no data for this location')
                        if(len(command)==3):
                            if len(j)<4:
                                for r_value in j:
                                    txt_msg = 'Party Name (Received From Party): '+(str(worksheet2.cell(r_value,5).value))+'\n1]Item Description: '+(str(worksheet2.cell(r_value,12).value))+'\nQuantity in Nos: '+(str(worksheet2.cell(r_value,13).value))                                
                                    if(str(worksheet2.cell(r_value,16).value)):
                                        txt_msg = txt_msg + '\n2]Item Description: '+(str(worksheet2.cell(r_value,16).value))+'\nQuantity in Nos: '+(str(worksheet2.cell(r_value,17).value))
                                        if(str(worksheet2.cell(r_value,20).value)):
                                            txt_msg = txt_msg + '\n3]Item Description: '+(str(worksheet2.cell(r_value,20).value))+'\nQuantity in Nos: '+(str(worksheet2.cell(r_value,21).value))
                                    txt_msg=txt_msg + '\nVehicle No: '+(str(worksheet2.cell(r_value,6).value))+'\nINVOICE >> Invoice No and Date: '+(str(worksheet2.cell(r_value,10).value))+'\nSITE: '+(str(worksheet2.cell(r_value,4).value))
                                    bot.sendMessage(chat_id,txt_msg)
                                    
                            elif len(j)>=4:
                                txt_msg =""
                                k=1
                                for r_value in j:                                 
                                    txt_msg = txt_msg+str(k)+'] Party Name (Received From Party): '+(str(worksheet2.cell(r_value,5).value))+'\n1]Item Description: '+(str(worksheet2.cell(r_value,12).value))+'\nQuantity in Nos: '+(str(worksheet2.cell(r_value,13).value))                                
                                    if(str(worksheet2.cell(r_value,16).value)):
                                        txt_msg = txt_msg + '\n2]Item Description: '+(str(worksheet2.cell(r_value,16).value))+'\nQuantity in Nos: '+(str(worksheet2.cell(r_value,17).value))
                                        if(str(worksheet2.cell(r_value,20).value)):
                                            txt_msg = txt_msg + '\n3]Item Description: '+(str(worksheet2.cell(r_value,20).value))+'\nQuantity in Nos: '+(str(worksheet2.cell(r_value,21).value))
                                    txt_msg=txt_msg + '\nVehicle No: '+(str(worksheet2.cell(r_value,6).value))+'\nINVOICE >> Invoice No and Date: '+(str(worksheet2.cell(r_value,10).value))+'\nSITE: '+(str(worksheet2.cell(r_value,4).value))+'\n\n+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-\n'
                                    k=k+1
                                with open("test.doc",'w') as f:
                                    f.write(txt_msg)
                                    f.close()
                                bot.sendDocument(chat_id,document=open('test.doc'))
                        elif(len(command)==4):
                            no_of_entries=command[3]
                            print(no_of_entries)
                            if int(no_of_entries) < 4:
                                print(len(j))
                                if len(j) > no_of_entries:
                                    while no_of_entries:
                                        for r_value in j:
                                            txt_msg = 'Party Name (Received From Party): '+(str(worksheet2.cell(r_value,5).value))+'\n1]Item Description: '+(str(worksheet2.cell(r_value,12).value))+'\nQuantity in Nos: '+(str(worksheet2.cell(r_value,13).value))                                
                                            if(str(worksheet2.cell(r_value,16).value)):
                                                txt_msg = txt_msg + '\n2]Item Description: '+(str(worksheet2.cell(r_value,16).value))+'\nQuantity in Nos: '+(str(worksheet2.cell(r_value,17).value))
                                                if(str(worksheet2.cell(r_value,20).value)):
                                                    txt_msg = txt_msg + '\n3]Item Description: '+(str(worksheet2.cell(r_value,20).value))+'\nQuantity in Nos: '+(str(worksheet2.cell(r_value,21).value))
                                            txt_msg = txt_msg + '\nVehicle No: '+(str(worksheet2.cell(r_value,6).value))+'\nINVOICE >> Invoice No and Date: '+(str(worksheet2.cell(r_value,10).value))+'\nSITE: '+(str(worksheet2.cell(r_value,4).value))
                                            bot.sendMessage(chat_id,txt_msg)
                                            no_of_entries = no_of_entries - 1                                
                                else:
                                    for r_value in j:
                                        txt_msg = 'Party Name (Received From Party): '+(str(worksheet2.cell(r_value,5).value))+'\n1]Item Description: '+(str(worksheet2.cell(r_value,12).value))+'\nQuantity in Nos: '+(str(worksheet2.cell(r_value,13).value))                                
                                        if(str(worksheet2.cell(r_value,16).value)):
                                            txt_msg = txt_msg + '\n2]Item Description: '+(str(worksheet2.cell(r_value,16).value))+'\nQuantity in Nos: '+(str(worksheet2.cell(r_value,17).value))
                                            if(str(worksheet2.cell(r_value,20).value)):
                                                txt_msg = txt_msg + '\n3]Item Description: '+(str(worksheet2.cell(r_value,20).value))+'\nQuantity in Nos: '+(str(worksheet2.cell(r_value,21).value))
                                        txt_msg=txt_msg + '\nVehicle No: '+(str(worksheet2.cell(r_value,6).value))+'\nINVOICE >> Invoice No and Date: '+(str(worksheet2.cell(r_value,10).value))+'\nSITE: '+(str(worksheet2.cell(r_value,4).value))
                                        bot.sendMessage(chat_id,txt_msg)
                            elif no_of_entries >= 4:
                                print('this?')
                                if len(j) > no_of_entries:
                                    while no_of_entries:
                                        for r_value in j:
                                            txt_msg = 'Party Name (Received From Party): '+(str(worksheet2.cell(r_value,5).value))+'\n1]Item Description: '+(str(worksheet2.cell(r_value,12).value))+'\nQuantity in Nos: '+(str(worksheet2.cell(r_value,13).value))
                                            if(str(worksheet2.cell(r_value,16).value)):
                                                txt_msg = txt_msg + '\n2]Item Description: '+(str(worksheet2.cell(r_value,16).value))+'\nQuantity in Nos: '+(str(worksheet2.cell(r_value,17).value))
                                                if(str(worksheet2.cell(r_value,20).value)):
                                                        txt_msg = txt_msg + '\n3]Item Description: '+(str(worksheet2.cell(r_value,20).value))+'\nQuantity in Nos: '+(str(worksheet2.cell(r_value,21).value))
                                                txt_msg = txt_msg + '\nVehicle No: '+(str(worksheet2.cell(r_value,6).value))+'\nINVOICE >> Invoice No and Date: '+(str(worksheet2.cell(r_value,10).value))+'\nSITE: '+(str(worksheet2.cell(r_value,4).value))
                                                no_of_entries = no_of_entries - 1
                                    with open("test.doc",'w') as f:
                                        f.write(txt_msg)
                                        f.close()
                                    bot.sendDocument(chat_id,document=open('test.doc'))
                                else:
                                    for r_value in j:
                                        txt_msg = 'Party Name (Received From Party): '+(str(worksheet2.cell(r_value,5).value))+'\n1]Item Description: '+(str(worksheet2.cell(r_value,12).value))+'\nQuantity in Nos: '+(str(worksheet2.cell(r_value,13).value))
                                        if(str(worksheet2.cell(r_value,16).value)):
                                            txt_msg = txt_msg + '\n2]Item Description: '+(str(worksheet2.cell(r_value,16).value))+'\nQuantity in Nos: '+(str(worksheet2.cell(r_value,17).value))
                                            if(str(worksheet2.cell(r_value,20).value)):
                                                    txt_msg = txt_msg + '\n3]Item Description: '+(str(worksheet2.cell(r_value,20).value))+'\nQuantity in Nos: '+(str(worksheet2.cell(r_value,21).value))
                                            txt_msg = txt_msg + '\nVehicle No: '+(str(worksheet2.cell(r_value,6).value))+'\nINVOICE >> Invoice No and Date: '+(str(worksheet2.cell(r_value,10).value))+'\nSITE: '+(str(worksheet2.cell(r_value,4).value))
                                            no_of_entries = no_of_entries - 1
                                    with open("test.doc",'w') as f:
                                        f.write(txt_msg)
                                        f.close()
                                    bot.sendDocument(chat_id,document=open('test.doc'))

                        else:
                            bot.sendMessage(chat_id,'Invalid usage of P2 command. \nLength of this command can be either 3 or 4.\n')                    
                    elif command[2]=='3':
                        print('P3')
                    elif command[2]=='4':
                        print('P4')
                    elif command[2]=='5':
                        print('P5')
                    else:
                        bot.sendMessage(chat_id,'Invalid command for P.')
                else:
                    bot.sendMessage(chat_id,'Sorry!\nYou are not authorised to use this command.')
            else:
                bot.sendMessage(chat_id,'Sorry!\nYou are not registered.')               
        elif str(command.lower()) == 's1':
            now = datetime.datetime.now()
            todays_date = now.strftime("%d/%m/%Y")
            worksheet = sh.worksheet('Telegram ID')
            telegram_id_list = worksheet.col_values(5)
            if str(chat_id) in telegram_id_list:
                row_value = int(telegram_id_list.index(str(chat_id)))+1
                if 'SUP' in str(worksheet.cell(row_value,6).value):
                    worksheet2 = sh.worksheet('S1- Attendence sheeet')
                    date_list = worksheet2.col_values(7)
                    i = 1
                    for date in date_list:
                        if str(date) == todays_date:
                            in_time = str(worksheet2.cell(i,8).value)
                            in_time = in_time.rstrip()
                            in_time = in_time.strip()
                            if len(in_time)==5:
                                if int(in_time[:2]) >= 10:
                                    try:
                                        if int(in_time[:2]) == 10:
                                            bot.sendMessage(int(worksheet2.cell(i,9).value),("You were late today by, "+ str(in_time[-2:])+" minutes."))
                                        else:
                                            bot.sendMessage(int(worksheet2.cell(i,9).value),("You were late today by, "+str(int(in_time[:2])-10)+' hours & '+ str(in_time[-2:])+" minutes."))
                                    except:
                                        print('There is no Telegram Id for '+str(worksheet2.cell(i,4).value))
                                    time.sleep(1)                            
                        i = i + 1
            bot.sendMessage(chat_id,"Message sent to all latecomers for today who's telgram id\'s are available.")
        elif str(command[1]).lower()=='b':
            worksheet = sh.worksheet('Telegram ID')
            telegram_id_list = worksheet.col_values(5)
            if str(chat_id) in telegram_id_list:
                row_value = int(telegram_id_list.index(str(chat_id)))+1
                if 'SUP' in str(worksheet.cell(row_value,6).value):
                    if str(command[2]).lower()=='1':
                        worksheet2 = sh.worksheet('B1: PRODUCTION SHEET')
                        date_list = worksheet2.col_values(7)
                        i=1
                        j=[]
                        for any_date in date_list:
                            if str(any_date)==str(dt_string[:10]):
                                j.append(i)
                            i = i + 1
                        if not j:
                            bot.sendMessage(chat_id,'No data for today!')
                        else:
                            for any_value in j:
                                if str(command[0]).lower()=='2':
                                    if str(worksheet2.cell(any_value,1).value)=='02 - IDEAL':
                                        txt_msg = "CONTROL NO:"+str(worksheet2.cell(any_value,2).value)+",\nTotal Box Packed:"
                                        if(str(worksheet2.cell(any_value,35).value) !=('0'or" ")):
                                            txt_msg = txt_msg +"\n4.5kg:"+str(worksheet2.cell(any_value,35).value)
                                        if(str(worksheet2.cell(any_value,36).value) !=('0'or" ")):
                                            txt_msg = txt_msg +"\n5 kg Boxes - L:"+str(worksheet2.cell(any_value,36).value)
                                        if(str(worksheet2.cell(any_value,37).value) !=('0'or" ")):
                                            txt_msg = txt_msg +"\n5 kg Boxes - XL:"+str(worksheet2.cell(any_value,37).value)
                                        if(str(worksheet2.cell(any_value,38).value) !=('0'or" ")):
                                            txt_msg = txt_msg +"\n8.2 kg Boxes:"+str(worksheet2.cell(any_value,38).value)
                                        txt_msg = txt_msg +"\nBerry %:"+str(worksheet2.cell(any_value,33).value)
                                        bot.sendMessage(chat_id,txt_msg)

                                elif str(command[0]).lower()=='3':
                                    if str(worksheet2.cell(any_value,1).value)=='03 - AMAYA':
                                        txt_msg = "CONTROL NO:"+str(worksheet2.cell(any_value,2).value)+",\nTotal Box Packed:"
                                        if(str(worksheet2.cell(any_value,35).value) !=('0'or" ")):
                                            txt_msg = txt_msg +"\n4.5kg:"+str(worksheet2.cell(any_value,35).value)
                                        if(str(worksheet2.cell(any_value,36).value) !=('0'or" ")):
                                            txt_msg = txt_msg +"\n5 kg Boxes - L:"+str(worksheet2.cell(any_value,36).value)
                                        if(str(worksheet2.cell(any_value,37).value) !=('0'or" ")):
                                            txt_msg = txt_msg +"\n5 kg Boxes - XL:"+str(worksheet2.cell(any_value,37).value)
                                        if(str(worksheet2.cell(any_value,38).value) !=('0'or" ")):
                                            txt_msg = txt_msg +"\n8.2 kg Boxes:"+str(worksheet2.cell(any_value,38).value)
                                        txt_msg = txt_msg +"\nBerry %:"+str(worksheet2.cell(any_value,33).value)
                                        bot.sendMessage(chat_id,txt_msg)

                                elif str(command[0]).lower()=='4':
                                    if str(worksheet2.cell(any_value,1).value)=='04 - MRJ':
                                        txt_msg = "CONTROL NO:"+str(worksheet2.cell(any_value,2).value)+",\nTotal Box Packed:"
                                        if(str(worksheet2.cell(any_value,35).value) !=('0'or" ")):
                                            txt_msg = txt_msg +"\n4.5kg:"+str(worksheet2.cell(any_value,35).value)
                                        if(str(worksheet2.cell(any_value,36).value) !=('0'or" ")):
                                            txt_msg = txt_msg +"\n5 kg Boxes - L:"+str(worksheet2.cell(any_value,36).value)
                                        if(str(worksheet2.cell(any_value,37).value) !=('0'or" ")):
                                            txt_msg = txt_msg +"\n5 kg Boxes - XL:"+str(worksheet2.cell(any_value,37).value)
                                        if(str(worksheet2.cell(any_value,38).value) !=('0'or" ")):
                                            txt_msg = txt_msg +"\n8.2 kg Boxes:"+str(worksheet2.cell(any_value,38).value)
                                        txt_msg = txt_msg +"\nBerry %:"+str(worksheet2.cell(any_value,33).value)
                                        bot.sendMessage(chat_id,txt_msg)
                                elif str(command[0]).lower()=='5':
                                    txt_msg = "CONTROL NO:"+str(worksheet2.cell(any_value,2).value)+",\nTotal Box Packed:"
                                    if(str(worksheet2.cell(any_value,35).value) !=('0'or" ")):
                                        txt_msg = txt_msg +"\n4.5kg:"+str(worksheet2.cell(any_value,35).value)
                                    if(str(worksheet2.cell(any_value,36).value) !=('0'or" ")):
                                        txt_msg = txt_msg +"\n5 kg Boxes - L:"+str(worksheet2.cell(any_value,36).value)
                                    if(str(worksheet2.cell(any_value,37).value) !=('0'or" ")):
                                        txt_msg = txt_msg +"\n5 kg Boxes - XL:"+str(worksheet2.cell(any_value,37).value)
                                    if(str(worksheet2.cell(any_value,38).value) !=('0'or" ")):
                                        txt_msg = txt_msg +"\n8.2 kg Boxes:"+str(worksheet2.cell(any_value,38).value)
                                    txt_msg = txt_msg +"\nBerry %:"+str(worksheet2.cell(any_value,33).value)
                                    bot.sendMessage(chat_id,txt_msg)
            
        elif str(command[1]).lower()=='c':
            worksheet = sh.worksheet('Telegram ID')
            telegram_id_list = worksheet.col_values(5)
            if str(chat_id) in telegram_id_list:
                row_value = int(telegram_id_list.index(str(chat_id)))+1
                if 'SUP' in str(worksheet.cell(row_value,6).value):
                    if str(command[2]).lower()=='1':
                        worksheet2 = sh.worksheet('C1- AF09 box making  report')
                        unit_list = worksheet2.col_values(4)
                        i=1
                        amaya=[]
                        miraj=[]
                        ideal=[]
                        for any_unit in unit_list:
                            if str(any_unit) == str("04-MRJ"):
                                miraj.append(i)
                            elif str(any_unit)==str("03 - AMAYA"):
                                amaya.append(i)
                            elif  str(any_unit)==str("02 - IDEAL"):
                                ideal.append(i)
                            i = i + 1
                        if str(command[0])=='2':
                            for k in range(-1,-4,-1):
                                txt_msg = "Date :"+str(worksheet2.cell(ideal[k],2).value)
                                if(str(worksheet2.cell(ideal[k],6).value) !=('0'or" ")):
                                    txt_msg = txt_msg + "\n5 kg BOX,COLOR:"+(str(worksheet2.cell(ideal[k],6).value))
                                    txt_msg = txt_msg + "\nTOTAL 5 kg BOX MADE:"+(str(worksheet2.cell(ideal[k],7).value))
                                if(str(worksheet2.cell(ideal[k],11).value) !=('0'or" ")):
                                    txt_msg = txt_msg + "\n8 kg BOX,COLOR:"+(str(worksheet2.cell(ideal[k],11).value))
                                    txt_msg = txt_msg + "\nTOTAL 8 kg BOX MADE:"+(str(worksheet2.cell(ideal[k],12).value))
                                if(str(worksheet2.cell(ideal[k],16).value) !=('0'or" ")):
                                    txt_msg = txt_msg + "\n4.5 kg BOX,COLOR:"+(str(worksheet2.cell(ideal[k],16).value))
                                    txt_msg = txt_msg + "\nTOTAL 4.5 kg BOX MADE:"+(str(worksheet2.cell(ideal[k],17).value))
                                bot.sendMessage(chat_id,txt_msg)

                        elif str(command[0])=='3':
                            for k in range(-1,-4,-1):
                                txt_msg = "Date :"+str(worksheet2.cell(amaya[k],2).value)
                                if(str(worksheet2.cell(amaya[k],6).value) !=('0'or" ")):
                                    txt_msg = txt_msg + "\n5 kg BOX,COLOR:"+(str(worksheet2.cell(amaya[k],6).value))
                                    txt_msg = txt_msg + "\nTOTAL 5 kg BOX MADE:"+(str(worksheet2.cell(amaya[k],7).value))
                                if(str(worksheet2.cell(amaya[k],11).value) !=('0'or" ")):
                                    txt_msg = txt_msg + "\n8 kg BOX,COLOR:"+(str(worksheet2.cell(amaya[k],11).value))
                                    txt_msg = txt_msg + "\nTOTAL 8 kg BOX MADE:"+(str(worksheet2.cell(amaya[k],12).value))
                                if(str(worksheet2.cell(amaya[k],16).value) !=('0'or" ")):
                                    txt_msg = txt_msg + "\n4.5 kg BOX,COLOR:"+(str(worksheet2.cell(amaya[k],16).value))
                                    txt_msg = txt_msg + "\nTOTAL 4.5 kg BOX MADE:"+(str(worksheet2.cell(amaya[k],17).value))
                                bot.sendMessage(chat_id,txt_msg)
                        elif str(command[0])=='4':
                            for k in range(-1,-4,-1):
                                txt_msg = "Date :"+str(worksheet2.cell(miraj[k],2).value)
                                if(str(worksheet2.cell(miraj[k],6).value) !=('0'or" ")):
                                    txt_msg = txt_msg + "\n5 kg BOX,COLOR:"+(str(worksheet2.cell(miraj[k],6).value))
                                    txt_msg = txt_msg + "\nTOTAL 5 kg BOX MADE:"+(str(worksheet2.cell(miraj[k],7).value))
                                if(str(worksheet2.cell(miraj[k],11).value) !=('0'or" ")):
                                    txt_msg = txt_msg + "\n8 kg BOX,COLOR:"+(str(worksheet2.cell(miraj[k],11).value))
                                    txt_msg = txt_msg + "\nTOTAL 8 kg BOX MADE:"+(str(worksheet2.cell(miraj[k],12).value))
                                if(str(worksheet2.cell(miraj[k],16).value) !=('0'or" ")):
                                    txt_msg = txt_msg + "\n4.5 kg BOX,COLOR:"+(str(worksheet2.cell(miraj[k],16).value))
                                    txt_msg = txt_msg + "\nTOTAL 4.5 kg BOX MADE:"+(str(worksheet2.cell(miraj[k],17).value))

                                bot.sendMessage(chat_id,txt_msg)
                        elif str(command[0]) == '5':
                            for k in range(-1,-4,-1):
                                txt_msg = "Date :"+str(worksheet2.cell(miraj[k],2).value)
                                if(str(worksheet2.cell(miraj[k],6).value) !=('0'or" ")):
                                    txt_msg = txt_msg + "\n5 kg BOX,COLOR:"+(str(worksheet2.cell(miraj[k],6).value))
                                    txt_msg = txt_msg + "\nTOTAL 5 kg BOX MADE:"+(str(worksheet2.cell(miraj[k],7).value))
                                if(str(worksheet2.cell(miraj[k],11).value) !=('0'or" ")):
                                    txt_msg = txt_msg + "\n8 kg BOX,COLOR:"+(str(worksheet2.cell(miraj[k],11).value))
                                    txt_msg = txt_msg + "\nTOTAL 8 kg BOX MADE:"+(str(worksheet2.cell(miraj[k],12).value))
                                if(str(worksheet2.cell(miraj[k],16).value) !=('0'or" ")):
                                    txt_msg = txt_msg + "\n4.5 kg BOX,COLOR:"+(str(worksheet2.cell(miraj[k],16).value))
                                    txt_msg = txt_msg + "\nTOTAL 4.5 kg BOX MADE:"+(str(worksheet2.cell(miraj[k],17).value))
                            bot.sendMessage(chat_id,txt_msg)
                            for k in range(-1,-4,-1):
                                txt_msg = "Date :"+str(worksheet2.cell(amaya[k],2).value)
                                if(str(worksheet2.cell(amaya[k],6).value) !=('0'or" ")):
                                    txt_msg = txt_msg + "\n5 kg BOX,COLOR:"+(str(worksheet2.cell(amaya[k],6).value))
                                    txt_msg = txt_msg + "\nTOTAL 5 kg BOX MADE:"+(str(worksheet2.cell(amaya[k],7).value))
                                if(str(worksheet2.cell(amaya[k],11).value) !=('0'or" ")):
                                    txt_msg = txt_msg + "\n8 kg BOX,COLOR:"+(str(worksheet2.cell(amaya[k],11).value))
                                    txt_msg = txt_msg + "\nTOTAL 8 kg BOX MADE:"+(str(worksheet2.cell(amaya[k],12).value))
                                if(str(worksheet2.cell(amaya[k],16).value) !=('0'or" ")):
                                    txt_msg = txt_msg + "\n4.5 kg BOX,COLOR:"+(str(worksheet2.cell(amaya[k],16).value))
                                    txt_msg = txt_msg + "\nTOTAL 4.5 kg BOX MADE:"+(str(worksheet2.cell(amaya[k],17).value))
                            bot.sendMessage(chat_id,txt_msg)
                            for k in range(-1,-4,-1):
                                txt_msg = "Date :"+str(worksheet2.cell(ideal[k],2).value)
                                if(str(worksheet2.cell(ideal[k],6).value) !=('0'or" ")):
                                    txt_msg = txt_msg + "\n5 kg BOX,COLOR:"+(str(worksheet2.cell(ideal[k],6).value))
                                    txt_msg = txt_msg + "\nTOTAL 5 kg BOX MADE:"+(str(worksheet2.cell(ideal[k],7).value))
                                if(str(worksheet2.cell(ideal[k],11).value) !=('0'or" ")):
                                    txt_msg = txt_msg + "\n8 kg BOX,COLOR:"+(str(worksheet2.cell(ideal[k],11).value))
                                    txt_msg = txt_msg + "\nTOTAL 8 kg BOX MADE:"+(str(worksheet2.cell(ideal[k],12).value))
                                if(str(worksheet2.cell(ideal[k],16).value) !=('0'or" ")):
                                    txt_msg = txt_msg + "\n4.5 kg BOX,COLOR:"+(str(worksheet2.cell(ideal[k],16).value))
                                    txt_msg = txt_msg + "\nTOTAL 4.5 kg BOX MADE:"+(str(worksheet2.cell(ideal[k],17).value))
                                bot.sendMessage(chat_id,txt_msg)                        
                        else:
                            bot.sendMessage(chat_id,"Please enter 2c/3c/4c/5c only!")
                        
        elif str(command[0]).lower()=='v':
            if str(command[1]) == '1':
                worksheet = sh.worksheet('Telegram ID')
                telegram_id_list = worksheet.col_values(5)
                if str(chat_id) in telegram_id_list:
                    row_value = int(telegram_id_list.index(str(chat_id)))+1
                    if 'SUP' in str(worksheet.cell(row_value,6).value):
                        worksheet2 = sh.worksheet('V1-GF04Nasik Shuttle')
                        last_row = len(worksheet2.col_values(2))
                        for i in range(0,-3,-1):
                            try:
                                row_value = last_row+i
                                txt_msg = "Date :"+str(worksheet2.cell(row_value,2).value)
                                txt_msg = txt_msg +"\nVehicle Number: "+str(worksheet2.cell(row_value,3).value)
                                txt_msg = txt_msg +"\nDriver Name: "+str(worksheet2.cell(row_value,4).value)
                                txt_msg = txt_msg +"\nRoute: "+str(worksheet2.cell(row_value,5).value)
                                txt_msg = txt_msg +"\Material: "+str(worksheet2.cell(row_value,6).value)
                                txt_msg = txt_msg +"\nApprox. Qty.: "+str(worksheet2.cell(row_value,7).value)
                                bot.sendMessage(chat_id,txt_msg)
                            except:
                                print("There was some problem with the rows.")
                        
        elif str(command[:2]).lower()=='jf':
            worksheet = sh.worksheet('JF')
            if len(command) == 2:
                bot.sendMessage(chat_id,'Following are the codes for different groups. Please reply with four digit code to get related forms.')
                txt_msg = str(worksheet.cell(3,2).value)
                bot.sendMessage(chat_id,txt_msg)
            elif len(command) == 4:
                jf_code_list = worksheet.col_values(1)
                keyword = command[2:]
                if keyword.lower() == 'of':
                    j = 1
                    k = []
                    for val in jf_code_list:
                        if 'of' in val.lower():
                            k.append(j)
                        j = j + 1
                    for val in k:
                        txt_msg = str((worksheet.cell(val,2).value).decode('utf-8'))+' : '+str((worksheet.cell(val,3).value).decode('utf-8'))
                        bot.sendMessage(chat_id,txt_msg)
                elif keyword.lower() == 'af':
                    j = 1
                    k = []
                    for val in jf_code_list:
                        if 'af' in val.lower():
                            k.append(j)
                        j = j + 1
                    for val in k:
                        txt_msg = str((worksheet.cell(val,2).value).decode('utf-8'))+' : '+str((worksheet.cell(val,3).value).decode('utf-8'))
                        bot.sendMessage(chat_id,txt_msg)
                elif keyword.lower() == 'ws':
                    j = 1
                    k = []
                    for val in jf_code_list:
                        if ('ws' in val.lower())or('jf13'==val.lower()):
                            k.append(j)
                        j = j + 1
                    for val in k:
                        txt_msg = str((worksheet.cell(val,2).value).decode('utf-8'))+' : '+str((worksheet.cell(val,3).value).decode('utf-8'))
                        bot.sendMessage(chat_id,txt_msg)
                elif keyword.lower() == 'gf':
                    j = 1
                    k = []
                    for val in jf_code_list:
                        if 'gf' in val.lower():
                            k.append(j)
                        j = j + 1
                    for val in k:
                        txt_msg = str((worksheet.cell(val,2).value).decode('utf-8'))+' : '+str((worksheet.cell(val,3).value).decode('utf-8'))
                        bot.sendMessage(chat_id,txt_msg)
                elif keyword.lower() == 'ff':
                    j = 1
                    k = []
                    for val in jf_code_list:
                        if 'ff' in val.lower():
                            k.append(j)
                        j = j + 1
                    for val in k:
                        txt_msg = str((worksheet.cell(val,2).value).decode('utf-8'))+' : '+str((worksheet.cell(val,3).value).decode('utf-8'))
                        bot.sendMessage(chat_id,txt_msg)
                elif keyword.lower() == 'hf':
                    j = 1
                    k = []
                    for val in jf_code_list:
                        if 'hf' in val.lower():
                            k.append(j)
                        j = j + 1
                    for val in k:
                        txt_msg = str((worksheet.cell(val,2).value).decode('utf-8'))+' : '+str((worksheet.cell(val,3).value).decode('utf-8'))
                        bot.sendMessage(chat_id,txt_msg)
                elif keyword.lower() == 'lf':
                    j = 1
                    k = []
                    for val in jf_code_list:
                        if 'lf' in val.lower():
                            k.append(j)
                        j = j + 1
                    for val in k:
                        txt_msg = str((worksheet.cell(val,2).value).decode('utf-8'))+' : '+str((worksheet.cell(val,3).value).decode('utf-8'))
                        bot.sendMessage(chat_id,txt_msg)
                elif keyword.lower() == 'qf':
                    j = 1
                    k = []
                    for val in jf_code_list:
                        if 'qf' in val.lower():
                            k.append(j)
                        j = j + 1
                    for val in k:
                        txt_msg = str((worksheet.cell(val,2).value).decode('utf-8'))+' : '+str((worksheet.cell(val,3).value).decode('utf-8'))
                        bot.sendMessage(chat_id,txt_msg)
                elif keyword.lower() == 'ro':
                    j = 1
                    k = []
                    for val in jf_code_list:
                        if ('rq' in val.lower())or('jf01'==val.lower()):
                            k.append(j)
                        j = j + 1
                    for val in k:
                        txt_msg = str((worksheet.cell(val,2).value).decode('utf-8'))+' : '+str((worksheet.cell(val,3).value).decode('utf-8'))
                        bot.sendMessage(chat_id,txt_msg)

        elif str(command).lower() == 's3,1':
            worksheet = sh.worksheet('Telegram ID')
            telegram_id_list = worksheet.col_values(5)
            if str(chat_id) in telegram_id_list:
                row_value = int(telegram_id_list.index(str(chat_id)))+1
                if 'SUP' in str(worksheet.cell(row_value,6).value):
                    worksheet = sh.worksheet('S3,1-Daily JF staff')
                    names_list = worksheet.col_values(6)
                    date_list = worksheet.col_values(2)
                    todays_now=str(now)
                    todays_date = str(todays_now[8:10])+'-'+str(todays_now[5:7])+'-'+str(todays_now[:4])
                    y_date = str(int(todays_now[8:10])-1)+'-'+str(todays_now[5:7])+'-'+str(todays_now[:4])
                    i = 1
                    j = []
                    for k in date_list:
                        if y_date in str(k):
                            j.append(i)
                        i = i + 1
                    present_list=[]
                    for l in j:
                        present_list.append(str((worksheet.cell(l,3).value).decode('utf-8')))
                    absent_list = Diff(names_list,present_list)
                    if absent_list:
                        print(absent_list)
                        for member in absent_list:
                            i=1
                            if member in  names_list:
                                try:
                                    bot.sendMessage(str((worksheet.cell(i,7).value).decode('utf-8')),'Please fill yesterdays Daily Report')
                                    bot.sendMessage(chat_id,str((worksheet.cell(i,6).value).decode('utf-8')))
                                except:
                                    temp='temp'
                            i = i + 1 
                    else:
                        bot.sendMessage(chat_id,'Everyone was present yesterday')
            else:
                bot.sendMessage(chat_id,'You are not authorised to use this command')
        elif 'q1'== command.lower()[:2]:
            if len(command) == 26:
                site_code = command[3]
                start_date = command[5:15]
                end_date = command[16:27]
                control_no_vs_berry(int(site_code),start_date,end_date,bot,chat_id)
                bot.sendDocument(chat_id,document=open('/home/pi/reports/Controlnovsberrypercentage.png')) 
            else:
                bot.sendMessage(chat_id,'You have put wrong format for Q1')
        elif 'showcolours'in command.lower():
            if (command[11:].isdigit() and int(command[11:])>0):
                bot.sendMessage(chat_id,'Please wait while we send color analysis pie chart')
                with open('/home/pi/RaisinsColorMonitor/status.txt', 'w') as f:
                    f.write('DONE')
                try:
                    with open('/home/pi/RaisinsColorMonitor/bands.txt', 'w') as f:
                        f.write(command[11:])
                    os.system('python /home/pi/RaisinsColorMonitor/opencv_approach.py')
                    with open('/home/pi/RaisinsColorMonitor/status.txt') as f:
                        lines = f.readlines()
                    if ('DONE'not in lines):
                        while("WIP" in lines):
                            bot.sendMessage(chat_id,'Please wait while we are working on finding Color Scheme')
                            sleep(5)
                            with open('/home/pi/RaisinsColorMonitor/status.txt') as s:
                                lines = s.readlines()
                    elif('DONE'in lines):
                        bot.sendDocument(chat_id,document=open('/home/pi/RaisinsColorMonitor/plot.png'))
                    with open('/home/pi/RaisinsColorMonitor/status.txt', 'w') as f:
                        f.write('DONE')
                except:
                    print("There was some error")
                    with open('/home/pi/RaisinsColorMonitor/status.txt', 'w') as f:
                        f.write('DONE')
            else:
                bot.sendMessage(chat_id,'Please send this command in following format:\n\nshowcolours5\n\n If you want to get top5 colours')                    
        else:
            bot.sendMessage(chat_id,'No such command')
    elif 'photo' in msg:
        #print(msg)
        try:
            if str(msg['from']['last_name']).decode('utf-8'):
                current_name = str(msg['from']['first_name'])+' '+str(msg['from']['last_name'])
            else:
                current_name = str(msg['from']['first_name'])
        except:
            if (msg['from']['last_name']).decode('utf-8')!='':
                current_name = str((msg['from']['first_name']).decode('utf-8'))+' '+str((msg['from']['last_name'].decode('utf-8')))
            else:
                current_name = str((msg['from']['first_name']).decode('utf-8'))
        #print(current_name)
        now = datetime.datetime.now()
        dt_string = now.strftime("%H_%M_%S_%d_%m_%Y")
        try:
            file_id=(msg['photo'][2]['file_id'])
        except:
            file_id=(msg['photo'][1]['file_id'])
        bot.download_file(file_id,'/home/pi/fresh_express_documents/'+dt_string+'_sent_by_'+current_name+'.jpg')
        im = Image.open('/home/pi/fresh_express_documents/'+dt_string+'_sent_by_'+current_name+'.jpg')
        width, height = im.size
        #if height>width:
        #    im = im.rotate(90)
        #    im.save('/home/pi/fresh_express_documents/'+dt_string+'.jpg')
        file_name = os.path.abspath('/home/pi/fresh_express_documents/'+dt_string+'_sent_by_'+current_name+'.jpg')
        with io.open(file_name, 'rb') as image_file:
            content = image_file.read()
        image = vision.types.Image(content=content)
        response = client.label_detection(image=image)
        labels = response.label_annotations
        #print("labels are:",labels)
        IMAGE_LABELS=[]
        for label in labels:
            #print(label.description)
            IMAGE_LABELS.append((label.description).encode("utf-8"))
        #print(IMAGE_LABELS)
        if (('Handwriting'in IMAGE_LABELS)or('Font'in IMAGE_LABELS)or('Paper'in IMAGE_LABELS)or('Document'in IMAGE_LABELS)or('Paper product'in IMAGE_LABELS)or('Label'in IMAGE_LABELS)):
            print("trying to detect text")
            try :
                detected_text = str(detectText('/home/pi/fresh_express_documents/'+dt_string+'_sent_by_'+current_name+'.jpg'))
                check_ocr(chat_id,detected_text,'/home/pi/fresh_express_documents/'+dt_string+'_sent_by_'+current_name+'.jpg')
            except :
                bot.sendMessage(chat_id,"OCR IS NOT POSSIBLE FOR THIS IMAGE")
        elif (("Plant"in IMAGE_LABELS)or("Grape leaves"in IMAGE_LABELS)or("Flowering plant"in IMAGE_LABELS)or("Deciduous"in IMAGE_LABELS)):
            try :
                msgtosend = findleafarea('/home/pi/fresh_express_documents/'+dt_string+'_sent_by_'+current_name+'.jpg')
                print(msgtosend)
                bot.sendMessage(chat_id,msgtosend)
                bot.sendDocument(chat_id,document=open('/home/pi/fresh_express_documents/'+dt_string+'_sent_by_'+current_name+'.jpg'))
                os.remove('/home/pi/fresh_express_documents/'+dt_string+'_sent_by_'+current_name+'.jpg')
            except:
                temp = "temp"
                print("there was some error")
        else:
            print("lable not in IMAGE_LABELS on line 1970")
    elif 'document' in msg:
        try:
            if str(msg['from']['last_name']).decode('utf-8'):
                current_name = str(msg['from']['first_name'])+' '+str(msg['from']['last_name'])
            else:
                current_name = str(msg['from']['first_name'])
        except:
            if (msg['from']['last_name']).decode('utf-8')!='':
                current_name = str((msg['from']['first_name']).decode('utf-8'))+' '+str((msg['from']['last_name'].decode('utf-8')))
            else:
                current_name = str((msg['from']['first_name']).decode('utf-8'))
        now = datetime.datetime.now()
        dt_string = now.strftime("%H_%M_%S_%d_%m_%Y")
        file_id=(msg['document']['file_id'])
        if '.jpg' or '.jpeg' or '.png' in str(msg['document']['file_name']):
            
            file_location_name = '/home/pi/fresh_express_documents/images/'+dt_string+'_sent_by_'+current_name+str(msg['document']['file_name'])
            bot.download_file(file_id,file_location_name)
            im = Image.open(file_location_name)
            width, height = im.size
            #if height>width:
            #    im = im.rotate(90)
            #    im.save(file_location_name)
            file_name = os.path.abspath(file_location_name)
            with io.open(file_name, 'rb') as image_file:
                content = image_file.read()
            image = vision.types.Image(content=content)
            response = client.label_detection(image=image)
            labels = response.label_annotations
            #print("labels are:",labels)
            IMAGE_LABELS=[]
            print('Labels:')
            for label in labels:
                #print(label.description)
                IMAGE_LABELS.append(label.description)
            if(('Handwriting'in IMAGE_LABELS)or('Font'in IMAGE_LABELS)or('Paper'in IMAGE_LABELS)or('Document'in IMAGE_LABELS)or('Paper product'in IMAGE_LABELS)or('Label'in IMAGE_LABELS)):
                try :
                    detected_text=str(detectText(file_location_name))
                    check_ocr(chat_id,detected_text,file_location_name)
                except :
                    bot.sendMessage(chat_id,"OCR IS NOT POSSIBLE FOR THIS IMAGE")
            elif(("Plant"in IMAGE_LABELS)or("Grape leaves"in IMAGE_LABELS)or("Flowering plant"in IMAGE_LABELS)or("Deciduous"in IMAGE_LABELS)):
                try :
                    msgtosend = findleafarea('/home/pi/fresh_express_documents/'+dt_string+'_sent_by_'+current_name+'.jpg')
                    bot.sendMessage(chat_id,msgtosend)
                except:
                    temp = "temp"
        else:
            bot.download_file(file_id,'/home/pi/fresh_express_documents/'+dt_string+str(msg['document']['file_name']))
        bot.sendMessage(chat_id,'Received '+str(msg['document']['file_name'])+'!')
    
    #except :#APIError:
    #    bot.sendMessage(chat_id,'There was some problem.\n Please try again')
            

try:
    bot.message_loop(handle)#intitate message loop and pass the handle function as a argument. This means that
except KeyError:
    print('There was Key error!')
print('Fresh Express bot is ready!')
now = datetime.datetime.now()
s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
s.connect(("8.8.8.8", 80))
start_msg='Fresh Express Telegram Bot started at '+ str(now)+ '\n with IP address :'+str(s.getsockname()[0])
s.close()

main_chat_id=[414553391,1210750385,1378878389,1043828479,1008930089,1383674791,1304264905,797890475]
for ids in main_chat_id:
    try:
        bot.sendMessage(ids,start_msg)
    except Exception as e:
        print(ids)
        print(e)
temp=0
while True:
    try:
        sleep(1)
        temp = temp + 1
        if temp == 150:
            bot.sendMessage(414553391,"You are the bestest ever!")
            temp = 0
    except KeyboardInterrupt:
        GPIO.cleanup()
        exit()
    except:
        print('Other error or exception occured!')
