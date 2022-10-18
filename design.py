#Get a patients drug data
#grab chemical name and build a library to store chemical name and merchant name and indiction and pharmalogical classification. 
##Resort the data according to the pharmacologic classification 
###Print out tree mapping drugs according to classification and date 
#### Put indications for a drug behind the drug name
print("reference: patient_list 1.5 mrd")
print("Author: M117 H.R.T")
print("modified: M118 T.Y.H")
import bs4
import requests
import pandas as pd
#import os
import re
import math
##########################################################################
import random
import time
import json
doc_num='DOC98021'
password_num="DOC98021"
global login_session
login_session=requests.Session()

global section_num
global vs_num
global ward_num
global bed_num
global chart_num
section_num = input("科別:")
vs_num = input("主治DOC:")
ward_num = input("病房:")
bed_num = ""
chart_num=''


def get_oc(address):
# 按照網頁的請求構造請求頭
    headers = {
            #'User-Agent': user_agent.random,
            'Connection': 'keep-alive',
            #'Referer': referlink,
            'Origin': 'http://f5-ws5.ndmctsgh.edu.tw',
            'Referer': 'http://f5-ws5.ndmctsgh.edu.tw/eForm/Account/Login'
        }



    data = {
        'login_id': doc_num,
        'password': password_num
    }
    # 按照服務端的需要構造請求，並獲得session，記錄在login_session
    response = login_session.post(address, data=data, verify=False, headers=headers)
    return(response)


def get_pt_oc():
    ptsearchdata = {
            'SearchChartno': chart_num,
            'SearchChinaname': "",
            'SearchIDNo': '', #身分證號
            'SearchSectionNo': section_num,
            'SearchNRCode': ward_num,
            'SearchBedCode': bed_num,
            'SearchVSDR': vs_num,
            'SearchSpecialNote': '',
            'SearchCcGroup[]': ''
    }
# 按照網頁的請求構造請求頭
    headers = {
            #'User-Agent': user_agent.random,
            'Connection': 'keep-alive',
            #'Referer': referlink,
            'Origin': 'http://f5-ws5.ndmctsgh.edu.tw',
            'Referer': 'http://f5-ws5.ndmctsgh.edu.tw/eForm/Account/Login'
        }

    # 按照服務端的需要構造請求，並獲得session，記錄在login_session
    response = login_session.post("http://f5-ws5.ndmctsgh.edu.tw/eForm/Patient/Result", data=ptsearchdata, verify=False, headers=headers)
    return(response)

def create_pt_list():
    pt_oc=get_pt_oc() ## get patient list page original code as Request object
    pt_list= pt_oc.json() ##retract a list of partient information form patient list original code
    global outdates
    global chartnums 
    global hcasenums 
    global NameGenderAges
    global NrBedNos
    global indatetimes
    global vsnames
    outdates=[]
    chartnums=[]
    #六碼者labdataurl會有Q   ===> ok
    hcasenums=[]

    NameGenderAges = []
    NrBedNos = []
    indatetimes = []
    vsnames = []


    for x in pt_list:
        outdate1 = x['OUTDATETIME']
        print("出院日" + outdate1)
        outdates.append(outdate1)
        
        if outdate1 == "":
            chartnum1 = x['CHARTNO']
            print("病歷號" + chartnum1)
            chartnums.append(chartnum1)
            
            hcasenum1 = x['HCASENO']
            print(hcasenum1)
            hcasenums.append(hcasenum1)
            
            NameGenderAge1 = x['NameGenderAge']
            NrBedNo1 = x['NrBedNo']
            indatetime1 = x['INDATETIME'][:7]
            vsname1 = x['VSDRNAME']
            NameGenderAges.append(NameGenderAge1)
            NrBedNos.append(NrBedNo1)
            indatetimes.append(indatetime1)
            vsnames.append(vsname1)
####
def aligns_last(string, length = 73):
    difference = length - len(string) 
    
    new_string = ""
    space = " "
    # 計算限定長度為20時需要補齊多少個空格
    if difference == 0: # 若差值為0則不需要補
        return string
    
    
    elif difference < 0: 
        print('錯誤：限定的對齊長度小於字元串長度!') 
        return None
    
    
    else: 
        new_string = string[:-31] + space * difference + string[-31:]
        return new_string
    # 若是全形，則不轉換 return new_string + space*(difference) 
    # 返回補齊空格後的字元串 

def aligns_long(string, length = 86):
    difference = length - len(string) 
    
    new_string = ""
    space = " "
    # 計算限定長度為20時需要補齊多少個空格
    if difference == 0: # 若差值為0則不需要補
        return string
    
    
    elif difference < 0: 
        print('錯誤：限定的對齊長 度小於字元串長度!') 
        return None
        
    
    
    else: 
        new_string = string[:-44] + space * difference + string[-44:]
        return new_string

def aligns_drugname(string, length = 20):
    difference = length - len(string) 
    
    new_string = ""
    space = " "
    # 計算限定長度為20時需要補齊多少個空格
    if difference == 0: # 若差值為0則不需要補
        return string
    
    
    elif difference < 0: 
        print('錯誤：限定的對齊長度小於字元串長度!') 
        return string[:20]
    
    
    else: 
        new_string = string + space * difference
        return new_string
    # 若是全形，則不轉換 return new_string + space*(difference)
    # 返回補齊空格後的字元串 
##################


def get_druglist():
    global patients_drugs
    global patients_drugs_forsearch 
    patients_drugs=pd.DataFrame(columns=["商品名", "dose", "freq", "method", "開始時間"])
    patients_drugs_forsearch = []
    for h in range(len(hcasenums)):
        print("drug" + str(h)*10)
        ###################################################################################################
        #drug
        
        drug_url = 'http://mobilereport.ndmctsgh.edu.tw/mr/HISEXNDREPORT.aspx?login_id=' + doc_num + '&special=n&cno=' + chartnums[h]
        headers = {
        #'User-Agent': user_agent.random,
        'Connection': 'keep-alive',
        #'Referer': referlink,
        'Origin': 'http://f5-ws5.ndmctsgh.edu.tw',
        'Referer': 'http://f5-ws5.ndmctsgh.edu.tw/eForm/Account/Login'
        }
        drug_oc = login_session.get(drug_url, verify=False, headers=headers)
        #print(response.content)
        
        
        if drug_oc.status_code == requests.codes.ok:
            print("取得成功")
        else:
            print("取得失敗")
            
        #print(htmlfile4.text)
        
        #import bs4
        drug_soup=bs4.BeautifulSoup(drug_oc.text, 'lxml')
        
        
        #div data-role="content"
        drugdata = drug_soup.find('div', {'data-role': 'content'})
        #print("資料型態", type(table))
        #print("串列長度", len(table))
        #print(drugdata)
        print(drugdata.text)
        print(drugdata.text[87:-1])  ##delete the most upper line
        
        #patients_drugs.append(drugdata.text[86:])
        
        
        
        
        druglist = drugdata.text[87:-1].splitlines() ##put a bunch of string seperated by \n into 

        
        drugs= []
        
        drug_dose = []
        
        drug_freq = []
        
        drug_method = []
        
        drug_start_time = []
        #drugothers = []
        
        #drugs_forsearch = []
        
        
        for drugline in druglist:
            #diff = 84-len(drugline)        
            #if len(drugline) > 75:
            #    drugline = aligns_long(drugline)
            #else:
            #    drugline = aligns_last(drugline)

            
           
            
            

            '''
            #drugname = drugline[:30]
            drugname_forsearch = drugline[:42]
            drugs_forsearch.append(drugname_forsearch)
            '''
            '''
            drugother = drugline[43:62]  #686
            drugothers.append(drugother)
            '''
            drugname = drugline[:42].title()
            drugs.append(drugname)
            
            #686
            drug_dose1 = drugline[42:48]
            drug_dose.append(drug_dose1)
            drug_freq1 = drugline[48:56]
            drug_freq.append(drug_freq1)
            drug_method1 = drugline[56:62]
            drug_method.append(drug_method1)
            
            drug_start_time1 = drugline[62:73]
            drug_start_time.append(drug_start_time1)
            #drugs.append(drugname + drugother)
            
            
        #drugdata_cut = '\n'.join(drugs)
        #patients_drugs.append(drugdata_cut)
        
        druginfo_dataframe = pd.DataFrame(list(zip(drugs, drug_dose, drug_freq, drug_method, drug_start_time)),
                                    columns=["商品名", "dose", "freq", "method", "開始時間"])
        
        
        
        
        
        
        
        
        patients_drugs.append(druginfo_dataframe)
        
        patients_drugs_forsearch.append(drugs)


        druginfo_dataframe

##MAIN CODE
progressnote_oc= get_oc("http://f5-ws5.ndmctsgh.edu.tw/eForm/Account/Login")## get progress note main page original code
progressnote_soup=bs4.BeautifulSoup(progressnote_oc.text, "lxml") ##Soup the progress note page

###above neccessary??

create_pt_list()
get_druglist()
#testing
print(NameGenderAges)
    #print("Prof:"+ a["NameGenderAge"] + "   CHARTNO"+ a["CHARTNO"] + "   VS:" + a["VSDRNAME"] +"   Bednumber:" + a['NrBedno'] + "\n")