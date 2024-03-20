import os
import pandas as pd
import win32com.client as win32
import datetime
import configparser
from subprocess import run, PIPE
import numpy as np
class DCR:
    def __init__(self, dcrNumber, dcrTitle, dcrClass, dcrPSB):
        self.dcrNumber = dcrNumber
        self.dcrTitle = dcrTitle
        self.dcrClass = dcrClass
        self.dcrPSB = dcrPSB

def sendmail(touser, excelpath, psbdictitem):
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = touser
    mail.Subject = 'Attention for DCR Due Date'
    mail.Body = 'Would you please reply this for testing'
    msg = 'For the following DCR,  the Due Date will be reached in few days:<br><font color="red">PSB number:{0}</font><br><font color="red">DCR number:{1}</font><br>Please initiate the necessary activities.'.format(psbdictitem.dcrPSB, psbdictitem.dcrNumber) 
    mail.HTMLBody = msg
    #mail.HTMLBody = '<h2>Draeger DCR Warning</h2>' #this field is optional

    # To attach a file to the email (optional):
    # attachment  = excelpath
    # mail.Attachments.Add(attachment)
    mail.Send()

if __name__ == "__main__":
    # check office network
    # cnt = 1
    # while True:
    #     r = run('ping 10.47.0.18',
    #         stdout=PIPE,
    #         stderr=PIPE,
    #         stdin=PIPE,
    #         shell=True)
    #     if r.returncode:
    #         print('office 网络连接失败，第{}次'.format(cnt))
    #         cnt += 1
    #     else:
    #         break    
    #     if cnt == 3:
    #         quit()
    # get config.ini    
    config = configparser.ConfigParser()
    config.read("config.ini", encoding='utf-8-sig')
    excelpath = config.get('config', 'ExcelPath')
    if excelpath == '':
        excelpath = 'EXPORT.XLSX'
    RegulatoryQualityDay = config.getint('config', 'RegulatoryQualityDay')
    OtherDay = config.getint('config', 'OtherDay')

    # read excel
    data = pd.read_excel(excelpath)
    df = pd.DataFrame(data, columns=['DCR Number','DCR Title','Origination Date', 'DCR Classification','PSB/CFT'])
    psbdict = {}
    index = 0
    print (df.index)
    for i in (df.index):
        raw_date2 = df.loc[i, "Origination Date"]
        dcrclass = df.loc[i, "DCR Classification"]
        df["PSB/CFT"]=df["PSB/CFT"].apply(lambda x: '{0:0>3}'.format(x))
        print (f'The PSB/CFT is {df.loc[i, "PSB/CFT"]}')
        print (f'The line index is {i} and datetime is {raw_date2}' )
        print (f'Origination Date is {pd.isna(df.loc[i, "Origination Date"])}')
        if  pd.isna(df.loc[i, "Origination Date"]) == False:
            formatted_date2 = datetime.datetime.strptime(raw_date2.strftime("%m/%d/%Y"), "%m/%d/%Y")
            curr_date = datetime.datetime.now().strftime("%m/%d/%Y")
            formatted_date1 = datetime.datetime.strptime(curr_date, "%m/%d/%Y")
            total_seconds = (formatted_date2 - formatted_date1).total_seconds()
            diffday = abs(total_seconds / 86400)
            if diffday > RegulatoryQualityDay and (dcrclass=='Regulatory mandate' or dcrclass=='Quality'):
             # add PSB to map
                psbdict[index] = DCR(df.loc[i, "DCR Number"], df.loc[i, "DCR Title"], df.loc[i, "DCR Classification"], df.loc[i, "PSB/CFT"])
                index+=1
            if diffday > OtherDay and (dcrclass!='Regulatory mandate' and dcrclass!='Quality'):
                # add PSB to map
                psbdict[index] = DCR(df.loc[i, "DCR Number"], df.loc[i, "DCR Title"], df.loc[i, "DCR Classification"], df.loc[i, "PSB/CFT"])
                index+=1

    mailsdata = pd.read_excel(r'list.xlsx')
    dfmail = pd.DataFrame(mailsdata, columns=['PSB/CFT', 'Originator'])
    dfmail["PSB/CFT"]=dfmail["PSB/CFT"].apply(lambda x: '{0:0>3}'.format(x))
    # get mail address list
    for i in range(len(psbdict)):
        tf = dfmail["PSB/CFT"] == psbdict[i].dcrPSB
        ttf = dfmail.loc[tf, :]
        mailaddress = ttf.iloc[0, 1]
        sendmail(mailaddress, excelpath, psbdict[i])