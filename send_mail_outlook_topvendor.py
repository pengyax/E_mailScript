
# coding=utf-8
from mailbox import Mailbox
import xlwings as xw
import pandas as pd
import win32com.client as win32
import numpy as np
import time

def send_mail(path,filename,year,month,sendlist):
    df = pd.read_excel(path+filename,sheet_name='Sheet2')
    df['To'] = df['To'].str.replace('\n', '').str.encode("UTF-8").str.decode("UTF-8")
    df['Bcc'] = df['Bcc'].str.replace('\n', '').str.encode("UTF-8").str.decode("gbk")
    df['CY_LM_rejection_rate'] = df['CY_LM_rejection_item'] / df['CY_LM_Inspection_Item']
    df['CY_CM_rejection_rate'] = df['CY_CM_rejection_item'] / df['CY_CM_Inspection_Item']
    df['LY_rejection_rate_YTD'] = df['LY_rejection_item_YTD'] / df['LY_Inspection_Item_YTD']
    df['CY_rejection_rate_YTD'] = df['CY_rejection_item_YTD'] / df['CY_Inspection_Item_YTD']
    df['CY_LM_rejection_rate'].replace(np.nan, 0,inplace=True)
    df['CY_CM_rejection_rate'].replace(np.nan, 0,inplace=True)
    df['LY_rejection_rate_YTD'].replace(np.nan, 0,inplace=True)
    df['CY_rejection_rate_YTD'].replace(np.nan, 0,inplace=True)
    red = """<span style="color:red">↑</span>"""
    green = """<span style="color:#00b050">↓</span>"""
    month_dict = {1:'January',
                  2:'February',
                  3:'March',
                  4:'April',
                  5:'May',
                  6:'June',
                  7:'July',
                  8:'August',
                  9:'September',
                  10:'October',
                  11:'November',
                  12:'December'}
    for index,row in df.iterrows():
        short_vendorname = row['short_vendorname']

        CY_Shipping_Item_YTD = row['CY_Shipping_Item_YTD']
        CY_Inspection_Item_YTD = row['CY_Inspection_Item_YTD']
        CY_rejection_item_YTD = row['CY_rejection_item_YTD']
        CY_Rework_Item_YTD = row['CY_Rework_Item_YTD']

        CY_LM_inspection_rate = row['CY_LM_inspection_rate']
        CY_CM_inspection_rate = row['CY_CM_inspection_rate']
        LY_inspection_rate_YTD = row['LY_inspection_rate_YTD']
        CY_inspection_rate_YTD = row['CY_inspection_rate_YTD']

        CY_LM_rejection_rate = row['CY_LM_rejection_rate']
        CY_CM_rejection_rate = row['CY_CM_rejection_rate']
        LY_rejection_rate_YTD = row['LY_rejection_rate_YTD']
        CY_rejection_rate_YTD = row['CY_rejection_rate_YTD']

        CY_LM_rework_rate = row['CY_LM_rework_rate']
        CY_CM_rework_rate = row['CY_CM_rework_rate']
        LY_rework_rate_YTD = row['LY_rework_rate_YTD']
        CY_rework_rate_YTD = row['CY_rework_rate_YTD']
        
        outlook = win32.Dispatch('Outlook.Application')
        mail = outlook.CreateItem(0) # 0: creat mail
        mail.To = row['To']
        # mail.CC = cc
        mail.BCC = row['Bcc']
        mail.Subject = f'''Monthly Inspection Report of {short_vendorname} {year} {month_dict[month]}'''
        mail.HTMLBody = f'''
        <!DOCTYPE html>
<html>
<head>
<meta charset="utf-8" />
<title>signature</title> 
<style type="text/css">
div.content {{font-family:Calibri;font-size:11.0pt;color:#000000}}
div.signature {{font-family:Calibri;font-size:11.0pt}}
table,table tr th, table tr td {{ border:0.5px solid #000000;}}
table {{ width: 200px; min-height: 25px; line-height: 25px; text-align: center; border-collapse: collapse; padding:2px;font-family:Calibri;font-size:11.0pt}}
td {{ text-align: left;white-space: nowrap;text-indent:0.5em}}
a {{text-decoration-line: underline;
    text-decoration-thickness: 1.25px;
    text-decoration-style: initial;
    text-decoration-color: #000000}}
</style>
</head>
<body>
<div class = "content">
<span>Hi {short_vendorname} Team,</span><br>
<br>
<span>Please find the attached report from {month_dict[1]} {year-1} to {month_dict[month]} {year}.</span><br>
<span>{year} YTD shipment number is {CY_Shipping_Item_YTD:,}, inspection number is {CY_Inspection_Item_YTD:,}, reject number is {CY_rejection_item_YTD:,}, rework number is {CY_Rework_Item_YTD:,}.</span><br>
<br>
<span>Please find the following data change:</span><br>
</div>
<table class ="data_table">
        <tr class = "header_td"><td></td><td>{year}.{month-1}</td><td>{year}.{month}</td><td>{year-1} YTD</td><td>{year} YTD</td></tr>
        <tr class = "header_td"><td>Inspection Rate</td><td>{CY_LM_inspection_rate:.2%}</td><td>{CY_CM_inspection_rate:.2%}</td><td>{LY_inspection_rate_YTD:.2%}</td><td>{CY_inspection_rate_YTD:.2%}</td></tr>
        <tr class = "header_td"><td>Rejection Rate</td><td>{CY_LM_rejection_rate:.2%}</td><td>{CY_CM_rejection_rate:.2%}{red if CY_CM_rejection_rate - CY_LM_rejection_rate >0 else (green if CY_CM_rejection_rate - CY_LM_rejection_rate<0 else '')}</td><td>{LY_rejection_rate_YTD:.2%}</td><td>{CY_rejection_rate_YTD:.2%}{red if CY_rejection_rate_YTD - LY_rejection_rate_YTD >0 else (green if CY_rejection_rate_YTD - LY_rejection_rate_YTD<0 else '')}</td></tr>
        <tr class = "header_td"><td>Rework Rate</td><td>{CY_LM_rework_rate:.2%}</td><td>{CY_CM_rework_rate:.2%}{red if CY_CM_rework_rate - CY_LM_rework_rate >0 else (green if CY_CM_rework_rate - CY_LM_rework_rate<0 else '')}</td><td>{LY_rework_rate_YTD:.2%}</td><td>{CY_rework_rate_YTD:.2%}{red if CY_rework_rate_YTD - LY_rework_rate_YTD >0 else (green if CY_rework_rate_YTD - LY_rework_rate_YTD<0 else '')}</td></tr>
</table><br>
<div class = "signature">
<span style="color:#000000">Best regards,</span><br>
<br>
<span style="color:#404040;font-weight: bold">{sender_name}</span><br>
<span style="color:#545454">QA Specialist (Data Analyst)</span><br>
<span style="color:#545454">GSO Wuhan</span><br>
<span style="color:#545454">Medline Industries, LP</span><br>
<a href="http://www.medline.com/?cmpid=eid:signature-link-US-Sales"><span style="color:#58585B">www.medline.com</span></a><br>
<br>
<span style="color:#545454">Cell: {sender_cell}</span><br>
<span style="font-size:10.0pt;color:gray;mso-themecolor:background1;mso-themeshade:
    128">Room1601, 16F, Zall International Center, No. 588<o:p></o:p></span><br>
<span style="font-size:10.0pt;color:gray;mso-themecolor:background1;mso-themeshade:
128">Jianshe Avenue,Jianghan District, Wuhan, China,<o:p></o:p></span>
</div>
</body> 
</html> 
        '''
        if  short_vendorname in sendlist:
            print(short_vendorname)
            print(LY_rejection_rate_YTD)
            print(f'{LY_rejection_rate_YTD:.2%}')
            print(row['To'])
            print(row['Bcc'])
            print("===="*6)
            mail.Attachments.Add(path + f'\\Monthly Inspection Report of {short_vendorname} {year} {month_dict[month]}.pdf')
            mail.Attachments.Add(path + f'\\Monthly Inspection Report of {short_vendorname} {year} {month_dict[month]}.xlsx')
            mail.Send()
            time.sleep(1)
        else:
            continue
    print("Transmission Completed!")

if __name__=="__main__":
    path = r'C:\Medline\medline_project\script\Outlook'
    # path = r'C:\Users\znie\Documents\1.0 DHR\06. Inspection Analysis on Top Vendors\Auto sending Email code & excel'
    filename = '/35 vendors database.xlsx'
    year = 2022
    month = 9
    sender_name = 'Z'
    sender_cell = '186-7234-6819'
    send_list= ['Cobes'
                'Allmed',
                'Amsino',
                'Bliss Health',
                'Cobes Myanmar',
                'Cobes',
                'Com Bridge',
                'Conod',
                'E Test',
                'Hong De',
                'Eco Medi Glove',
                'Emerald',
                'Gcmedica',
                'Dunli',
                'Hengxiang',
                'Hongray',
                'Safeway',
                'Intco',
                'Jianerkang',
                'Jie Gao',
                'Kossan',
                'Master and Frank',
                'Medisafe',
                'Shengyurui',
                'Polaris',
                'Principle & Will',
                'Raise',
                'SES',
                'Minhua',
                'Sino',
                'Space Hero',
                'Jingle',
                'Trolli King',
                'Beauty and Health',
                'Assure'
                ]
    send_mail(path,filename,year,month,send_list)


