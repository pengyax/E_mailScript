# coding=utf-8
from mailbox import Mailbox
import xlwings as xw
import pandas as pd
import win32com.client as win32
import numpy as np
import time

def goal(vendor,reworkgoal=0,cpmgoal=0,rework=0,cpm=0):
    if vendor == 'Both':
        # 使用三元运算符来简化条件判断
        rework_result = "meet" if rework <= reworkgoal else "don't meet"
        cpm_result = "meet" if cpm <= cpmgoal else "don't meet"
        content = f'''
            <span>US rework rate is {rework:.2%}, goal is {reworkgoal:.2%},</span><span id="goal1"> {rework_result}</span><span> the goal.</span><br>
            <span>US CPM is {cpm:.2f}, goal is {cpmgoal:.2f},</span><span id="goal1"> {cpm_result}</span><span> the goal.</span><br>
            <br>'''
    elif vendor == 'CPM':
        cpm_result = "meet" if cpm <= cpmgoal else "don't meet"
        content = f'''
            <span>US CPM is {cpm:.2f}, goal is {cpmgoal:.2f},</span><span id="goal1"> {cpm_result}</span><span> the goal.</span><br>
            <br>'''
    elif vendor == 'Rework':
        rework_result = "meet" if rework <= reworkgoal else "don't meet"
        content = f'''
            <span>US rework rate is {rework:.2%}, goal is {reworkgoal:.2%},</span><span id="goal1"> {rework_result}</span><span> the goal.</span><br>
            <br>'''
    else:
        content = ''
    return content

def send_mail(path,filename,year,month,sendlist):
    df = pd.read_excel(path+filename,sheet_name='Sheet2')
    df_blitz = pd.read_excel(path+filename,sheet_name='Vendor List')
    df = df.merge(df_blitz,how='left',left_on='vendor_name',right_on='Name')
    df['To'] = df['To'].str.replace('\n', '').str.encode("UTF-8").str.decode("UTF-8")
    df['Bcc'] = df['Bcc'].str.replace('\n', '').str.encode("UTF-8").str.decode("gbk")
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

        CY_Total_Items_YTD = row['CY_Total_Items_YTD']
        CY_Inspection_Item_YTD = row['CY_Inspection_Item_YTD']
        CY_Rework_Item_YTD = row['CY_Rework_Item_YTD']
        CY_Complaints_YTD = row['CY_Complaints_YTD']

        CY_LM_inspection_rate = row['CY_LM_inspection_rate']
        CY_CM_inspection_rate = row['CY_CM_inspection_rate']
        LY_inspection_rate_YTD = row['LY_inspection_rate_YTD']
        CY_inspection_rate_YTD = row['CY_inspection_rate_YTD']

        CY_LM_rework_rate = row['CY_LM_rework_rate']
        CY_CM_rework_rate = row['CY_CM_rework_rate']
        LY_rework_rate_YTD = row['LY_rework_rate_YTD']
        CY_rework_rate_YTD = row['CY_rework_rate_YTD']

        CY_LM_Complaints_YTD = row['CY_LM_Complaints_YTD']
        CY_CM_Complaints_YTD = row['CY_CM_Complaints_YTD']
        LY_Complaints_YTD = row['LY_Complaints_YTD']
        CY_Complaints_YTD = row['CY_Complaints_YTD']
        
        rework_rate = row['Current_Rework_Rate_US']
        cpm = row['Current_CPM_US']
        rework_rate_goal = row['2023_Rework_Rate_Goal']
        cpm_goal = row['2023_CPM_Goal']
        vendor_type = row['Type']
        
        
        outlook = win32.Dispatch('Outlook.Application')
        mail = outlook.CreateItem(0) # 0: creat mail
        mail.To = row['To']
        # mail.CC = cc
        mail.BCC = row['Bcc']
        mail.Subject = f'''Monthly Analysis on Top Vendor Report of {short_vendorname} {year} {month_dict[month]}'''
        mail.HTMLBody = f'''
        <!DOCTYPE html>
<html>
<head>
<meta charset="utf-8" />
<title>top_vendor</title> 
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
#goal1 {{font-family:Calibri;font-size:20.0pt;font-weight:bold;color:#00b050}}
#goal2 {{font-family:Calibri;font-size:20.0pt;font-weight:bold;color:#d73333}}
span {{font-family:Calibri;font-size:11.0pt;color:#2a2b2e}}
</style>
</head>
<body>
<div class = "content">
<span>Hi {short_vendorname} Team,</span><br>
<br>
<span>In this year, we will share the Top Vendor's Report to you on each month. The data includes the inspection results total items and complaints information in the previous month. Total data range from 2022 January to the latest month of 2023.</span><br>
<br>
<span>{year} YTD Total number is {CY_Total_Items_YTD:,}, inspection number is {CY_Inspection_Item_YTD:,}, rework number is {CY_Rework_Item_YTD:,}, complaint number is {CY_Complaints_YTD:,}.</span><br>
<br>
<span>Please find the following data change (US + Non-US):</span><br>
</div>
<table class ="data_table">
        <tr class = "header_td"><td></td><td>{year - 1 if month-1 == 0 else year}.{12 if month-1 == 0 else month-1}</td><td>{year}.{month}</td><td>{year-1} YTD</td><td>{year} YTD</td></tr>
        <tr class = "header_td"><td>Inspection Rate</td><td>{CY_LM_inspection_rate:.2%}</td><td>{CY_CM_inspection_rate:.2%}</td><td>{LY_inspection_rate_YTD:.2%}</td><td>{CY_inspection_rate_YTD:.2%}</td></tr>
        <tr class = "header_td"><td>Rework Rate</td><td>{CY_LM_rework_rate:.2%}</td><td>{CY_CM_rework_rate:.2%}{red if CY_CM_rework_rate - CY_LM_rework_rate >0 else (green if CY_CM_rework_rate - CY_LM_rework_rate<0 else '')}</td><td>{LY_rework_rate_YTD:.2%}</td><td>{CY_rework_rate_YTD:.2%}{red if CY_rework_rate_YTD - LY_rework_rate_YTD >0 else (green if CY_rework_rate_YTD - LY_rework_rate_YTD<0 else '')}</td></tr>
        <tr class = "header_td"><td>Complaint Number</td><td>{CY_LM_Complaints_YTD:,}</td><td>{CY_CM_Complaints_YTD:,}{red if CY_CM_Complaints_YTD - CY_LM_Complaints_YTD > 0 else (green if CY_CM_Complaints_YTD - CY_LM_Complaints_YTD<0 else '')}</td><td>{LY_Complaints_YTD:,}</td><td>{CY_Complaints_YTD}{red if CY_Complaints_YTD - LY_Complaints_YTD > 0 else (green if CY_Complaints_YTD - LY_Complaints_YTD <0 else '')}</td>
</table><br>
{goal(vendor=vendor_type,reworkgoal=rework_rate_goal,cpmgoal=cpm_goal,rework=rework_rate,cpm=cpm)}
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
            print(row['To'])
            print(row['Bcc'])
            print("===="*6)
            # mail.Attachments.Add(path + f'\\Monthly Analysis on Top Vendor Report of {short_vendorname} {year} {month_dict[month]}.pdf')
            # mail.Attachments.Add(path + f'\\Monthly Analysis on Top Vendor Report of {short_vendorname} {year} {month_dict[month]}.xlsx')
            # mail.Attachments.Add(path + f'\\Monthly Top Vendor Report Guidance.pdf')
            mail.Save() 
            time.sleep(1)
        else:
            continue
    print("Transmission Completed!")

if __name__=="__main__":
    path = r'C:\Medline\Top Vendor\script'
    filename = '/29 vendors database-test.xlsx'
    year = 2023
    month = 1
    sender_name = 'Zixin Nie'
    sender_cell = '186-7234-6819'
    send_list= ['Amsino',
                'Cobes',
                'Com Bridge',
                'Conod',
                'Danameco',
                'Dieu Thuong',
                'E Test',
                'Hong De',
                'Eco Medi Glove',
                'Gcmedica',
                'Transtek',
                'Dunli',
                'Hartalega',
                'Jianerkang',
                'Jumao',
                'Jie Gao',
                'Kossan',
                'Lotus',
                'Medisafe',
                'Premier Towels',
                'Principle & Will',
                'Raise',
                'Rang Dong',
                'SES',
                'Minhua',
                'Sino',
                'Trolli King',
                'YTY',
                'Assure'
                ]
    send_mail(path,filename,year,month,send_list)

