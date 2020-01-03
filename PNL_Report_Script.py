import win32com.client
import os
import datetime as datetime
import numpy as np
import pandas as pd
import xlsxwriter
import datetime
import io

def subject_name(file_text):
    now = datetime.datetime.now()
    weekday = now.weekday()
    if weekday == 0:
        today = now - datetime.timedelta(days=3)
        yesterday = today - datetime.timedelta(days=1)
    elif weekday == 1:
        today = now - datetime.timedelta(days=1)
        yesterday = today - datetime.timedelta(days=3)
    else:
        today = now - datetime.timedelta(days=1)
        yesterday = today - datetime.timedelta(days=1)
    current = today.strftime("%Y-%m-%d")
    current = str(file_text)+str(current)
    PNL_Report_Date = today.strftime("%m.%d.%Y")
    PNL_Report_Date = str(PNL_Report_Date)+'PNL Discrepancy'
    now = datetime.datetime.now()
    PNL_Report_Write_to_Date = now.strftime('%m.%d.%Y')
    PNL_Report_Write_to_Date = str(PNL_Report_Write_to_Date)+'PNL Discrepancy'
    today = file_text + str(today)
    yesterday = yesterday.strftime("%Y-%m-%d")
    yesterday_1 = yesterday
    yesterday = file_text + str(yesterday)
    return current,yesterday,PNL_Report_Date,PNL_Report_Write_to_Date

def correct_position_type(Inventory):
    x = Inventory['Position']
    x = pd.to_numeric(x)
    Inventory['Position'] = x
    x = Inventory['P&L']
    x = pd.to_numeric(x)
    Inventory['P&L'] = x
    x = Inventory['MTG Position']
    x = pd.to_numeric(x)
    Inventory['MTG Position'] = x
    return Inventory


file_text = 'Inventory Margin Report for '
x = subject_name(file_text)

"""
Pull and Generate new QTY_DSP_Cleared_Positions file
"""
PNL_Report_Date = x[2]
PNL_Report_File_Most_Recent = 'P:/2. Corps/PNL_Daily_Report/Reports/PNL_Report.xlsx'
Cleared_Yesterday_PNL_Report = pd.read_excel(PNL_Report_File_Most_Recent,sheet_name = 'Quantity Diff')
Cleared_Yesterday_PNL_Report = Cleared_Yesterday_PNL_Report[['Security','Account','Cusip','QTY DSP','Position Notes']]
Cleared_Yesterday_PNL_Report.dropna(inplace = True)
QTY_DSP_Cleared_Positions = 'P:/2. Corps/PNL_Daily_Report/Cleared_Position_File/QTY_DSP_Cleared_Positions.xlsx'
QTY_DSP_Cleared_Positions= pd.read_excel(QTY_DSP_Cleared_Positions,index = False)
QTY_DSP_Cleared_Positions = QTY_DSP_Cleared_Positions.append(Cleared_Yesterday_PNL_Report)

QTY_DSP_Cleared_Positions.drop_duplicates(keep='first',inplace = True)
writer = pd.ExcelWriter('P:/2. Corps/PNL_Daily_Report/Cleared_Position_File/QTY_DSP_Cleared_Positions.xlsx', engine='xlsxwriter')
QTY_DSP_Cleared_Positions.to_excel(writer)
writer.save()


today = x[0]
yesterday = x[1]
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder(6) # "6" refers to the index of a folder - in this case the inbox. You can change that number to reference
messages = inbox.Items
messages.Sort("[ReceivedTime]", True)
message = messages.GetFirst()

i = 0
while i < 20:
    if message.Subject == today:
        try:
            attachments = message.Attachments
            attachment = attachments.Item(1)
            attachment.SaveASFile('P:/2. Corps/PNL_Daily_Report/HT_Files/' + str(today)+'.xlsx') #Saves to the attachment to current folder
            print('HT File Found')
            message = messages.GetNext()

        except:
            message = messages.GetNext()
    i += 1
    message = messages.GetNext()


file_text = 'Report "TW 16 22" is generated.'
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder(6) # "6" refers to the index of a folder - in this case the inbox. You can change that number to reference
messages = inbox.Items
messages.Sort("[ReceivedTime]", True)
message = messages.GetFirst()
i = 0
while i < 10:
   
    if message.Subject == file_text:
        print(message.Subject)
    
        try:
            attachments = message.Attachments
            attachment = attachments.Item(2)
            attachment.SaveASFile('C:\\Users\\ccraig\\Desktop\\PNL Project' + '\\Bloomberg TW'+'.xls') #Saves to the attachment to current folder
            message = messages.GetNext()

        except:
            message = messages.GetNext()
    i += 1
    message = messages.GetNext()
    
Bloomberg_Inventory_16_22 = pd.read_table('C:/Users/ccraig/Desktop/PNL Project/Bloomberg TW.xls')
Bloomberg_Inventory_16_22 = Bloomberg_Inventory_16_22.transpose()
Bloomberg_Inventory_16_22.reset_index(inplace = True)
Bloomberg_Inventory_16_22.drop('index',axis = 1, inplace = True)
Bloomberg_Inventory_16_22.drop(0,axis = 1, inplace = True)
Bloomberg_Inventory_16_22 = Bloomberg_Inventory_16_22.transpose()
Bloomberg_Inventory_16_22 = Bloomberg_Inventory_16_22[:-1]
Bloomberg_Inventory_16_22 = Bloomberg_Inventory_16_22[2:]
Bloomberg_Inventory_16_22 = Bloomberg_Inventory_16_22[0].str.split(',',expand=True)
Bloomberg_Inventory_16_22.rename(columns={0: 'Security'}, inplace=True)
Bloomberg_Inventory_16_22.rename(columns={1: 'P&L'},inplace =True)
Bloomberg_Inventory_16_22.rename(columns={2: 'Cusip'},inplace =True)
Bloomberg_Inventory_16_22.rename(columns={3: 'Position'},inplace =True)
Bloomberg_Inventory_16_22.rename(columns={4: 'Symbol'},inplace =True)
Bloomberg_Inventory_16_22.rename(columns={5: 'Book'},inplace =True)
Bloomberg_Inventory_16_22.rename(columns={6: 'MTG Position'},inplace =True)
Bloomberg_Inventory_16_22 = correct_position_type(Bloomberg_Inventory_16_22)

file_text = 'Report "TW 1 5" is generated.'
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder(6) # "6" refers to the index of a folder - in this case the inbox. You can change that number to reference
messages = inbox.Items
messages.Sort("[ReceivedTime]", True)
message = messages.GetFirst()
i = 0
while i < 10:

    if message.Subject == file_text:
        print(message.Subject)
    
        try:
            attachments = message.Attachments
            attachment = attachments.Item(2)
            attachment.SaveASFile('C:\\Users\\ccraig\\Desktop\\PNL Project' + '\\Bloomberg TW'+'.xls') #Saves to the attachment to current folder
            message = messages.GetNext()

        except:
            message = messages.GetNext()
    i += 1
    message = messages.GetNext()
    
Bloomberg_Inventory_1_5 = pd.read_table('C:/Users/ccraig/Desktop/PNL Project/Bloomberg TW.xls')
Bloomberg_Inventory_1_5 = Bloomberg_Inventory_1_5.transpose()
Bloomberg_Inventory_1_5.reset_index(inplace = True)
Bloomberg_Inventory_1_5.drop('index',axis = 1, inplace = True)
Bloomberg_Inventory_1_5.drop(0,axis = 1, inplace = True)
Bloomberg_Inventory_1_5 = Bloomberg_Inventory_1_5.transpose()
Bloomberg_Inventory_1_5 = Bloomberg_Inventory_1_5[:-1]
Bloomberg_Inventory_1_5 = Bloomberg_Inventory_1_5[2:]
Bloomberg_Inventory_1_5 = Bloomberg_Inventory_1_5[0].str.split(',',expand=True)
Bloomberg_Inventory_1_5.rename(columns={0: 'Security'}, inplace=True)
Bloomberg_Inventory_1_5.rename(columns={1: 'P&L'},inplace =True)
Bloomberg_Inventory_1_5.rename(columns={2: 'Cusip'},inplace =True)
Bloomberg_Inventory_1_5.rename(columns={3: 'Position'},inplace =True)
Bloomberg_Inventory_1_5.rename(columns={4: 'Symbol'},inplace =True)
Bloomberg_Inventory_1_5.rename(columns={5: 'Book'},inplace =True)
Bloomberg_Inventory_1_5.rename(columns={6: 'MTG Position'},inplace =True)
Bloomberg_Inventory_1_5 = correct_position_type(Bloomberg_Inventory_1_5)

file_text = 'Report "TW 6 10" is generated.'
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder(6) # "6" refers to the index of a folder - in this case the inbox. You can change that number to reference
messages = inbox.Items
messages.Sort("[ReceivedTime]", True)
message = messages.GetFirst()
i = 0
while i < 10:
   
    if message.Subject == file_text:
        print(message.Subject)
    
        try:
            attachments = message.Attachments
            attachment = attachments.Item(2)
            attachment.SaveASFile('C:\\Users\\ccraig\\Desktop\\PNL Project' + '\\Bloomberg TW'+'.xls') #Saves to the attachment to current folder
            message = messages.GetNext()

        except:
            message = messages.GetNext()
    i += 1
    message = messages.GetNext()
    
Bloomberg_Inventory_6_10 = pd.read_table('C:/Users/ccraig/Desktop/PNL Project/Bloomberg TW.xls')
Bloomberg_Inventory_6_10 = Bloomberg_Inventory_6_10.transpose()
Bloomberg_Inventory_6_10.reset_index(inplace = True)
Bloomberg_Inventory_6_10.drop('index',axis = 1, inplace = True)
Bloomberg_Inventory_6_10.drop(0,axis = 1, inplace = True)
Bloomberg_Inventory_6_10 = Bloomberg_Inventory_6_10.transpose()
Bloomberg_Inventory_6_10 = Bloomberg_Inventory_6_10[:-1]
Bloomberg_Inventory_6_10 = Bloomberg_Inventory_6_10[2:]
Bloomberg_Inventory_6_10 = Bloomberg_Inventory_6_10[0].str.split(',',expand=True)
Bloomberg_Inventory_6_10.rename(columns={0: 'Security'}, inplace=True)
Bloomberg_Inventory_6_10.rename(columns={1: 'P&L'},inplace =True)
Bloomberg_Inventory_6_10.rename(columns={2: 'Cusip'},inplace =True)
Bloomberg_Inventory_6_10.rename(columns={3: 'Position'},inplace =True)
Bloomberg_Inventory_6_10.rename(columns={4: 'Symbol'},inplace =True)
Bloomberg_Inventory_6_10.rename(columns={5: 'Book'},inplace =True)
Bloomberg_Inventory_6_10.rename(columns={6: 'MTG Position'},inplace =True)
Bloomberg_Inventory_6_10 = correct_position_type(Bloomberg_Inventory_6_10)


file_text = 'Report "TW 11 15" is generated.'
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder(6) # "6" refers to the index of a folder - in this case the inbox. You can change that number to reference
messages = inbox.Items
messages.Sort("[ReceivedTime]", True)
message = messages.GetFirst()
i = 0
while i < 10:
   
    if message.Subject == file_text:
        print(message.Subject)
        try:
            attachments = message.Attachments
            attachment = attachments.Item(2)
            attachment.SaveASFile('C:\\Users\\ccraig\\Desktop\\PNL Project' + '\\Bloomberg TW'+'.xls') #Saves to the attachment to current folder
            message = messages.GetNext()

        except:
            message = messages.GetNext()
    i += 1
    message = messages.GetNext()
    
Bloomberg_Inventory_11_15 = pd.read_table('C:/Users/ccraig/Desktop/PNL Project/Bloomberg TW.xls')
Bloomberg_Inventory_11_15 = Bloomberg_Inventory_11_15.transpose()
Bloomberg_Inventory_11_15.reset_index(inplace = True)
Bloomberg_Inventory_11_15.drop('index',axis = 1, inplace = True)
Bloomberg_Inventory_11_15.drop(0,axis = 1, inplace = True)
Bloomberg_Inventory_11_15 = Bloomberg_Inventory_11_15.transpose()
Bloomberg_Inventory_11_15 = Bloomberg_Inventory_11_15[:-1]
Bloomberg_Inventory_11_15 = Bloomberg_Inventory_11_15[2:]
Bloomberg_Inventory_11_15 = Bloomberg_Inventory_11_15[0].str.split(',',expand=True)
Bloomberg_Inventory_11_15.rename(columns={0: 'Security'}, inplace=True)
Bloomberg_Inventory_11_15.rename(columns={1: 'P&L'},inplace =True)
Bloomberg_Inventory_11_15.rename(columns={2: 'Cusip'},inplace =True)
Bloomberg_Inventory_11_15.rename(columns={3: 'Position'},inplace =True)
Bloomberg_Inventory_11_15.rename(columns={4: 'Symbol'},inplace =True)
Bloomberg_Inventory_11_15.rename(columns={5: 'Book'},inplace =True)
Bloomberg_Inventory_11_15.rename(columns={6: 'MTG Position'},inplace =True)
Bloomberg_Inventory_11_15 = correct_position_type(Bloomberg_Inventory_11_15)

Bloomberg_Inventory = pd.concat([Bloomberg_Inventory_16_22,
                                 Bloomberg_Inventory_1_5,
                                 Bloomberg_Inventory_6_10,
                                 Bloomberg_Inventory_11_15],ignore_index = True)

"""
Read in Excel files from Hilltop and Bloomberg

"""

# 'P:/2. Corps/PNL_Daily_Report/HT_Files/' + str(today)+'.xlsx'
TW_Inventory_Date_Text = today[28:]
Bloomberg_Inventory['Cusip'] = Bloomberg_Inventory['Symbol'].str[1:10]
writer = pd.ExcelWriter('P:/2. Corps/Daily_TW_Files/' + TW_Inventory_Date_Text + ' TW Inventory.xlsx', engine='xlsxwriter')
Bloomberg_Inventory.to_excel(writer,index = False)
writer.save()


Bloomberg_Inventory = Bloomberg_Inventory[['Cusip', 'P&L', 'Security', 'Position','Symbol','Book','MTG Position']]
Bloomberg_Inventory['Position'] = Bloomberg_Inventory['Position']*1000
Bloomberg_Inventory['MTG Position'] = Bloomberg_Inventory['MTG Position']*1000
# Bloomberg_Inventory['Cusip'] = Bloomberg_Inventory['Cusip'].astype(str)
print(today)
print(yesterday)
Recent = 'P:/2. Corps/PNL_Daily_Report/HT_Files/' + str(today)+'.xlsx'#Recent = 'C:/Users/ccraig/Desktop/PNL Project/'+str(today)+'.xlsx'
Old = 'P:/2. Corps/PNL_Daily_Report/HT_Files/' + str(yesterday)+'.xlsx'
Hilltop_Recent_x = pd.read_excel(io=Recent, sheet_name='Detail')
HT_Detail = Hilltop_Recent_x
# Hilltop_Recent_x['Cusip'] = Hilltop_Recent['Cusip'].astype(str)
Hilltop_Old_y = pd.read_excel(io=Old, sheet_name='Detail')
# Hilltop_Old_y['Cusip'] = Hilltop_Old_y['Cusip'].astype(str)
Hilltop_Recent_s = pd.read_excel(io=Recent, sheet_name='Summary')
Hilltop_Old_s = pd.read_excel(io=Old, sheet_name='Summary')
Hilltop_Recent_s = Hilltop_Recent_s.head(10)
Hilltop_Old_s = Hilltop_Old_s.head(10)
Hilltop_Recent_x['Cusip_group_by'] = Hilltop_Recent_x['Cusip']
Hilltop_Recent_x['Cusip_group_by'] = 'C'+ Hilltop_Recent_x['Cusip_group_by']
Bloomberg_Inventory = Bloomberg_Inventory.groupby(['Cusip']).agg({'P&L':'sum',
                                                                   'Security':'first',
                                                                   'Position':'sum',
                                                                   'Symbol':'first',
                                                                   'Book':'first',
                                                                   'MTG Position':'sum'})
TW_Detail = Bloomberg_Inventory

"""
Fix MTG Position
"""
# Bloomberg_Inventory.loc[(Bloomberg_Inventory['Book'] == '8763') | (Bloomberg_Inventory['Book']=='IO'), 'Position'] = 'MTG Position'


Hilltop_Recent = Hilltop_Recent_x.groupby(['Cusip_group_by']).agg({'Quantity':'sum',
                                                                   'Unreal PNL':'sum',
                                                                   'Real PNL':'sum',
                                                                   'Requirement':'sum',
                                                                   'Cusip':'first',
                                                                   'Description':'first',
                                                                   'Price':'mean'})

Hilltop_Old_y['Cusip_group_by'] = Hilltop_Old_y['Cusip']
Hilltop_Old = Hilltop_Old_y.groupby(['Cusip_group_by']).agg({'Quantity':'sum',
                                                             'Unreal PNL':'sum',
                                                             'Real PNL':'sum',
                                                             'Requirement':'sum',
                                                             'Cusip':'first',
                                                             'Description':'first',
                                                             'Price':'mean'})

"""
Merge Hilltop Recent and Old together

"""

Hilltop_Recent = pd.merge(Hilltop_Recent, Hilltop_Old, on='Cusip', how='left')

Hilltop_Recent = pd.merge(Hilltop_Recent, Bloomberg_Inventory, on='Cusip', how='outer')

Hilltop_Recent_x = Hilltop_Recent_x[['Cusip','Account Name']]

Hilltop_Recent = pd.merge(Hilltop_Recent, Hilltop_Recent_x,on='Cusip',how='outer')

Hilltop_Recent = Hilltop_Recent.fillna(0)
Hilltop_Recent.loc[Hilltop_Recent['Security']==0,'Security'] = Hilltop_Recent['Description_x']

"""
Fix Account Names returning 0
"""
Hilltop_Recent.loc[(Hilltop_Recent['Account Name'] == 0) & (Hilltop_Recent['Book']=='8701'), 'Account Name'] = 'K74 Corporates'
Hilltop_Recent.loc[(Hilltop_Recent['Account Name'] == 0) & (Hilltop_Recent['Book']=='CORPSP'), 'Account Name'] = 'P02 Corp SP'
Hilltop_Recent.loc[(Hilltop_Recent['Account Name'] == 0) & (Hilltop_Recent['Book']=='CORPFRN'), 'Account Name'] = 'P01 Corp Floate'
Hilltop_Recent.loc[(Hilltop_Recent['Account Name'] == 0) & (Hilltop_Recent['Book']=='8503'), 'Account Name'] = 'K72 Muni Inv FI'
Hilltop_Recent.loc[(Hilltop_Recent['Account Name'] == 0) & (Hilltop_Recent['Book']=='8763'), 'Account Name'] = 'K76 S P Inv'
Hilltop_Recent.loc[(Hilltop_Recent['Account Name'] == 0) & (Hilltop_Recent['Book']=='8782'), 'Account Name'] = 'K77 CD Inv'
Hilltop_Recent.loc[(Hilltop_Recent['Account Name'] == 0) & (Hilltop_Recent['Book']=='8937'), 'Account Name'] = 'K78 Taxable Mun'
Hilltop_Recent.loc[(Hilltop_Recent['Account Name'] == 0) & (Hilltop_Recent['Book']=='8938'), 'Account Name'] = 'K79 Cali Tax Ex'
Hilltop_Recent.loc[(Hilltop_Recent['Account Name'] == 0) & (Hilltop_Recent['Book']=='8939'), 'Account Name'] = 'K80 Muni Tax Ex'
Hilltop_Recent.loc[(Hilltop_Recent['Account Name'] == 0) & (Hilltop_Recent['Book']=='8940'), 'Account Name'] = 'K81 Muni Tax'
Hilltop_Recent.loc[(Hilltop_Recent['Account Name'] == 0) & (Hilltop_Recent['Book']=='8941'), 'Account Name'] = 'K82 Tax 0 Muni'
Hilltop_Recent.loc[(Hilltop_Recent['Account Name'] == 0) & (Hilltop_Recent['Book']=='CORPIG'), 'Account Name'] = 'N88 Corp Notes'
Hilltop_Recent.loc[(Hilltop_Recent['Account Name'] == 0) & (Hilltop_Recent['Book']=='CORPNOTE'), 'Account Name'] = 'N90 CD'
Hilltop_Recent.loc[(Hilltop_Recent['Account Name'] == 0) & (Hilltop_Recent['Book']=='IO'), 'Account Name'] = 'M64 Sierra MBS'
Hilltop_Recent = Hilltop_Recent[Hilltop_Recent['Book'] != '8659']
Hilltop_Recent = Hilltop_Recent[Hilltop_Recent['Book'] != '8720']
Hilltop_Recent['Position'] = np.where((Hilltop_Recent['Account Name'] == 'K76 S P Inv'),Hilltop_Recent['MTG Position'],Hilltop_Recent['Position'])

"""
Calculate nessicary values

"""

Hilltop_Recent['Real_PNL_Change'] = Hilltop_Recent['Real PNL_x'] - Hilltop_Recent['Real PNL_y']
Hilltop_Recent['Real Discrepancy'] = Hilltop_Recent['Real_PNL_Change'] - Hilltop_Recent['P&L']
Hilltop_Recent['Quantity Change'] = Hilltop_Recent['Quantity_x'] - Hilltop_Recent['Quantity_y']

Hilltop_Individual_Book_Summary = Hilltop_Recent

Hilltop_Recent.rename(columns={'Quantity Change': 'HT Quantity Change',
                               'Quantity_x': 'HT New Quantity',
                               'Quantity_y': 'HT Old Quantity',
                               'Quantity_x_y': '-  HT Quantity  =',
                               'Real Discrepancy_y': 'TW vs. HT Real Discrepancy',
                               'Position': 'TW Quantity',
                               'Real PNL_x': 'HT New Real PNL',
                               'Account Name': 'Account',
                               'Quantity_y':'HT Old Quantity',
                               'Unreal PNL_x':'HT New Unreal PNL',
                               'Unreal PNL_y':'HT Old Unreal PNL',
                               'Real PNL_y':'HT Old Real PNL',
                               'P&L':'TW PNL',
                               'Quantity_y':'HT Old Quantity',
                               'Price_x':'Price'}, inplace=True)

# # Hilltop_Recent['TW Quantity'] = Hilltop_Recent['TW Quantity'] * 1000

Hilltop_Recent['HT Change in Quantity'] = Hilltop_Recent['HT New Quantity']-Hilltop_Recent['HT Old Quantity']
Hilltop_Recent['TW - HT Quantity Discrepancy'] = Hilltop_Recent['TW Quantity']-Hilltop_Recent['HT New Quantity']
Hilltop_Recent['Real PNL Change'] = Hilltop_Recent['HT New Real PNL']-Hilltop_Recent['HT Old Real PNL']
Hilltop_Recent['HT Real PNL Change'] = Hilltop_Recent['HT New Real PNL']-Hilltop_Recent['HT Old Real PNL']
Hilltop_Recent['Adj Unreal PNL Change'] = Hilltop_Recent['HT New Unreal PNL']-Hilltop_Recent['HT Old Unreal PNL']+Hilltop_Recent['HT Real PNL Change']
Hilltop_Recent['HT-TW PNL Discrepancy'] = Hilltop_Recent['HT Real PNL Change']-Hilltop_Recent['TW PNL']
Hilltop_Recent['Requirement Change'] = Hilltop_Recent['Requirement_x']-Hilltop_Recent['Requirement_y']
Hilltop_Recent['Filter Column'] = Hilltop_Recent['TW - HT Quantity Discrepancy'] + Hilltop_Recent['HT Change in Quantity'] + Hilltop_Recent['Adj Unreal PNL Change'] + Hilltop_Recent['HT-TW PNL Discrepancy'] + Hilltop_Recent['Requirement Change']
# Hilltop_Recent = Hilltop_Recent[(Hilltop_Recent['Filter Column'] != 0)]
Hilltop_Recent = pd.merge(HT_Detail,Hilltop_Recent, on='Cusip', how='left')

Hilltop_Recent = Hilltop_Recent[[
                                 'HT Quantity Change',                         
                                 'Security',                                       #A
                                 'Cusip',                                          #B
                                 'Account',                                        #C
                                 'Price_x',                                        #D
                                 'TW Quantity',                                    #E
                                 'HT New Quantity',                                #F
                                 'TW - HT Quantity Discrepancy',                   #G
                                 'HT Change in Quantity',                          #H
                                 'HT New Unreal PNL',                              #I
                                 'HT Old Unreal PNL',                              #J
                                 'Real PNL Change',                                #K
                                 'Adj Unreal PNL Change',                          #L
                                 'TW PNL',                                         #M
                                 'HT-TW PNL Discrepancy',                          #N
                                 'Requirement Change',
                                 'Requirement_x'
]]
Hilltop_Recent.dropna(thresh = 5,inplace = True) 

"""
# Drop Duplicated Values for Cusip and HT Quantity Change

"""
Hilltop_Recent = Hilltop_Recent.drop_duplicates(['Cusip', 'HT Quantity Change'])


"""
# Set up excel file naming path w/ today's date

"""
time = datetime.datetime.today()
current = time.strftime("%m.%d.%Y")
"""
# Write file to excel

"""
Hilltop_Recent.drop(
    [
        'HT Quantity Change'
    ],
    axis=1, inplace=True)


Hilltop_Individual_Summary = Hilltop_Recent
Hilltop_Recent.sort_values('TW PNL', axis=0, ascending=False, inplace=True)

filepath = 'P:/2. Corps/PNL_Daily_Report/Reports/PNL_Report.xlsx'

writer = pd.ExcelWriter(filepath, engine='xlsxwriter')
"""
            Summary Code

"""
Daily_Change_x = pd.read_excel(io=Recent, sheet_name='Summary')
Daily_Change_y = pd.read_excel(io=Old,sheet_name='Summary')
Daily_Change_x = Daily_Change_x[12:]
Daily_Change_y = Daily_Change_y[12:]
Daily_Change_x.rename(columns={'Account Number': 'Inventory Sub Totals',
                                'Total Available Funds':'Account Name',
                                'Unnamed: 2':'Position Type',
                                'Unnamed: 3':'Cost',
                                'Unnamed: 4':'Market Value',
                                'Unnamed: 5':'Unreal PNL',
                                'Unnamed: 6':'Requirement',
                                'Unnamed: 7':'Real PNL'}, inplace=True)
Daily_Change_x.reset_index(inplace = True)
Daily_Change_y.rename(columns={'Account Number': 'Inventory Sub Totals',
                                'Total Available Funds':'Account Name',
                                'Unnamed: 2':'Position Type',
                                'Unnamed: 3':'Cost',
                                'Unnamed: 4':'Market Value',
                                'Unnamed: 5':'Unreal PNL',
                                'Unnamed: 6':'Requirement',
                                'Unnamed: 7':'Real PNL'}, inplace=True)

Daily_Change_y.reset_index(inplace = True)
Daily_Change_x=pd.merge(Daily_Change_x, Daily_Change_y, on='index', how='left')
Daily_Change_x['Cost'] = Daily_Change_x['Cost_x']-Daily_Change_x['Cost_y']
Daily_Change_x['Market Value']=Daily_Change_x['Market Value_x']-Daily_Change_x['Market Value_y']
Daily_Change_x['Requirement']=Daily_Change_x['Requirement_x']-Daily_Change_x['Requirement_y']
Daily_Change_x['Unreal PNL']=Daily_Change_x['Unreal PNL_x']-Daily_Change_x['Unreal PNL_y']
Daily_Change_x['Real PNL']=Daily_Change_x['Real PNL_x']-Daily_Change_x['Real PNL_y']
Daily_Change_x = Daily_Change_x[['index','Account Name_x','Position Type_x','Cost','Market Value','Requirement','Unreal PNL','Real PNL']]
Daily_Change = Daily_Change_x.groupby('Account Name_x')['Cost','Market Value','Requirement', 'Unreal PNL','Real PNL'].sum()
Daily_Change['Account Name_x']=['K72 Muni Inv','K74 Corporates ','K76 S P Inv',
                                'K77 CD Inv','K78 Taxable Mun',
                                'K79 Cali Tax Ex','K80 Muni Tax Ex',
                                'K81 Muni Tax','K82 Tax 0 Muni','L81 Sierra Comp',
                                'M64 Sierra MBS','N90 CD','P01 Corp Floate','P02 Corp Sp',
                                'N88 Corp Notes','N87 Corp HY','P03 New Corp 01']
Daily_Change = Daily_Change_x.append(Daily_Change, ignore_index = True,sort = False)
Daily_Change['Account Name_x'] = Daily_Change['Account Name_x'].replace({'K72 Muni Inv':'K72 Muni Inv Fl'})
Daily_Change.sort_values(['Account Name_x','Position Type_x'],inplace = True)
Daily_Change.rename(columns={'Account Name_x':'Account Name','Position Type_x':'Position Type'},inplace = True)
Daily_Change.drop('index',axis = 1, inplace = True)

"""
Create and Format Summary sheet

# """
"""
            Summary Code

"""
Summary_Recent = pd.read_excel(io=Recent, sheet_name='Summary')
Summary_Old = pd.read_excel(io=Old,sheet_name='Summary')
Summary_Recent = Summary_Recent[12:]
Summary_Old = Summary_Old[12:]
Summary_Recent.rename(columns={'Account Number': 'Inventory Sub Totals',
                                'Total Available Funds':'Account Name',
                                'Unnamed: 2':'Position Type',
                                'Unnamed: 3':'Cost',
                                'Unnamed: 4':'Market Value',
                                'Unnamed: 5':'Unreal PNL',
                                'Unnamed: 6':'Requirement',
                                'Unnamed: 7':'Real PNL'}, inplace=True)
Summary_Recent.reset_index(inplace = True)
Summary_Old.rename(columns={'Account Number': 'Inventory Sub Totals',
                                'Total Available Funds':'Account Name',
                                'Unnamed: 2':'Position Type',
                                'Unnamed: 3':'Cost',
                                'Unnamed: 4':'Market Value',
                                'Unnamed: 5':'Unreal PNL',
                                'Unnamed: 6':'Requirement',
                                'Unnamed: 7':'Real PNL'}, inplace=True)
Summary_Old.reset_index(inplace = True)

Daily_Change=pd.merge(Summary_Recent, Summary_Old, on='index', how='left')

Daily_Change['Cost'] = Daily_Change['Cost_x']-Daily_Change['Cost_y']
Daily_Change['Market Value']=Daily_Change['Market Value_x']-Daily_Change['Market Value_y']
Daily_Change['Requirement']=Daily_Change['Requirement_x']-Daily_Change['Requirement_y']
Daily_Change['Unreal PNL']=Daily_Change['Unreal PNL_x']-Daily_Change['Unreal PNL_y']
Daily_Change['Real PNL']=Daily_Change['Real PNL_x']-Daily_Change['Real PNL_y']
Daily_Change = Daily_Change[['Account Name_x','Position Type_x','Cost','Market Value','Requirement','Unreal PNL','Real PNL']]
Daily_Change.rename(columns={'Account Name_x': 'Account Name',
                                'Position Type_x':'Position Type'
                            }, inplace=True)
Summary_Recent = Summary_Recent[['Account Name','Position Type','Cost','Market Value','Requirement','Unreal PNL','Real PNL']]
Muni_Summary_Recent = Summary_Recent.reindex([0,1,8,9,10,11,12,13,14,15,16,17]) 
Muni_Daily_Change =  Daily_Change.reindex([0,1,8,9,10,11,12,13,14,15,16,17])
Muni_Summary_Recent_Grouped = Muni_Summary_Recent.groupby(['Account Name']).agg({'Position Type':'first',
                                                                                 'Cost':'sum',
                                                                     'Market Value':'sum',
                                                                    'Requirement':'sum',
                                                                    'Unreal PNL':'sum',
                                                                    'Real PNL':'sum'})
Muni_Summary_Recent.reset_index(inplace = True)
Muni_Summary_Recent_Grouped.reset_index(inplace = True)
Muni_Summary_Recent_Short = Muni_Summary_Recent[Muni_Summary_Recent['Position Type'] == 'Short']

Muni_Summary_Recent_Short['Short Total'] = abs(Muni_Summary_Recent_Short['Cost'] + Muni_Summary_Recent_Short['Market Value'] + Muni_Summary_Recent_Short['Requirement'] + Muni_Summary_Recent_Short['Unreal PNL'] + 
                                             Muni_Summary_Recent_Short['Real PNL'])
Muni_Summary_Recent_Short = Muni_Summary_Recent_Short[Muni_Summary_Recent_Short['Position Type'] == 'Short']
Muni_Summary_Recent_Short = Muni_Summary_Recent_Short[Muni_Summary_Recent_Short['Short Total'] > 0]
Muni_Summary_Recent_Short = Muni_Summary_Recent_Short['Account Name'].tolist()


Muni_Daily_Change_Grouped = Muni_Daily_Change.groupby(['Account Name']).agg({'Position Type':'first',
                                                                             'Cost':'sum',
                                                                     'Market Value':'sum',
                                                                    'Requirement':'sum',
                                                                    'Unreal PNL':'sum',
                                                                    'Real PNL':'sum'})
Muni_Daily_Change.reset_index(inplace = True)
Muni_Daily_Change_Grouped.reset_index(inplace = True)
Muni_Daily_Change_Short = Muni_Daily_Change[Muni_Daily_Change['Position Type'] == 'Short']

Muni_Daily_Change_Short['Short Total'] = abs(Muni_Daily_Change_Short['Cost'] + Muni_Daily_Change_Short['Market Value'] + Muni_Daily_Change_Short['Requirement'] + Muni_Daily_Change_Short['Unreal PNL'] + 
                                             Muni_Daily_Change_Short['Real PNL'])
Muni_Daily_Change_Short = Muni_Daily_Change_Short[Muni_Daily_Change_Short['Position Type'] == 'Short']
Muni_Daily_Change_Short = Muni_Daily_Change_Short[Muni_Daily_Change_Short['Short Total'] > 0]
Muni_Daily_Change_Short = Muni_Daily_Change_Short['Account Name'].tolist()

N87_Summary_Recent  = Summary_Recent.reindex([22,23])
N87_Daily_Change  = Daily_Change.reindex([22,23])

N88_Summary_Recent  = Summary_Recent.reindex([24,25])
N88_Daily_Change  = Daily_Change.reindex([24,25])

N90_Summary_Recent  = Summary_Recent.reindex([26,27])
N90_Daily_Change  = Daily_Change.reindex([26,27])

P01_Summary_Recent  = Summary_Recent.reindex([28,29])
P01_Daily_Change  = Daily_Change.reindex([28,29])

K74_Summary_Recent  = Summary_Recent.reindex([2,3])
K74_Daily_Change  = Daily_Change.reindex([2,3])

L81_Summary_Recent  = Summary_Recent.reindex([18,19])
L81_Daily_Change  = Daily_Change.reindex([18,19])

P02_Summary_Recent  = Summary_Recent.reindex([30,31])
P02_Daily_Change  = Daily_Change.reindex([30,31])

P03_Summary_Recent  = Summary_Recent.reindex([32,33])
P03_Daily_Change  = Daily_Change.reindex([32,33])

CD_Summary_Recent  = Summary_Recent.reindex([6])
CD_Daily_Change  = Daily_Change.reindex([6])

CMO_Summary_Recent  = Summary_Recent.reindex([4,20])
CMO_Daily_Change  = Daily_Change.reindex([4,20])





"""
Create and Format Summary sheet

# """
Hilltop_Summary_new = pd.read_excel(io=Recent, sheet_name='Detail')
Hilltop_Summary_new = Hilltop_Summary_new.groupby(['Account Name']).agg({
                                                          'Unreal PNL':'sum',
                                                          'Real PNL':'sum',
                                                          'Requirement':'sum',
                                                          'Cost':'sum',
                                                          'Market Value':'sum'})
Hilltop_Summary_new = Hilltop_Summary_new[['Cost','Market Value','Requirement','Unreal PNL','Real PNL']]

Hilltop_Summary_new.loc['Column_Total'] = Hilltop_Summary_new.sum(numeric_only=True, axis=0)



# Daily_Change.to_excel(writer,sheet_name ='Summary',index=False,startrow=20)

Muni_Summary_Recent_Grouped.to_excel(writer,sheet_name = 'Summary',index = False,startrow =11 )
Muni_Daily_Change_Grouped.to_excel(writer,sheet_name = 'Summary',index = False,startrow =11, startcol = 8 )
N88_Summary_Recent.to_excel(writer,sheet_name = 'Summary',index = False,startrow =19 )
N88_Daily_Change.to_excel(writer,sheet_name = 'Summary',index = False,startrow =19, startcol = 8)
N90_Summary_Recent.to_excel(writer,sheet_name = 'Summary',index = False,startrow =23 )
N90_Daily_Change.to_excel(writer,sheet_name = 'Summary',index = False,startrow =23, startcol = 8 )
P01_Summary_Recent.to_excel(writer,sheet_name = 'Summary',index = False,startrow =27 )
P01_Daily_Change.to_excel(writer,sheet_name = 'Summary',index = False,startrow =27, startcol = 8 )
P02_Summary_Recent.to_excel(writer,sheet_name = 'Summary',index = False,startrow =31 )
P02_Daily_Change.to_excel(writer,sheet_name = 'Summary',index = False,startrow =31, startcol = 8 )
K74_Summary_Recent.to_excel(writer,sheet_name = 'Summary',index = False,startrow =35 )
K74_Daily_Change.to_excel(writer,sheet_name = 'Summary',index = False,startrow =35, startcol = 8 )
L81_Summary_Recent.to_excel(writer,sheet_name = 'Summary',index = False,startrow =39 )
L81_Daily_Change.to_excel(writer,sheet_name = 'Summary',index = False,startrow =39, startcol = 8 )
P03_Summary_Recent.to_excel(writer,sheet_name = 'Summary',index = False,startrow =43 )
P03_Daily_Change.to_excel(writer,sheet_name = 'Summary',index = False,startrow =43, startcol = 8 )
N87_Summary_Recent.to_excel(writer,sheet_name = 'Summary',index = False,startrow =47 )
N87_Daily_Change.to_excel(writer,sheet_name = 'Summary',index = False,startrow =47, startcol = 8 )

CD_Summary_Recent.to_excel(writer,sheet_name = 'Summary',index = False,startrow = 52)
CD_Daily_Change.to_excel(writer,sheet_name = 'Summary',index = False,startrow = 52, startcol = 8)
CMO_Summary_Recent.to_excel(writer,sheet_name = 'Summary',index = False,startrow = 54)
CMO_Daily_Change.to_excel(writer,sheet_name = 'Summary',index = False,startrow = 54, startcol = 8)

"""
Calculate Column Totals
"""
#       MUNI
Muni_Cost_Summary_Recent = Muni_Summary_Recent['Cost'].sum()
Muni_Market_Value_Summary_Recent = Muni_Summary_Recent['Market Value'].sum()
Muni_Requirement_Summary_Recent = Muni_Summary_Recent['Requirement'].sum()
Muni_Unreal_PNL_Summary_Recent = Muni_Summary_Recent['Unreal PNL'].sum()
Muni_Real_PNL_Summary_Recent = Muni_Summary_Recent['Real PNL'].sum()

Muni_Cost_Daily_Change = Muni_Daily_Change['Cost'].sum()
Muni_Market_Value_Daily_Change = Muni_Daily_Change['Market Value'].sum()
Muni_Requirement_Daily_Change = Muni_Daily_Change['Requirement'].sum()
Muni_Unreal_PNL_Daily_Change = Muni_Daily_Change['Unreal PNL'].sum()
Muni_Real_PNL_Daily_Change = Muni_Daily_Change['Real PNL'].sum()

#      Corp
Corp_N88_Cost_Summary_Recent = N88_Summary_Recent['Cost'].sum()
Corp_N88_Market_Value_Summary_Recent = N88_Summary_Recent['Market Value'].sum()
Corp_N88_Requirement_Summary_Recent = N88_Summary_Recent['Requirement'].sum()
Corp_N88_Unreal_PNL_Summary_Recent = N88_Summary_Recent['Unreal PNL'].sum()
Corp_N88_Real_PNL_Summary_Recent = N88_Summary_Recent['Real PNL'].sum()

Corp_N88_Cost_Daily_Change = N88_Daily_Change['Cost'].sum()
Corp_N88_Market_Value_Daily_Change = N88_Daily_Change['Market Value'].sum()
Corp_N88_Requirement_Daily_Change = N88_Daily_Change['Requirement'].sum()
Corp_N88_Unreal_PNL_Daily_Change = N88_Daily_Change['Unreal PNL'].sum()
Corp_N88_Real_PNL_Daily_Change = N88_Daily_Change['Real PNL'].sum()


Corp_N90_Cost_Summary_Recent = N90_Summary_Recent['Cost'].sum()
Corp_N90_Market_Value_Summary_Recent = N90_Summary_Recent['Market Value'].sum()
Corp_N90_Requirement_Summary_Recent = N90_Summary_Recent['Requirement'].sum()
Corp_N90_Unreal_PNL_Summary_Recent = N90_Summary_Recent['Unreal PNL'].sum()
Corp_N90_Real_PNL_Summary_Recent = N90_Summary_Recent['Real PNL'].sum()

Corp_N90_Cost_Daily_Change = N90_Daily_Change['Cost'].sum()
Corp_N90_Market_Value_Daily_Change = N90_Daily_Change['Market Value'].sum()
Corp_N90_Requirement_Daily_Change = N90_Daily_Change['Requirement'].sum()
Corp_N90_Unreal_PNL_Daily_Change = N90_Daily_Change['Unreal PNL'].sum()
Corp_N90_Real_PNL_Daily_Change = N90_Daily_Change['Real PNL'].sum()


Corp_P01_Cost_Summary_Recent = P01_Summary_Recent['Cost'].sum()
Corp_P01_Market_Value_Summary_Recent = P01_Summary_Recent['Market Value'].sum()
Corp_P01_Requirement_Summary_Recent = P01_Summary_Recent['Requirement'].sum()
Corp_P01_Unreal_PNL_Summary_Recent = P01_Summary_Recent['Unreal PNL'].sum()
Corp_P01_Real_PNL_Summary_Recent = P01_Summary_Recent['Real PNL'].sum()

Corp_P01_Cost_Daily_Change = P01_Daily_Change['Cost'].sum()
Corp_P01_Market_Value_Daily_Change = P01_Daily_Change['Market Value'].sum()
Corp_P01_Requirement_Daily_Change = P01_Daily_Change['Requirement'].sum()
Corp_P01_Unreal_PNL_Daily_Change = P01_Daily_Change['Unreal PNL'].sum()
Corp_P01_Real_PNL_Daily_Change = P01_Daily_Change['Real PNL'].sum()


Corp_P02_Cost_Summary_Recent = P02_Summary_Recent['Cost'].sum()
Corp_P02_Market_Value_Summary_Recent = P02_Summary_Recent['Market Value'].sum()
Corp_P02_Requirement_Summary_Recent = P02_Summary_Recent['Requirement'].sum()
Corp_P02_Unreal_PNL_Summary_Recent = P02_Summary_Recent['Unreal PNL'].sum()
Corp_P02_Real_PNL_Summary_Recent = P02_Summary_Recent['Real PNL'].sum()

Corp_P02_Cost_Daily_Change = P02_Daily_Change['Cost'].sum()
Corp_P02_Market_Value_Daily_Change = P02_Daily_Change['Market Value'].sum()
Corp_P02_Requirement_Daily_Change = P02_Daily_Change['Requirement'].sum()
Corp_P02_Unreal_PNL_Daily_Change = P02_Daily_Change['Unreal PNL'].sum()
Corp_P02_Real_PNL_Daily_Change = P02_Daily_Change['Real PNL'].sum()


Corp_K74_Cost_Summary_Recent = K74_Summary_Recent['Cost'].sum()
Corp_K74_Market_Value_Summary_Recent = K74_Summary_Recent['Market Value'].sum()
Corp_K74_Requirement_Summary_Recent = K74_Summary_Recent['Requirement'].sum()
Corp_K74_Unreal_PNL_Summary_Recent = K74_Summary_Recent['Unreal PNL'].sum()
Corp_K74_Real_PNL_Summary_Recent = K74_Summary_Recent['Real PNL'].sum()

Corp_K74_Cost_Daily_Change = K74_Daily_Change['Cost'].sum()
Corp_K74_Market_Value_Daily_Change = K74_Daily_Change['Market Value'].sum()
Corp_K74_Requirement_Daily_Change = K74_Daily_Change['Requirement'].sum()
Corp_K74_Unreal_PNL_Daily_Change = K74_Daily_Change['Unreal PNL'].sum()
Corp_K74_Real_PNL_Daily_Change = K74_Daily_Change['Real PNL'].sum()


Corp_L81_Cost_Summary_Recent = L81_Summary_Recent['Cost'].sum()
Corp_L81_Market_Value_Summary_Recent = L81_Summary_Recent['Market Value'].sum()
Corp_L81_Requirement_Summary_Recent = L81_Summary_Recent['Requirement'].sum()
Corp_L81_Unreal_PNL_Summary_Recent = L81_Summary_Recent['Unreal PNL'].sum()
Corp_L81_Real_PNL_Summary_Recent = L81_Summary_Recent['Real PNL'].sum()

Corp_L81_Cost_Daily_Change = L81_Daily_Change['Cost'].sum()
Corp_L81_Market_Value_Daily_Change = L81_Daily_Change['Market Value'].sum()
Corp_L81_Requirement_Daily_Change = L81_Daily_Change['Requirement'].sum()
Corp_L81_Unreal_PNL_Daily_Change = L81_Daily_Change['Unreal PNL'].sum()
Corp_L81_Real_PNL_Daily_Change = L81_Daily_Change['Real PNL'].sum()

Corp_P03_Cost_Summary_Recent = P03_Summary_Recent['Cost'].sum()
Corp_P03_Market_Value_Summary_Recent = P03_Summary_Recent['Market Value'].sum()
Corp_P03_Requirement_Summary_Recent = P03_Summary_Recent['Requirement'].sum()
Corp_P03_Unreal_PNL_Summary_Recent = P03_Summary_Recent['Unreal PNL'].sum()
Corp_P03_Real_PNL_Summary_Recent = P03_Summary_Recent['Real PNL'].sum()

Corp_P03_Cost_Daily_Change = P03_Daily_Change['Cost'].sum()
Corp_P03_Market_Value_Daily_Change = P03_Daily_Change['Market Value'].sum()
Corp_P03_Requirement_Daily_Change = P03_Daily_Change['Requirement'].sum()
Corp_P03_Unreal_PNL_Daily_Change = P03_Daily_Change['Unreal PNL'].sum()
Corp_P03_Real_PNL_Daily_Change = P03_Daily_Change['Real PNL'].sum()

Corp_N87_Cost_Summary_Recent = N87_Summary_Recent['Cost'].sum()
Corp_N87_Market_Value_Summary_Recent = N87_Summary_Recent['Market Value'].sum()
Corp_N87_Requirement_Summary_Recent = N87_Summary_Recent['Requirement'].sum()
Corp_N87_Unreal_PNL_Summary_Recent =N87_Summary_Recent['Unreal PNL'].sum()
Corp_N87_Real_PNL_Summary_Recent = N87_Summary_Recent['Real PNL'].sum()

Corp_N87_Cost_Daily_Change =N87_Daily_Change['Cost'].sum()
Corp_N87_Market_Value_Daily_Change = N87_Daily_Change['Market Value'].sum()
Corp_N87_Requirement_Daily_Change = N87_Daily_Change['Requirement'].sum()
Corp_N87_Unreal_PNL_Daily_Change = N87_Daily_Change['Unreal PNL'].sum()
Corp_N87_Real_PNL_Daily_Change = N87_Daily_Change['Real PNL'].sum()


# Overall Totals
Corp_Total_Cost_Summary = (Corp_N88_Cost_Summary_Recent,
                              Corp_N90_Cost_Summary_Recent,
                              Corp_P01_Cost_Summary_Recent,
                              Corp_P02_Cost_Summary_Recent,
                              Corp_K74_Cost_Summary_Recent,
                              Corp_L81_Cost_Summary_Recent,
                              Corp_P03_Cost_Summary_Recent,
                              Corp_N87_Cost_Summary_Recent)
Corp_Total_Cost_Summary = sum(Corp_Total_Cost_Summary)


Corp_Total_Cost_Daily = (Corp_N88_Cost_Daily_Change,
                              Corp_N90_Cost_Daily_Change,
                              Corp_P01_Cost_Daily_Change,
                              Corp_P02_Cost_Daily_Change,
                              Corp_K74_Cost_Daily_Change,
                              Corp_L81_Cost_Daily_Change,
                        Corp_P03_Cost_Daily_Change,
                        Corp_N87_Cost_Daily_Change)
Corp_Total_Cost_Daily = sum(Corp_Total_Cost_Daily)

Corp_Total_Market_Value_Summary = (Corp_N88_Market_Value_Summary_Recent,
                                      Corp_N90_Market_Value_Summary_Recent,
                                      Corp_P01_Market_Value_Summary_Recent,
                                      Corp_P02_Market_Value_Summary_Recent,
                                      Corp_K74_Market_Value_Summary_Recent,
                                      Corp_L81_Market_Value_Summary_Recent,
                                  Corp_P03_Market_Value_Summary_Recent,
                                  Corp_N87_Market_Value_Summary_Recent)
Corp_Total_Market_Value_Summary = sum(Corp_Total_Market_Value_Summary)

Corp_Total_Market_Value_Daily = (Corp_N88_Market_Value_Daily_Change,
                                    Corp_N90_Market_Value_Daily_Change,
                                    Corp_P01_Market_Value_Daily_Change,
                                    Corp_P02_Market_Value_Daily_Change,
                                    Corp_K74_Market_Value_Daily_Change,
                                    Corp_L81_Market_Value_Daily_Change,
                                Corp_P03_Market_Value_Daily_Change,
                                Corp_N87_Market_Value_Daily_Change)
Corp_Total_Market_Value_Daily = sum(Corp_Total_Market_Value_Daily)

Corp_Total_Requirement_Summary = (Corp_N88_Requirement_Summary_Recent,
                              Corp_N90_Requirement_Summary_Recent,
                              Corp_P01_Requirement_Summary_Recent,
                              Corp_P02_Requirement_Summary_Recent,
                              Corp_K74_Requirement_Summary_Recent,
                              Corp_L81_Requirement_Summary_Recent,
                                 Corp_P03_Requirement_Summary_Recent,
                                 Corp_N87_Requirement_Summary_Recent)
Corp_Total_Requirement_Summary = sum(Corp_Total_Requirement_Summary)

Corp_Total_Requirement_Daily = (Corp_N88_Requirement_Daily_Change,
                              Corp_N90_Requirement_Daily_Change,
                              Corp_P01_Requirement_Daily_Change,
                              Corp_P02_Requirement_Daily_Change,
                              Corp_K74_Requirement_Daily_Change,
                              Corp_L81_Requirement_Daily_Change,
                              Corp_P03_Requirement_Daily_Change,
                              Corp_N87_Requirement_Daily_Change)
Corp_Total_Requirement_Daily = sum(Corp_Total_Requirement_Daily)

Corp_Total_Unreal_PNL_Summary = (Corp_N88_Unreal_PNL_Summary_Recent,
                              Corp_N90_Unreal_PNL_Summary_Recent,
                              Corp_P01_Unreal_PNL_Summary_Recent,
                              Corp_P02_Unreal_PNL_Summary_Recent,
                              Corp_K74_Unreal_PNL_Summary_Recent,
                              Corp_L81_Unreal_PNL_Summary_Recent,
                              Corp_P03_Unreal_PNL_Summary_Recent,
                              Corp_N87_Unreal_PNL_Summary_Recent)
Corp_Total_Unreal_PNL_Summary = sum(Corp_Total_Unreal_PNL_Summary)

Corp_Total_Unreal_PNL_Daily = (Corp_N88_Unreal_PNL_Daily_Change,
                              Corp_N90_Unreal_PNL_Daily_Change,
                              Corp_P01_Unreal_PNL_Daily_Change,
                              Corp_P02_Unreal_PNL_Daily_Change,
                              Corp_K74_Unreal_PNL_Daily_Change,
                              Corp_L81_Unreal_PNL_Daily_Change,
                              Corp_P03_Unreal_PNL_Daily_Change,
                              Corp_N87_Unreal_PNL_Daily_Change)
Corp_Total_Unreal_PNL_Daily = sum(Corp_Total_Unreal_PNL_Daily)


Corp_Total_Real_PNL_Summary = (Corp_N88_Real_PNL_Summary_Recent,
                              Corp_N90_Real_PNL_Summary_Recent,
                              Corp_P01_Real_PNL_Summary_Recent,
                              Corp_P02_Real_PNL_Summary_Recent,
                              Corp_K74_Real_PNL_Summary_Recent,
                              Corp_L81_Real_PNL_Summary_Recent,
                              Corp_P03_Real_PNL_Summary_Recent,
                              Corp_N87_Real_PNL_Summary_Recent)
Corp_Total_Real_PNL_Summary = sum(Corp_Total_Real_PNL_Summary )

Corp_Total_Real_PNL_Daily = (Corp_N88_Real_PNL_Daily_Change,
                              Corp_N90_Real_PNL_Daily_Change,
                              Corp_P01_Real_PNL_Daily_Change,
                              Corp_P02_Real_PNL_Daily_Change,
                              Corp_K74_Real_PNL_Daily_Change,
                              Corp_L81_Real_PNL_Daily_Change,
                              Corp_P03_Real_PNL_Daily_Change,
                              Corp_N87_Real_PNL_Daily_Change)
Corp_Total_Real_PNL_Daily = sum(Corp_Total_Real_PNL_Daily)


# CD

CD_Cost_Summary_Recent = CD_Summary_Recent['Cost'].sum()
CD_Market_Value_Summary_Recent = CD_Summary_Recent['Market Value'].sum()
CD_Requirement_Summary_Recent = CD_Summary_Recent['Requirement'].sum()
CD_Unreal_PNL_Summary_Recent = CD_Summary_Recent['Unreal PNL'].sum()
CD_Real_PNL_Summary_Recent = CD_Summary_Recent['Real PNL'].sum()

CD_Cost_Daily_Change = CD_Daily_Change['Cost'].sum()
CD_Market_Value_Daily_Change = CD_Daily_Change['Market Value'].sum()
CD_Requirement_Daily_Change = CD_Daily_Change['Requirement'].sum()
CD_Unreal_PNL_Daily_Change = CD_Daily_Change['Unreal PNL'].sum()
CD_Real_PNL_Daily_Change = CD_Daily_Change['Real PNL'].sum()


#CMO
CMO_Cost_Summary_Recent = CMO_Summary_Recent['Cost'].sum()
CMO_Market_Value_Summary_Recent = CMO_Summary_Recent['Market Value'].sum()
CMO_Requirement_Summary_Recent = CMO_Summary_Recent['Requirement'].sum()
CMO_Unreal_PNL_Summary_Recent = CMO_Summary_Recent['Unreal PNL'].sum()
CMO_Real_PNL_Summary_Recent = CMO_Summary_Recent['Real PNL'].sum()

CMO_Cost_Daily_Change = CMO_Daily_Change['Cost'].sum()
CMO_Market_Value_Daily_Change = CMO_Daily_Change['Market Value'].sum()
CMO_Requirement_Daily_Change = CMO_Daily_Change['Requirement'].sum()
CMO_Unreal_PNL_Daily_Change = CMO_Daily_Change['Unreal PNL'].sum()
CMO_Real_PNL_Daily_Change = CMO_Daily_Change['Real PNL'].sum()




Firm_Cost_Summary_Total = (Muni_Cost_Summary_Recent,Corp_Total_Cost_Summary,CD_Cost_Summary_Recent,CMO_Cost_Summary_Recent)
Firm_Market_Value_Summary_Total = (Muni_Market_Value_Summary_Recent,Corp_Total_Market_Value_Summary,CD_Market_Value_Summary_Recent,CMO_Market_Value_Summary_Recent )
Firm_Requirement_Summary_Total = (Muni_Requirement_Summary_Recent,Corp_Total_Requirement_Summary,CD_Requirement_Summary_Recent,CMO_Requirement_Summary_Recent)
Firm_Unreal_PNL_Summary_Total = (Muni_Unreal_PNL_Summary_Recent,Corp_Total_Unreal_PNL_Summary,CD_Unreal_PNL_Summary_Recent,CMO_Unreal_PNL_Summary_Recent)
Firm_Real_PNL_Summary_Total = (Muni_Real_PNL_Summary_Recent,Corp_Total_Real_PNL_Summary,CD_Real_PNL_Summary_Recent,CMO_Real_PNL_Summary_Recent)

Firm_Cost_Daily_Total = (Muni_Cost_Daily_Change,Corp_Total_Cost_Daily,CD_Cost_Daily_Change,CMO_Cost_Daily_Change)
Firm_Market_Value_Daily_Total = (Muni_Market_Value_Daily_Change,Corp_Total_Market_Value_Daily,CD_Market_Value_Daily_Change,CMO_Market_Value_Daily_Change)
Firm_Requirement_Daily_Total = (Muni_Requirement_Daily_Change,Corp_Total_Requirement_Daily,CD_Requirement_Daily_Change,CMO_Requirement_Daily_Change)
Firm_Unreal_PNL_Daily_Total = (Muni_Unreal_PNL_Daily_Change,Corp_Total_Unreal_PNL_Daily,CD_Unreal_PNL_Daily_Change,CMO_Unreal_PNL_Daily_Change)
Firm_Real_PNL_Daily_Total = (Muni_Real_PNL_Daily_Change,Corp_Total_Real_PNL_Daily,CD_Real_PNL_Daily_Change,CMO_Real_PNL_Daily_Change)



Firm_Cost_Summary_Total = sum(Firm_Cost_Summary_Total)
Firm_Market_Value_Summary_Total = sum(Firm_Market_Value_Summary_Total)
Firm_Requirement_Summary_Total = sum(Firm_Requirement_Summary_Total)
Firm_Unreal_PNL_Summary_Total = sum(Firm_Unreal_PNL_Summary_Total)
Firm_Real_PNL_Summary_Total = sum(Firm_Real_PNL_Summary_Total)
Firm_Cost_Daily_Total = sum(Firm_Cost_Daily_Total)
Firm_Market_Value_Daily_Total = sum(Firm_Market_Value_Daily_Total)
Firm_Requirement_Daily_Total = sum(Firm_Requirement_Daily_Total)
Firm_Unreal_PNL_Daily_Total = sum(Firm_Unreal_PNL_Daily_Total)
Firm_Real_PNL_Daily_Total = sum(Firm_Real_PNL_Daily_Total)







# Hilltop_Summary_new.to_excel(writer,sheet_name ='Summary',index=True,startrow=1)

Hilltop_Recent_s['Change'] = Hilltop_Recent_s['Total Available Funds']-Hilltop_Old_s['Total Available Funds']
Hilltop_Recent_s = Hilltop_Recent_s[3:]
Hilltop_Recent_s = Hilltop_Recent_s[['Account Number','Total Available Funds','Change']]
Hilltop_Recent_s.to_excel(writer,sheet_name ='Summary',index=False,startrow=1,startcol=30)
workbook = writer.book


format_mini_total = workbook.add_format({'num_format': '#,##0',
                                         'font_size':'8',
                                         'bold': True,
                                         'top':1})
format_general = workbook.add_format({'num_format': '#,##0',
                                         'font_size':'8'})
format_blank_blue = workbook.add_format({'bg_color':'#4267b8',
                                         'font_size':'8',
                                         'font_color':'white'})
format_top_summary = workbook.add_format({'bg_color':'#000e6b',
                                          'font_size':'10',
                                          'font_color':'white'})

format_grey_columnhead = workbook.add_format({'bg_color':'#d4d4d4',
                                              'font_size':'8',
                                              'bottom':1})

format_subtotal = workbook.add_format({'num_format': '#,##0',
                                       'bold':True,
                                       'font_size':'10',
                                       'bottom':2,
                                       'top':1})
format_general_row = workbook.add_format({'font_size':'8',
                                          'num_format': '#,##0'})
format_general_row_green = workbook.add_format({'font_size':'8',
                                                'num_format': '#,##0',
                                                'font_color':'green'})
format_general_row_red = workbook.add_format({'font_size':'8',
                                              'num_format': '#,##0',
                                              'font_color':'red'})
format_subtotal_row = workbook.add_format({'font_size':'10',
                                          'num_format': '#,##0',
                                          'bottom':1,
                                          'top':1})
format_group_total = workbook.add_format({'font_size':'10',
                                          'num_format':'#,##0',
                                          'bold': True})
format_column = workbook.add_format({'bottom':0,
                                     'top':0,
                                     'border_color':'white'})
format_url_links = workbook.add_format({'font_size':'10',
                                       'font_color':'blue',
                                       'underline': 1})

merge_format = workbook.add_format({
    'font_size':'10',
    'bold': 1,
    'border': 0,
    'align': 'left',
    'valign': 'vcenter',
    'fg_color':'#000e6b',
    'font_color':'white'})

worksheet_summary = writer.sheets['Summary']

worksheet_summary.write('D1', 'Summary',format_top_summary) 
worksheet_summary.write('E1', ' ',format_top_summary)
worksheet_summary.write('F1', ' ',format_top_summary)
worksheet_summary.write('D2', 'Item',format_grey_columnhead)
worksheet_summary.write('E2', 'Available Funds',format_grey_columnhead)
worksheet_summary.write('F2', 'Change',format_grey_columnhead)

worksheet_summary.write_formula('D3','=AE3',format_general)
worksheet_summary.write_formula('D4','=AE4',format_general)
worksheet_summary.write_formula('D5','=AE5',format_mini_total)
worksheet_summary.write_formula('D6','=AE6',format_general)
worksheet_summary.write_formula('D7','=AE7',format_general)
worksheet_summary.write_formula('D8','=AE8',format_general)
worksheet_summary.write_formula('D9','=AE9',format_mini_total)
worksheet_summary.write_formula('E3','=AF3',format_general)
worksheet_summary.write_formula('E4','=AF4',format_general)
worksheet_summary.write_formula('E5','=AF5',format_mini_total)
worksheet_summary.write_formula('E6','=AF6',format_general)
worksheet_summary.write_formula('E7','=AF7',format_general)
worksheet_summary.write_formula('E8','=AF8',format_general)
worksheet_summary.write_formula('E9','=AF9',format_mini_total)
worksheet_summary.write_formula('F3','=AG3',format_general)
worksheet_summary.write_formula('F4','=AG4',format_general)
worksheet_summary.write_formula('F5','=AG5',format_mini_total)
worksheet_summary.write_formula('F6','=AG6',format_general)
worksheet_summary.write_formula('F7','=AG7',format_general)
worksheet_summary.write_formula('F8','=AG8',format_general)
worksheet_summary.write_formula('F9','=AG9',format_mini_total)

worksheet_summary.write('A19', 'Muni Total',format_subtotal)
worksheet_summary.write('B19', ' ',format_subtotal)
worksheet_summary.write('I19', 'Muni Total',format_subtotal)
worksheet_summary.write('J19', ' ',format_subtotal)
worksheet_summary.write('C19', Muni_Cost_Summary_Recent,format_subtotal)
worksheet_summary.write('D19', Muni_Market_Value_Summary_Recent,format_subtotal)
worksheet_summary.write('E19', Muni_Requirement_Summary_Recent,format_subtotal)
worksheet_summary.write('F19', Muni_Unreal_PNL_Summary_Recent,format_subtotal)
worksheet_summary.write('G19', Muni_Real_PNL_Summary_Recent,format_subtotal)
worksheet_summary.write('K19', Muni_Cost_Daily_Change,format_subtotal)
worksheet_summary.write('L19', Muni_Market_Value_Daily_Change,format_subtotal)
worksheet_summary.write('M19', Muni_Requirement_Daily_Change,format_subtotal)
worksheet_summary.write('N19', Muni_Unreal_PNL_Daily_Change,format_subtotal)
worksheet_summary.write('O19', Muni_Real_PNL_Daily_Change,format_subtotal)

worksheet_summary.write('A12', 'Account Name',format_grey_columnhead)
worksheet_summary.write('B12', 'Position Type',format_grey_columnhead)
worksheet_summary.write('C12', 'Cost',format_grey_columnhead)
worksheet_summary.write('D12', 'Market Value',format_grey_columnhead)
worksheet_summary.write('E12', 'Requirement',format_grey_columnhead)
worksheet_summary.write('F12', 'Unreal PNL',format_grey_columnhead)
worksheet_summary.write('G12', 'Real PNL',format_grey_columnhead)
worksheet_summary.write('I12', 'Account Name',format_grey_columnhead)
worksheet_summary.write('J12', 'Position Type',format_grey_columnhead)
worksheet_summary.write('K12', 'Cost',format_grey_columnhead)
worksheet_summary.write('L12', 'Market Value',format_grey_columnhead)
worksheet_summary.write('M12', 'Requirement',format_grey_columnhead)
worksheet_summary.write('N12', 'UnrealPNL',format_grey_columnhead)
worksheet_summary.write('O12', 'Real PNL',format_grey_columnhead)

if 'K72 Muni Inv Fl' in Muni_Summary_Recent_Short:
    worksheet_summary.write('B13', '*short*')
else:
    worksheet_summary.write('B13', '     -')
if 'K78 Taxable Mun' in Muni_Summary_Recent_Short:
    worksheet_summary.write('B14', '*short*')
else:
    worksheet_summary.write('B14', '     -')
if 'K79 Cali Tax Ex' in Muni_Summary_Recent_Short:
    worksheet_summary.write('B15', '*short*')
else:
    worksheet_summary.write('B15', '     -')
if 'K80 Muni Tax Ex' in Muni_Summary_Recent_Short:
    worksheet_summary.write('B16', '*short*')
else:
    worksheet_summary.write('B16', '     -')
if 'K81 Muni Tax' in Muni_Summary_Recent_Short:
    worksheet_summary.write('B17', '*short*')
else:
    worksheet_summary.write('B17', '     -')
if 'K82 Tax 0 Muni' in Muni_Summary_Recent_Short:
    worksheet_summary.write('B18', '*short*')
else:
    worksheet_summary.write('B18', '     -')


if 'K72 Muni Inv Fl' in Muni_Daily_Change_Short:
    worksheet_summary.write('J13', '*short*')
else:
    worksheet_summary.write('J13', '     -')
if 'K78 Taxable Mun' in Muni_Daily_Change_Short:
    worksheet_summary.write('J14', '*short*')
else:
    worksheet_summary.write('J14', '     -')
if 'K79 Cali Tax Ex' in Muni_Daily_Change_Short:
    worksheet_summary.write('J15', '*short*')
else:
    worksheet_summary.write('J15', '     -')
if 'K80 Muni Tax Ex' in Muni_Daily_Change_Short:
    worksheet_summary.write('J16', '*short*')
else:
    worksheet_summary.write('J16', '     -')
if 'K81 Muni Tax' in Muni_Daily_Change_Short:
    worksheet_summary.write('J17', '*short*')
else:
    worksheet_summary.write('J17', '     -')
if 'K82 Tax 0 Muni' in Muni_Daily_Change_Short:
    worksheet_summary.write('J18', '*short*')
else:
    worksheet_summary.write('J18', '     -')



worksheet_summary.write('A20', ' ')
worksheet_summary.write('B20', ' ')
worksheet_summary.write('C20', ' ')
worksheet_summary.write('D20', ' ')
worksheet_summary.write('E20', ' ')
worksheet_summary.write('F20', ' ')
worksheet_summary.write('G20', ' ')
worksheet_summary.write('I20', ' ')
worksheet_summary.write('J20', ' ')
worksheet_summary.write('K20', ' ')
worksheet_summary.write('L20', ' ')
worksheet_summary.write('M20', ' ')
worksheet_summary.write('N20', ' ')
worksheet_summary.write('O20', ' ')

worksheet_summary.write('A13', 'K72 MUNI',format_general)
worksheet_summary.write('A14', 'K78 MUNTAX',format_general)
worksheet_summary.write('A15', 'K79 MUNCC',format_general)
worksheet_summary.write('A16', 'K80 MUNBT',format_general)
worksheet_summary.write('A17', 'K81 MUNCCTAX',format_general)
worksheet_summary.write('A18', 'K82 MUNBTTAX',format_general)
worksheet_summary.write('A21', 'N88 CORPIG',format_general)
worksheet_summary.write('A25', 'N90 CORPNOTE',format_general)
worksheet_summary.write('A29', 'P01 CORPFRN',format_general)
worksheet_summary.write('A33', 'P02 CORPSP',format_general)
worksheet_summary.write('A37', 'K74 CORPHY',format_general)
worksheet_summary.write('A41', 'L81 Corp Other',format_general)
worksheet_summary.write('A45', 'P03 CORPDIST',format_general)
worksheet_summary.write('A49', 'N87 CORPXOVR',format_general)
worksheet_summary.write('A56', 'K76 CMO',format_general)
worksheet_summary.write('A57', 'M64 IO',format_general)

worksheet_summary.write('I13', 'K72 MUNI',format_general)
worksheet_summary.write('I14', 'K78 MUNTAX',format_general)
worksheet_summary.write('I15', 'K79 MUNCC',format_general)
worksheet_summary.write('I16', 'K80 MUNBT',format_general)
worksheet_summary.write('I17', 'K81 MUNCCTAX',format_general)
worksheet_summary.write('I18', 'K82 MUNBTTAX',format_general)
worksheet_summary.write('I21', 'N88 CORPIG',format_general)
worksheet_summary.write('I25', 'N90 CORPNOTE',format_general)
worksheet_summary.write('I29', 'P01 CORPFRN',format_general)
worksheet_summary.write('I33', 'P02 CORPSP',format_general)
worksheet_summary.write('I37', 'K74 CORPHY',format_general)
worksheet_summary.write('I41', 'L81 Corp Other',format_general)
worksheet_summary.write('I45', 'P03 CORPDIST',format_general)
worksheet_summary.write('I49', 'N87 CORPXOVR',format_general)
worksheet_summary.write('I56', 'K76 CMO',format_general)
worksheet_summary.write('I57', 'M64 IO',format_general)

worksheet_summary.write('A22', ' ')
worksheet_summary.write('I22', ' ')
worksheet_summary.write('B23', 'Total',format_mini_total)
worksheet_summary.write('J23', 'Total',format_mini_total)
worksheet_summary.write('C23', Corp_N88_Cost_Summary_Recent,format_mini_total)
worksheet_summary.write('D23', Corp_N88_Market_Value_Summary_Recent,format_mini_total)
worksheet_summary.write('E23', Corp_N88_Requirement_Summary_Recent,format_mini_total)
worksheet_summary.write('F23', Corp_N88_Unreal_PNL_Summary_Recent,format_mini_total)
worksheet_summary.write('G23', Corp_N88_Real_PNL_Summary_Recent,format_mini_total)
worksheet_summary.write('I23', ' ')
worksheet_summary.write('K23', Corp_N88_Cost_Daily_Change,format_mini_total)
worksheet_summary.write('L23', Corp_N88_Market_Value_Daily_Change,format_mini_total)
worksheet_summary.write('M23', Corp_N88_Requirement_Daily_Change ,format_mini_total)
worksheet_summary.write('N23', Corp_N88_Unreal_PNL_Daily_Change,format_mini_total)
worksheet_summary.write('O23', Corp_N88_Real_PNL_Daily_Change,format_mini_total)


worksheet_summary.write('A24', ' ')
worksheet_summary.write('B24', ' ')
worksheet_summary.write('C24', ' ')
worksheet_summary.write('D24', ' ')
worksheet_summary.write('E24', ' ')
worksheet_summary.write('F24', ' ')
worksheet_summary.write('G24', ' ')
worksheet_summary.write('I24', ' ')
worksheet_summary.write('J24', ' ')
worksheet_summary.write('K24', ' ')
worksheet_summary.write('L24', ' ')
worksheet_summary.write('M24', ' ')
worksheet_summary.write('N24', ' ')
worksheet_summary.write('O24', ' ')

worksheet_summary.write('I26', ' ')
worksheet_summary.write('A26', ' ')
worksheet_summary.write('B27', 'Total',format_mini_total)
worksheet_summary.write('J27', 'Total',format_mini_total)
worksheet_summary.write('C27', Corp_N90_Cost_Summary_Recent,format_mini_total)
worksheet_summary.write('D27', Corp_N90_Market_Value_Summary_Recent,format_mini_total)
worksheet_summary.write('E27', Corp_N90_Requirement_Summary_Recent,format_mini_total)
worksheet_summary.write('F27', Corp_N90_Unreal_PNL_Summary_Recent,format_mini_total)
worksheet_summary.write('G27', Corp_N90_Real_PNL_Summary_Recent,format_mini_total)
worksheet_summary.write('I27', ' ')
worksheet_summary.write('K27', Corp_N90_Cost_Daily_Change,format_mini_total)
worksheet_summary.write('L27', Corp_N90_Market_Value_Daily_Change,format_mini_total)
worksheet_summary.write('M27', Corp_N90_Requirement_Daily_Change,format_mini_total)
worksheet_summary.write('N27', Corp_N90_Unreal_PNL_Daily_Change,format_mini_total)
worksheet_summary.write('O27', Corp_N90_Real_PNL_Daily_Change,format_mini_total)

worksheet_summary.write('A28', ' ')
worksheet_summary.write('B28', ' ')
worksheet_summary.write('C28', ' ')
worksheet_summary.write('D28', ' ')
worksheet_summary.write('E28', ' ')
worksheet_summary.write('F28', ' ')
worksheet_summary.write('G28', ' ')
worksheet_summary.write('I28', ' ')
worksheet_summary.write('J28', ' ')
worksheet_summary.write('K28', ' ')
worksheet_summary.write('L28', ' ')
worksheet_summary.write('M28', ' ')
worksheet_summary.write('N28', ' ')
worksheet_summary.write('O28', ' ')

worksheet_summary.write('A30', ' ')
worksheet_summary.write('I30', ' ')
worksheet_summary.write('B31', 'Total',format_mini_total)
worksheet_summary.write('J31', 'Total',format_mini_total)
worksheet_summary.write('C31', Corp_P01_Cost_Summary_Recent ,format_mini_total)
worksheet_summary.write('D31', Corp_P01_Market_Value_Summary_Recent,format_mini_total)
worksheet_summary.write('E31', Corp_P01_Requirement_Summary_Recent,format_mini_total)
worksheet_summary.write('F31', Corp_P01_Unreal_PNL_Summary_Recent,format_mini_total)
worksheet_summary.write('G31', Corp_P01_Real_PNL_Summary_Recent,format_mini_total)
worksheet_summary.write('I31', ' ')
worksheet_summary.write('K31', Corp_P01_Cost_Daily_Change,format_mini_total)
worksheet_summary.write('L31', Corp_P01_Market_Value_Daily_Change,format_mini_total)
worksheet_summary.write('M31', Corp_P01_Requirement_Daily_Change,format_mini_total)
worksheet_summary.write('N31', Corp_P01_Unreal_PNL_Daily_Change,format_mini_total)
worksheet_summary.write('O31', Corp_P01_Real_PNL_Daily_Change,format_mini_total)

worksheet_summary.write('A32', ' ')
worksheet_summary.write('B32', ' ')
worksheet_summary.write('C32', ' ')
worksheet_summary.write('D32', ' ')
worksheet_summary.write('E32', ' ')
worksheet_summary.write('F32', ' ')
worksheet_summary.write('G32', ' ')
worksheet_summary.write('I32', ' ')
worksheet_summary.write('J32', ' ')
worksheet_summary.write('K32', ' ')
worksheet_summary.write('L32', ' ')
worksheet_summary.write('M32', ' ')
worksheet_summary.write('N32', ' ')
worksheet_summary.write('O32', ' ')


worksheet_summary.write('A34', ' ')
worksheet_summary.write('I34', ' ')
worksheet_summary.write('B35', 'Total',format_mini_total)
worksheet_summary.write('J35', 'Total',format_mini_total)
worksheet_summary.write('C35', Corp_P02_Cost_Summary_Recent ,format_mini_total)
worksheet_summary.write('D35', Corp_P02_Market_Value_Summary_Recent,format_mini_total)
worksheet_summary.write('E35', Corp_P02_Requirement_Summary_Recent,format_mini_total)
worksheet_summary.write('F35', Corp_P02_Unreal_PNL_Summary_Recent,format_mini_total)
worksheet_summary.write('G35', Corp_P02_Real_PNL_Summary_Recent,format_mini_total)
worksheet_summary.write('I35', ' ')
worksheet_summary.write('K35', Corp_P02_Cost_Daily_Change,format_mini_total)
worksheet_summary.write('L35', Corp_P02_Market_Value_Daily_Change,format_mini_total)
worksheet_summary.write('M35', Corp_P02_Requirement_Daily_Change,format_mini_total)
worksheet_summary.write('N35', Corp_P02_Unreal_PNL_Daily_Change,format_mini_total)
worksheet_summary.write('O35', Corp_P02_Real_PNL_Daily_Change,format_mini_total)


Corp_P02_Cost_Summary_Recent = P02_Summary_Recent['Cost'].sum()
Corp_P02_Market_Value_Summary_Recent = P02_Summary_Recent['Market Value'].sum()
Corp_P02_Requirement_Summary_Recent = P02_Summary_Recent['Requirement'].sum()
Corp_P02_Unreal_PNL_Summary_Recent = P02_Summary_Recent['Unreal PNL'].sum()
Corp_P02_Real_PNL_Summary_Recent = P02_Summary_Recent['Real PNL'].sum()

Corp_P02_Cost_Daily_Change = P02_Daily_Change['Cost'].sum()
Corp_P02_Market_Value_Daily_Change = P02_Daily_Change['Market Value'].sum()
Corp_P02_Requirement_Daily_Change = P02_Daily_Change['Requirement'].sum()
Corp_P02_Unreal_PNL_Daily_Change = P02_Daily_Change['Unreal PNL'].sum()
Corp_P02_Real_PNL_Daily_Change = P02_Daily_Change['Real PNL'].sum()


worksheet_summary.write('A36', ' ')
worksheet_summary.write('B36', ' ')
worksheet_summary.write('C36', ' ')
worksheet_summary.write('D36', ' ')
worksheet_summary.write('E36', ' ')
worksheet_summary.write('F36', ' ')
worksheet_summary.write('G36', ' ')
worksheet_summary.write('I36', ' ')
worksheet_summary.write('J36', ' ')
worksheet_summary.write('K36', ' ')
worksheet_summary.write('L36', ' ')
worksheet_summary.write('M36', ' ')
worksheet_summary.write('N36', ' ')
worksheet_summary.write('O36', ' ')

worksheet_summary.write('A38', ' ')
worksheet_summary.write('I38', ' ')
worksheet_summary.write('B39', 'Total',format_mini_total)
worksheet_summary.write('J39', 'Total',format_mini_total)
worksheet_summary.write('C39', Corp_K74_Cost_Summary_Recent ,format_mini_total)
worksheet_summary.write('D39', Corp_K74_Market_Value_Summary_Recent,format_mini_total)
worksheet_summary.write('E39', Corp_K74_Requirement_Summary_Recent,format_mini_total)
worksheet_summary.write('F39', Corp_K74_Unreal_PNL_Summary_Recent,format_mini_total)
worksheet_summary.write('G39', Corp_K74_Real_PNL_Summary_Recent,format_mini_total)
worksheet_summary.write('I39', ' ')
worksheet_summary.write('K39', Corp_K74_Cost_Daily_Change,format_mini_total)
worksheet_summary.write('L39', Corp_K74_Market_Value_Daily_Change,format_mini_total)
worksheet_summary.write('M39', Corp_K74_Requirement_Daily_Change,format_mini_total)
worksheet_summary.write('N39', Corp_K74_Unreal_PNL_Daily_Change,format_mini_total)
worksheet_summary.write('O39', Corp_K74_Real_PNL_Daily_Change,format_mini_total)

worksheet_summary.write('A42', ' ')
worksheet_summary.write('I42', ' ')
worksheet_summary.write('B43', 'Total',format_mini_total)
worksheet_summary.write('J43', 'Total',format_mini_total)
worksheet_summary.write('C43', Corp_L81_Cost_Summary_Recent ,format_mini_total)
worksheet_summary.write('D43', Corp_L81_Market_Value_Summary_Recent,format_mini_total)
worksheet_summary.write('E43', Corp_L81_Requirement_Summary_Recent,format_mini_total)
worksheet_summary.write('F43', Corp_L81_Unreal_PNL_Summary_Recent,format_mini_total)
worksheet_summary.write('G43', Corp_L81_Real_PNL_Summary_Recent,format_mini_total)
worksheet_summary.write('I43', ' ')
worksheet_summary.write('K43', Corp_L81_Cost_Daily_Change,format_mini_total)
worksheet_summary.write('L43', Corp_L81_Market_Value_Daily_Change,format_mini_total)
worksheet_summary.write('M43', Corp_L81_Requirement_Daily_Change,format_mini_total)
worksheet_summary.write('N43', Corp_L81_Unreal_PNL_Daily_Change,format_mini_total)
worksheet_summary.write('O43', Corp_L81_Real_PNL_Daily_Change,format_mini_total)

worksheet_summary.write('A44', ' ')
worksheet_summary.write('B44', ' ')
worksheet_summary.write('C44', ' ')
worksheet_summary.write('D44', ' ')
worksheet_summary.write('E44', ' ')
worksheet_summary.write('F44', ' ')
worksheet_summary.write('G44', ' ')
worksheet_summary.write('I44', ' ')
worksheet_summary.write('J44', ' ')
worksheet_summary.write('K44', ' ')
worksheet_summary.write('L44', ' ')
worksheet_summary.write('M44', ' ')
worksheet_summary.write('N44', ' ')
worksheet_summary.write('O44', ' ')




worksheet_summary.write('A48', ' ')
worksheet_summary.write('B48', ' ')
worksheet_summary.write('C48', ' ')
worksheet_summary.write('D48', ' ')
worksheet_summary.write('E48', ' ')
worksheet_summary.write('F48', ' ')
worksheet_summary.write('G48', ' ')
worksheet_summary.write('I48', ' ')
worksheet_summary.write('J48', ' ')
worksheet_summary.write('K48', ' ')
worksheet_summary.write('L48', ' ')
worksheet_summary.write('M48', ' ')
worksheet_summary.write('N48', ' ')
worksheet_summary.write('O48', ' ')

worksheet_summary.write('A50', ' ')
worksheet_summary.write('I50', ' ')
worksheet_summary.write('B51', 'Total',format_mini_total)
worksheet_summary.write('J51', 'Total',format_mini_total)
worksheet_summary.write('C51', Corp_N87_Cost_Summary_Recent ,format_mini_total)
worksheet_summary.write('D51', Corp_N87_Market_Value_Summary_Recent,format_mini_total)
worksheet_summary.write('E51', Corp_N87_Requirement_Summary_Recent,format_mini_total)
worksheet_summary.write('F51', Corp_N87_Unreal_PNL_Summary_Recent,format_mini_total)
worksheet_summary.write('G51', Corp_N87_Real_PNL_Summary_Recent,format_mini_total)
worksheet_summary.write('I51', ' ')
worksheet_summary.write('K51', Corp_N87_Cost_Daily_Change,format_mini_total)
worksheet_summary.write('L51', Corp_N87_Market_Value_Daily_Change,format_mini_total)
worksheet_summary.write('M51', Corp_N87_Requirement_Daily_Change,format_mini_total)
worksheet_summary.write('N51', Corp_N87_Unreal_PNL_Daily_Change,format_mini_total)
worksheet_summary.write('O51', Corp_N87_Real_PNL_Daily_Change,format_mini_total)


worksheet_summary.write('A40', ' ')
worksheet_summary.write('B40', ' ')
worksheet_summary.write('C40', ' ')
worksheet_summary.write('D40', ' ')
worksheet_summary.write('E40', ' ')
worksheet_summary.write('F40', ' ')
worksheet_summary.write('G40', ' ')
worksheet_summary.write('I40', ' ')
worksheet_summary.write('J40', ' ')
worksheet_summary.write('K40', ' ')
worksheet_summary.write('L40', ' ')
worksheet_summary.write('M40', ' ')
worksheet_summary.write('N40', ' ')
worksheet_summary.write('O40', ' ')

worksheet_summary.write('A46', ' ')
worksheet_summary.write('I46', ' ')
worksheet_summary.write('B47', 'Total',format_mini_total)
worksheet_summary.write('J47', 'Total',format_mini_total)
worksheet_summary.write('C47', Corp_P03_Cost_Summary_Recent ,format_mini_total)
worksheet_summary.write('D47', Corp_P03_Market_Value_Summary_Recent,format_mini_total)
worksheet_summary.write('E47', Corp_P03_Requirement_Summary_Recent,format_mini_total)
worksheet_summary.write('F47', Corp_P03_Unreal_PNL_Summary_Recent,format_mini_total)
worksheet_summary.write('G47', Corp_P03_Real_PNL_Summary_Recent,format_mini_total)
worksheet_summary.write('I47', ' ')
worksheet_summary.write('K47', Corp_P03_Cost_Daily_Change,format_mini_total)
worksheet_summary.write('L47', Corp_P03_Market_Value_Daily_Change,format_mini_total)
worksheet_summary.write('M47', Corp_P03_Requirement_Daily_Change,format_mini_total)
worksheet_summary.write('N47', Corp_P03_Unreal_PNL_Daily_Change,format_mini_total)
worksheet_summary.write('O47', Corp_P03_Real_PNL_Daily_Change,format_mini_total)


worksheet_summary.write('A52', 'Corp Total',format_subtotal)
worksheet_summary.write('B52', ' ',format_subtotal)
worksheet_summary.write('I52', 'Corp Total',format_subtotal)
worksheet_summary.write('J52', ' ',format_subtotal)
worksheet_summary.write('C52', Corp_Total_Cost_Summary,format_subtotal)
worksheet_summary.write('D52', Corp_Total_Market_Value_Summary,format_subtotal)
worksheet_summary.write('E52', Corp_Total_Requirement_Summary,format_subtotal)
worksheet_summary.write('F52', Corp_Total_Unreal_PNL_Summary,format_subtotal)
worksheet_summary.write('G52', Corp_Total_Real_PNL_Summary ,format_subtotal)
worksheet_summary.write('K52', Corp_Total_Cost_Daily,format_subtotal)
worksheet_summary.write('L52', Corp_Total_Market_Value_Daily ,format_subtotal)
worksheet_summary.write('M52', Corp_Total_Requirement_Daily,format_subtotal)
worksheet_summary.write('N52', Corp_Total_Unreal_PNL_Daily,format_subtotal)
worksheet_summary.write('O52', Corp_Total_Real_PNL_Daily,format_subtotal)

worksheet_summary.write('A53', ' ')
worksheet_summary.write('B53', ' ')
worksheet_summary.write('C53', ' ')
worksheet_summary.write('D53', ' ')
worksheet_summary.write('E53', ' ')
worksheet_summary.write('F53', ' ')
worksheet_summary.write('G53', ' ')
worksheet_summary.write('I53', ' ')
worksheet_summary.write('J53', ' ')
worksheet_summary.write('K53', ' ')
worksheet_summary.write('L53', ' ')
worksheet_summary.write('M53', ' ')
worksheet_summary.write('N53', ' ')
worksheet_summary.write('O53', ' ')


worksheet_summary.write('A54', 'CD Total',format_subtotal)
worksheet_summary.write('B54', ' ',format_subtotal)
worksheet_summary.write('I54', 'CD Total',format_subtotal)
worksheet_summary.write('J54', ' ',format_subtotal)
worksheet_summary.write('C54', CD_Cost_Summary_Recent ,format_subtotal)
worksheet_summary.write('D54', CD_Market_Value_Summary_Recent,format_subtotal)
worksheet_summary.write('E54', CD_Requirement_Summary_Recent,format_subtotal)
worksheet_summary.write('F54', CD_Unreal_PNL_Summary_Recent,format_subtotal)
worksheet_summary.write('G54', CD_Real_PNL_Summary_Recent,format_subtotal)
worksheet_summary.write('K54', CD_Cost_Daily_Change,format_subtotal)
worksheet_summary.write('L54', CD_Market_Value_Daily_Change,format_subtotal)
worksheet_summary.write('M54', CD_Requirement_Daily_Change ,format_subtotal)
worksheet_summary.write('N54', CD_Unreal_PNL_Daily_Change,format_subtotal)
worksheet_summary.write('O54', CD_Real_PNL_Daily_Change,format_subtotal)


worksheet_summary.write('A55', ' ')
worksheet_summary.write('B55', ' ')
worksheet_summary.write('C55', ' ')
worksheet_summary.write('D55', ' ')
worksheet_summary.write('E55', ' ')
worksheet_summary.write('F55', ' ')
worksheet_summary.write('G55', ' ')
worksheet_summary.write('I55', ' ')
worksheet_summary.write('J55', ' ')
worksheet_summary.write('K55', ' ')
worksheet_summary.write('L55', ' ')
worksheet_summary.write('M55', ' ')
worksheet_summary.write('N55', ' ')
worksheet_summary.write('O55', ' ')

worksheet_summary.write('B56', '     -')
worksheet_summary.write('B57', '     -')
worksheet_summary.write('J56', '     -')
worksheet_summary.write('J57', '     -')

worksheet_summary.write('A58', 'CMO Total',format_subtotal)
worksheet_summary.write('B58', ' ',format_subtotal)
worksheet_summary.write('I58', 'CMO Total',format_subtotal)
worksheet_summary.write('J58', ' ',format_subtotal)
worksheet_summary.write('C58', CMO_Cost_Summary_Recent,format_subtotal)
worksheet_summary.write('D58', CMO_Market_Value_Summary_Recent,format_subtotal)
worksheet_summary.write('E58', CMO_Requirement_Summary_Recent,format_subtotal)
worksheet_summary.write('F58', CMO_Unreal_PNL_Summary_Recent,format_subtotal)
worksheet_summary.write('G58', CMO_Real_PNL_Summary_Recent,format_subtotal)
worksheet_summary.write('K58', CMO_Cost_Daily_Change,format_subtotal)
worksheet_summary.write('L58', CMO_Market_Value_Daily_Change,format_subtotal)
worksheet_summary.write('M58', CMO_Requirement_Daily_Change,format_subtotal)
worksheet_summary.write('N58', CMO_Unreal_PNL_Daily_Change,format_subtotal)
worksheet_summary.write('O58', CMO_Real_PNL_Daily_Change,format_subtotal)


worksheet_summary.write('A60', 'Firm Total',format_subtotal)
worksheet_summary.write('B60', ' ',format_subtotal)
worksheet_summary.write('D60', Firm_Market_Value_Summary_Total,format_subtotal)
worksheet_summary.write('E60', Firm_Requirement_Summary_Total,format_subtotal)
worksheet_summary.write('F60', Firm_Unreal_PNL_Summary_Total,format_subtotal)
worksheet_summary.write('G60', Firm_Real_PNL_Summary_Total,format_subtotal)

worksheet_summary.write('I60', 'Firm Total',format_subtotal)
worksheet_summary.write('J60', ' ',format_subtotal)
worksheet_summary.write('L60', Firm_Market_Value_Daily_Total,format_subtotal)
worksheet_summary.write('M60', Firm_Requirement_Daily_Total,format_subtotal)
worksheet_summary.write('N60', Firm_Unreal_PNL_Daily_Total,format_subtotal)
worksheet_summary.write('O60', Firm_Real_PNL_Daily_Total,format_subtotal)


                                      
worksheet_summary.merge_range('A11:G11', 'Month Summary',merge_format)

worksheet_summary.merge_range('I11:O11', 'Daily Change',merge_format)
worksheet_summary.set_row(1,None,format_general_row) 
worksheet_summary.set_row(2,None,format_general_row) 
worksheet_summary.set_row(3,None,format_general_row) 
worksheet_summary.set_row(4,None,format_general_row) 
worksheet_summary.set_row(5,None,format_general_row) 
worksheet_summary.set_row(6,None,format_general_row) 
worksheet_summary.set_row(7,None,format_general_row) 
worksheet_summary.set_row(8,None,format_general_row) 

worksheet_summary.set_row(9,None,format_general_row)
worksheet_summary.set_row(10,None,format_general_row)
worksheet_summary.set_row(12,None,format_general_row)  
worksheet_summary.set_row(13,None,format_general_row)  
worksheet_summary.set_row(14,None,format_general_row)  
worksheet_summary.set_row(15,None,format_general_row) 
worksheet_summary.set_row(16,None,format_general_row) 
worksheet_summary.set_row(17,None,format_general_row) 
worksheet_summary.set_row(18,None,format_general_row) 
worksheet_summary.set_row(19,3,format_general_row) 
worksheet_summary.set_row(20,None,format_general_row) 
worksheet_summary.set_row(21,None,format_general_row) 
worksheet_summary.set_row(22,None,format_general_row) 
worksheet_summary.set_row(23,3,format_general_row) 
worksheet_summary.set_row(24,None,format_general_row) 
worksheet_summary.set_row(25,None,format_general_row) 
worksheet_summary.set_row(26,None,format_general_row) 
worksheet_summary.set_row(27,3,format_general_row) 
worksheet_summary.set_row(28,None,format_general_row) 
worksheet_summary.set_row(29,None,format_general_row) 
worksheet_summary.set_row(30,None,format_general_row) 
worksheet_summary.set_row(31,3,format_general_row) 
worksheet_summary.set_row(32,None,format_general_row) 
worksheet_summary.set_row(33,None,format_general_row) 
worksheet_summary.set_row(34,None,format_general_row) 
worksheet_summary.set_row(35,3,format_general_row) 
worksheet_summary.set_row(36,None,format_general_row) 
worksheet_summary.set_row(37,None,format_general_row) 
worksheet_summary.set_row(38,None,format_general_row) 
worksheet_summary.set_row(39,3,format_general_row) 
worksheet_summary.set_row(40,None,format_general_row) 
worksheet_summary.set_row(41,None,format_general_row) 
worksheet_summary.set_row(42,None,format_general_row) 
worksheet_summary.set_row(43,3,format_general_row) 
worksheet_summary.set_row(44,None,format_general_row) 
worksheet_summary.set_row(45,None,format_general_row) 
worksheet_summary.set_row(46,None,format_general_row) 
worksheet_summary.set_row(47,3,format_general_row) 
worksheet_summary.set_row(48,None,format_general_row) 
worksheet_summary.set_row(49,None,format_general_row) 
worksheet_summary.set_row(50,None,format_general_row) 
worksheet_summary.set_row(51,None,format_general_row) 
worksheet_summary.set_row(52,3) 
worksheet_summary.set_row(53,None,format_general_row) 
worksheet_summary.set_row(54,3) 
worksheet_summary.set_row(55,None,format_general_row)
worksheet_summary.set_row(56,None,format_general_row)
worksheet_summary.set_row(57,None,format_general_row)
worksheet_summary.set_row(58,3)
worksheet_summary.set_row(59,None,format_general_row)
# worksheet_summary.set_row(60,None,format_subtotal_row)

worksheet_summary.set_column('A:A',15,None)
worksheet_summary.set_column('B:B',13,None)
worksheet_summary.set_column('C:C',13,None,{'hidden':True})
worksheet_summary.set_column('D:D',15,None)
worksheet_summary.set_column('E:E',13,None)
worksheet_summary.set_column('F:F',13,None)
worksheet_summary.set_column('G:G',13,None)
worksheet_summary.set_column('H:H',2,None)
worksheet_summary.set_column('I:I',15,None)
worksheet_summary.set_column('J:J',13,None)
worksheet_summary.set_column('K:K',13,None,{'hidden':True})
worksheet_summary.set_column('L:L',13,None)
worksheet_summary.set_column('M:M',13,None)
worksheet_summary.set_column('N:N',13,None)
worksheet_summary.set_column('O:O',13,None)

worksheet_summary.write('A1', 'Page Links',format_top_summary)

worksheet_summary.write_url('A2',"internal:'Quantity Diff'!A1",format_url_links,string = '1. Quantity Diff')
worksheet_summary.write_url('A3',"internal:'PNL Diff'!A1",format_url_links,string = '2. PNL Diff')
worksheet_summary.write_url('A4',"internal:'Adj Unrealized PNL Change'!A1",format_url_links,string = '3. Adj Unrealized PNL Change')
worksheet_summary.write_url('A5',"internal:'Requirement Change'!A1",format_url_links,string = '4. Requirement Change')
worksheet_summary.write_url('A6',"internal:'HT Detail'!A1",format_url_links,string = '5. HT Detail')
worksheet_summary.write_url('A7',"internal:'TW Detail'!A1",format_url_links ,string = '6. TW Detail')

# Conditional Formating
worksheet_summary.conditional_format('M13:M60', {'type':'cell',
                                        'criteria': '<',
                                        'value':    0,
                                        'format':   format_general_row_green})
worksheet_summary.conditional_format('M13:M60', {'type':'cell',
                                        'criteria': '>',
                                        'value':    0,
                                        'format':   format_general_row_red})
worksheet_summary.conditional_format('N13:O60', {'type':'cell',
                                        'criteria': '<',
                                        'value':    0,
                                        'format':   format_general_row_red})
worksheet_summary.conditional_format('N13:O60', {'type':'cell',
                                        'criteria': '>',
                                        'value':    0,
                                        'format':   format_general_row_green})
"""
Create and format Detail sheet
"""



format1 = workbook.add_format({'num_format': '#,##0',
                               'font_size':'9'})
format2 = workbook.add_format({'num_format': '#,##0.00',
                               'font_size':'9'})

worksheet_summary.insert_image('M1', 'P:/1. Individual Folders/Chad/Python Scripts/PNL Report/Logo.png',{'x_scale':.7,'y_scale':.7})

# Format each colum to fit and display data correclty


Hilltop_x = Hilltop_Individual_Summary
Hilltop_y = Hilltop_x


def Summary_Individual_Sheets(Hilltop_Individual_Summary):
    column_summary = ('TW - HT Quantity Discrepancy',
                      'HT Change in Quantity',
                      'Adj Unreal PNL Change',
                      'HT-TW PNL Discrepancy',
                      'Requirement Change')
    for item in column_summary:
        Hilltop_Individual_Summary[item] = Hilltop_Individual_Summary[item]
    Hilltop_QTY_DSP = Hilltop_Individual_Summary[(Hilltop_Individual_Summary['TW - HT Quantity Discrepancy'] != 0)]
    Hilltop_HT_QTY_Change = Hilltop_Individual_Summary[(Hilltop_Individual_Summary['HT Change in Quantity'] != 0)]
    Hilltop_Adj_Unreal_PNL_Change = Hilltop_Individual_Summary[(Hilltop_Individual_Summary['Adj Unreal PNL Change'] != 0)]
    Hilltop_HT_TW_PNL_DSP = Hilltop_Individual_Summary[(Hilltop_Individual_Summary['HT-TW PNL Discrepancy'] != 0)]
    Hilltop_Requirement_Change = Hilltop_Individual_Summary[(Hilltop_Individual_Summary['Requirement Change'] != 0)]
    return Hilltop_QTY_DSP,Hilltop_HT_QTY_Change,Hilltop_Adj_Unreal_PNL_Change, Hilltop_HT_TW_PNL_DSP,Hilltop_Requirement_Change

Individual_Sheets = Summary_Individual_Sheets(Hilltop_Individual_Summary)


Hilltop_QTY_DSP  = Individual_Sheets[0]
Hilltop_QTY_DSP = pd.merge(Hilltop_QTY_DSP, Hilltop_x, on='Cusip', how='left')
Hilltop_QTY_DSP['TW - HT Quantity Discrepancy_x'] = Hilltop_QTY_DSP['TW - HT Quantity Discrepancy_x'].abs()
Hilltop_QTY_DSP.sort_values('TW - HT Quantity Discrepancy_x', axis=0, ascending=False, inplace=True)
Hilltop_QTY_DSP = Hilltop_QTY_DSP[['Security_x','Cusip','Account_x','TW - HT Quantity Discrepancy_y']]
Hilltop_QTY_DSP.rename(columns={'Security_x': 'Security','Account_x':'Account','TW - HT Quantity Discrepancy_y':"QTY DSP"}, inplace=True)
QTY_DSP_Cleared_Positions_Drop = QTY_DSP_Cleared_Positions.drop('Position Notes',axis = 1)
Hilltop_QTY_DSP = Hilltop_QTY_DSP.append(QTY_DSP_Cleared_Positions_Drop)
Hilltop_QTY_DSP.drop_duplicates(subset ='Cusip',keep = False, inplace = True)
Hilltop_QTY_DSP = Hilltop_QTY_DSP[['Security','Account','Cusip','QTY DSP']]
QTY_DSP_Cleared_Positions= QTY_DSP_Cleared_Positions[['Security','Account','Cusip','QTY DSP','Position Notes']]


Hilltop_HT_TW_PNL_DSP = Individual_Sheets[3]
Hilltop_HT_TW_PNL_DSP.sort_values('HT-TW PNL Discrepancy', axis=0, ascending=False, inplace=True)
Hilltop_HT_TW_PNL_DSP = pd.merge(Hilltop_HT_TW_PNL_DSP, Hilltop_x, on='Cusip', how='left')
Hilltop_HT_TW_PNL_DSP['HT-TW PNL Discrepancy_x'] = Hilltop_HT_TW_PNL_DSP['HT-TW PNL Discrepancy_x'].abs()
Hilltop_HT_TW_PNL_DSP.sort_values('HT-TW PNL Discrepancy_x', axis=0, ascending=False, inplace=True)
Hilltop_HT_TW_PNL_DSP = Hilltop_HT_TW_PNL_DSP[['Security_x','Cusip','Account_x','HT-TW PNL Discrepancy_y']]
Hilltop_HT_TW_PNL_DSP_Lower = Hilltop_HT_TW_PNL_DSP[(Hilltop_HT_TW_PNL_DSP['HT-TW PNL Discrepancy_y'] < -10)]
Hilltop_HT_TW_PNL_DSP_Upper = Hilltop_HT_TW_PNL_DSP[(Hilltop_HT_TW_PNL_DSP['HT-TW PNL Discrepancy_y'] > 10)]
Hilltop_Chunks = [Hilltop_HT_TW_PNL_DSP_Upper,Hilltop_HT_TW_PNL_DSP_Lower]
Hilltop_HT_TW_PNL_DSP = pd.concat(Hilltop_Chunks)
Hilltop_HT_TW_PNL_DSP.sort_values(by = 'HT-TW PNL Discrepancy_y',ascending = False)
x = subject_name(file_text)
y = x[3]
PNL_DSP_Date = y[:10]
Hilltop_HT_TW_PNL_DSP['Date'] = PNL_DSP_Date
Hilltop_HT_TW_PNL_DSP = Hilltop_HT_TW_PNL_DSP[['Date','Security_x','Account_x','Cusip','HT-TW PNL Discrepancy_y',   ]]


"""
Pull and Generate new PNL DIFF Position File
"""
PNL_Report_Date = x[2]
PNL_Report_File_Most_Recent = 'P:/2. Corps/PNL_Daily_Report/Reports/PNL_Report.xlsx'
PNL_DSP_Yesterday = pd.read_excel(PNL_Report_File_Most_Recent,sheet_name = 'PNL Diff')  #DSP Items from previous report
Additions_to_Running_PNL_DSP = PNL_DSP_Yesterday[['Date','Security','Account','Cusip','PNL DSP']]  #DSP Items from previous report sorted
Additions_to_Running_PNL_DSP.dropna(inplace = True)
Current_Running_PNL_DSP_Filepath = 'P:/2. Corps/PNL_Daily_Report/Cleared_Position_File/PNL_DSP_History.xlsx' 
Current_Running_PNL_DSP = pd.read_excel(Current_Running_PNL_DSP_Filepath)# reads running PNL DSP file
Current_Running_PNL_DSP = Current_Running_PNL_DSP[Current_Running_PNL_DSP['Previous PNL DSP'] > 5] # filters out 'closed' Positions
Additions_to_Running_PNL_DSP['Previous PNL DSP'] = Additions_to_Running_PNL_DSP['PNL DSP']
Additions_to_Running_PNL_DSP = Additions_to_Running_PNL_DSP[['Account','Cusip','Date','Previous PNL DSP','Security']]
Current_Running_PNL_DSP_List = [Current_Running_PNL_DSP,Additions_to_Running_PNL_DSP]                                  # creates a list to concat the dataframes together
Complete_Running_PNL_DSP = pd.concat(Current_Running_PNL_DSP_List)     # Concatinate the running PNL DSP and the Most recent DSP
Complete_Running_PNL_DSP.drop_duplicates(subset = 'Cusip', keep = 'first', inplace = True)
Complete_Running_PNL_DSP_with_Detail = pd.merge(Complete_Running_PNL_DSP,Hilltop_Recent, on = 'Cusip', how = 'left')
Complete_Running_PNL_DSP_with_Detail['Net PNL DSP'] = Complete_Running_PNL_DSP_with_Detail['Previous PNL DSP'] + Complete_Running_PNL_DSP_with_Detail['HT-TW PNL Discrepancy']
Complete_Running_PNL_DSP_with_Detail = Complete_Running_PNL_DSP_with_Detail[['Date','Security_x','Account_x','Cusip','Previous PNL DSP','HT-TW PNL Discrepancy','Net PNL DSP']]

Complete_Running_PNL_DSP_with_Detail.rename(columns={'Date':'Date',
                                        'Security_x':'Security',
                                        'Account_x':'Account',
                                        'Cusip':'Cusip',
                                        'Previous PNL DSP':'Previous PNL DSP',
                                         'HT-TW PNL Discrepancy':'Current PNL DSP'
                                        },inplace = True)

Hilltop_Adj_Unreal_PNL_Change = Individual_Sheets[2]
Hilltop_Adj_Unreal_PNL_Change = pd.merge(Hilltop_Adj_Unreal_PNL_Change, Hilltop_x, on='Cusip', how='left')
Hilltop_Adj_Unreal_PNL_Change['Adj Unreal PNL Change_x'] = Hilltop_Adj_Unreal_PNL_Change['Adj Unreal PNL Change_x'].abs()
Hilltop_Adj_Unreal_PNL_Change.sort_values('Adj Unreal PNL Change_x', axis=0, ascending=False, inplace=True)
Hilltop_Adj_Unreal_PNL_Change = Hilltop_Adj_Unreal_PNL_Change[['Security_x','Cusip','Account_x','Adj Unreal PNL Change_y']]



Hilltop_Requirement_Change = Individual_Sheets[4]
Hilltop_Requirement_Change = pd.merge(Hilltop_Requirement_Change, Hilltop_x, on='Cusip', how='left')
Hilltop_Requirement_Change = Hilltop_Requirement_Change[['Cusip','Security_x','Account_x','Requirement Change_x','Requirement Change_y']]
Hilltop_Requirement_Change['Requirement Change_x'] = Hilltop_Requirement_Change['Requirement Change_x'].abs()
Hilltop_Requirement_Change.sort_values('Requirement Change_x', axis=0, ascending=False, inplace=True)
Hilltop_Requirement_Change = Hilltop_Requirement_Change[['Security_x','Cusip','Account_x','Requirement Change_y']]


"""
Write to Excel

"""

QTY_DSP_Cleared_Positions.to_excel(writer,sheet_name ='Quantity Diff',index=False,startrow=1,startcol=6)
Hilltop_QTY_DSP.to_excel(writer, sheet_name = 'Quantity Diff', index=False)
worksheet_Hilltop_QTY_DSP = writer.sheets['Quantity Diff']

Hilltop_HT_TW_PNL_DSP.to_excel(writer, sheet_name = 'PNL Diff', index=False)
Complete_Running_PNL_DSP_with_Detail.to_excel(writer, sheet_name = 'PNL Diff', index=False,startrow = 1,startcol=7)
worksheet_Hilltop_HT_TW_PNL_DSP = writer.sheets['PNL Diff']

Hilltop_Adj_Unreal_PNL_Change.to_excel(writer, sheet_name = 'Adj Unrealized PNL Change', index=False)
worksheet_Hilltop_Adj_Unreal_PNL_Change = writer.sheets['Adj Unrealized PNL Change']

Hilltop_Requirement_Change.to_excel(writer, sheet_name = 'Requirement Change', index=False)
worksheet_Hilltop_Requirement_Change = writer.sheets['Requirement Change']


Hilltop_Recent.to_excel(writer, sheet_name='HT Detail', index=False)
worksheet = writer.sheets['HT Detail']

worksheet_Hilltop_QTY_DSP.set_column('A:A', 35,format1)#, format7) #2
worksheet_Hilltop_QTY_DSP.set_column('B:B', 15,format1)#, format2) #2
worksheet_Hilltop_QTY_DSP.set_column('C:C', 12,format1)#, format7) #3
worksheet_Hilltop_QTY_DSP.set_column('D:D', 12,format1)#, format1)#4
worksheet_Hilltop_QTY_DSP.set_column('D:D', 12,format1)#, format1)#4
worksheet_Hilltop_QTY_DSP.set_column('E:E', 30,format1)#, format1)

worksheet_Hilltop_QTY_DSP.write('A1', 'Security',format_top_summary)#,format5)
worksheet_Hilltop_QTY_DSP.write('B1', 'Account',format_top_summary)#,format5)
worksheet_Hilltop_QTY_DSP.write('C1', 'Cusip',format_top_summary)#,format5)
worksheet_Hilltop_QTY_DSP.write('D1', 'QTY DSP',format_top_summary)#,format5)
worksheet_Hilltop_QTY_DSP.write('E1', 'Position Notes',format_top_summary)#,format5)

worksheet_Hilltop_QTY_DSP.write('G2', 'Security',format_top_summary)#,format5)
worksheet_Hilltop_QTY_DSP.write('H2', 'Account',format_top_summary)#,format5)
worksheet_Hilltop_QTY_DSP.write('I2', 'Cusip',format_top_summary)#,format5)
worksheet_Hilltop_QTY_DSP.write('J2', 'QTY DSP',format_top_summary)#,format5)
worksheet_Hilltop_QTY_DSP.write('K2', 'Position Notes',format_top_summary)#,format5)

worksheet_Hilltop_QTY_DSP.merge_range('G1:K1', 'Cleared QTY DSP',merge_format)

worksheet_Hilltop_QTY_DSP.set_column('G:G', 20,format1)#, format1) #2
worksheet_Hilltop_QTY_DSP.set_column('H:H', 12,format1)#, format1) #2
worksheet_Hilltop_QTY_DSP.set_column('I:I', 12,format1)#, format1) #3
worksheet_Hilltop_QTY_DSP.set_column('J:J', 15,format1)#, format1)#4
worksheet_Hilltop_QTY_DSP.set_column('K:K', 30,format1)#, format1)#4

# worksheet_Hilltop_QTY_DSP.freeze_panes(1, 1)
worksheet_Hilltop_QTY_DSP.autofilter('A1:V20000')
# worksheet_Hilltop_QTY_DSP.hide_gridlines(2)
# worksheet_Hilltop_QTY_DSP.protect('welcome123')
worksheet_Hilltop_QTY_DSP.set_zoom(90)
"""
"""
worksheet_Hilltop_Adj_Unreal_PNL_Change.set_column('A:A', 35,format1)#, format7) #2
worksheet_Hilltop_Adj_Unreal_PNL_Change.set_column('B:B', 12,format1)#, format2) #2
worksheet_Hilltop_Adj_Unreal_PNL_Change.set_column('C:C', 15,format1)#, format7) #3
worksheet_Hilltop_Adj_Unreal_PNL_Change.set_column('D:D', 12,format1)#, format1)#4

worksheet_Hilltop_Adj_Unreal_PNL_Change.write('A1', 'Security',format_top_summary)#,format5)
worksheet_Hilltop_Adj_Unreal_PNL_Change.write('B1', 'Cusip',format_top_summary)#,format5)
worksheet_Hilltop_Adj_Unreal_PNL_Change.write('C1', 'Account',format_top_summary)#,format5)
worksheet_Hilltop_Adj_Unreal_PNL_Change.write('D1', 'Adj Unreal PNL Change',format_top_summary)#,format5)
worksheet_Hilltop_Adj_Unreal_PNL_Change.set_zoom(90)
"""
"""
worksheet_Hilltop_HT_TW_PNL_DSP.set_column('A:A', 11,format1)#, format7) #2
worksheet_Hilltop_HT_TW_PNL_DSP.set_column('B:B', 25,format1)#, format2) #2
worksheet_Hilltop_HT_TW_PNL_DSP.set_column('C:C', 13,format1)#, format7) #3
worksheet_Hilltop_HT_TW_PNL_DSP.set_column('D:D', 11,format1)#, format1)#4
worksheet_Hilltop_HT_TW_PNL_DSP.set_column('E:E', 11,format1)#, format1)#4
worksheet_Hilltop_HT_TW_PNL_DSP.set_column('F:F', 15,format1)#, format1)#4

worksheet_Hilltop_HT_TW_PNL_DSP.write('A1', 'Date',format_top_summary)#,format5)
worksheet_Hilltop_HT_TW_PNL_DSP.write('B1', 'Security',format_top_summary)#,format5)
worksheet_Hilltop_HT_TW_PNL_DSP.write('C1', 'Account',format_top_summary)#,format5)
worksheet_Hilltop_HT_TW_PNL_DSP.write('D1', 'Cusip',format_top_summary)#,format5)
worksheet_Hilltop_HT_TW_PNL_DSP.write('E1', 'PNL DSP',format_top_summary)#,format5)
# worksheet_Hilltop_HT_TW_PNL_DSP.write('F1', 'Position Notes',format_top_summary)#,format5)


worksheet_Hilltop_HT_TW_PNL_DSP.write('H2', 'Date',format_top_summary)#,format5)
worksheet_Hilltop_HT_TW_PNL_DSP.write('I2', 'Security',format_top_summary)#,format5)
worksheet_Hilltop_HT_TW_PNL_DSP.write('J2', 'Account',format_top_summary)#,format5)
worksheet_Hilltop_HT_TW_PNL_DSP.write('K2', 'Cusip',format_top_summary)#,format5)
worksheet_Hilltop_HT_TW_PNL_DSP.write('L2', 'Previous PNL',format_top_summary)#,format5)
worksheet_Hilltop_HT_TW_PNL_DSP.write('M2', 'Current PNL',format_top_summary)#,format5)
worksheet_Hilltop_HT_TW_PNL_DSP.write('N2', 'Net PNL',format_top_summary)#,format5)

worksheet_Hilltop_HT_TW_PNL_DSP.merge_range('H1:N1', 'Unresolved PNL DSP',merge_format)

worksheet_Hilltop_HT_TW_PNL_DSP.set_column('G:G', 3,format1)#, format1) #2
worksheet_Hilltop_HT_TW_PNL_DSP.set_column('H:H', 12,format1)#, format1) #2
worksheet_Hilltop_HT_TW_PNL_DSP.set_column('I:I', 12,format1)#, format1) #3
worksheet_Hilltop_HT_TW_PNL_DSP.set_column('J:J', 15,format1)#, format1)#4
worksheet_Hilltop_HT_TW_PNL_DSP.set_column('K:K', 15,format1)#, format1)#4
worksheet_Hilltop_HT_TW_PNL_DSP.set_column('L:L', 15,format1)#, format1)#4
worksheet_Hilltop_HT_TW_PNL_DSP.set_column('M:M', 15,format1)#, format1)#4
worksheet_Hilltop_HT_TW_PNL_DSP.set_column('N:N', 15,format1)#, format1)#4
worksheet_Hilltop_HT_TW_PNL_DSP.set_zoom(90)

"""
"""
worksheet_Hilltop_Requirement_Change.set_column('A:A', 35,format1)#, format7) #2
worksheet_Hilltop_Requirement_Change.set_column('B:B', 15,format1)#, format2) #2
worksheet_Hilltop_Requirement_Change.set_column('C:C', 15,format1)#, format7) #3
worksheet_Hilltop_Requirement_Change.set_column('D:D', 15,format1)#, format1)#4

worksheet_Hilltop_Requirement_Change.write('A1', 'Security',format_top_summary)#,format5)
worksheet_Hilltop_Requirement_Change.write('B1', 'Cusip',format_top_summary)#,format5)
worksheet_Hilltop_Requirement_Change.write('C1', 'Account',format_top_summary)#,format5)
worksheet_Hilltop_Requirement_Change.write('D1', 'Requirement Change',format_top_summary)#,format5)
worksheet_Hilltop_Requirement_Change.set_zoom(90)


#
# worksheet_Hilltop_QTY_DSP.freeze_panes(1, 1)
worksheet_Hilltop_QTY_DSP.autofilter('A1:E20000')
# worksheet_Hilltop_QTY_DSP.hide_gridlines(2)

# worksheet_Hilltop_HT_TW_PNL_DSP.freeze_panes(1, 1)
worksheet_Hilltop_HT_TW_PNL_DSP.autofilter('A1:E20000')
# worksheet_Hilltop_HT_TW_PNL_DSP.hide_gridlines(2)


# worksheet_Hilltop_Adj_Unreal_PNL_Change.freeze_panes(1, 1)
worksheet_Hilltop_Adj_Unreal_PNL_Change.autofilter('A1:D20000')
# worksheet_Hilltop_Adj_Unreal_PNL_Change.hide_gridlines(2)

# worksheet_Hilltop_Requirement_Change.freeze_panes(1, 1)
worksheet_Hilltop_Requirement_Change.autofilter('A1:D20000')
# worksheet_Hilltop_Requirement_Change.hide_gridlines(2)

worksheet_summary.hide_gridlines(2)
# worksheet_TW_Detail = writer.sheets['TW Detail']


worksheet.set_column('A:A', 34,format1)#, format7) #2
worksheet.set_column('B:B', 13,format1)#, format2) #2
worksheet.set_column('C:C', 14,format1)#, format2) #3
worksheet.set_column('D:D', 11,format2)#, format1)#4
worksheet.set_column('E:E', 11,format1)#, format1)#5
worksheet.set_column('F:F', 11,format1)#, format1)#6
worksheet.set_column('G:G', 11,format1)#, format1)#7
worksheet.set_column('H:H', 11,format1)#, format1)#8
worksheet.set_column('I:I', 11,format1)#, format1)#9
worksheet.set_column('J:J', 11,format1)#, format1)#10
worksheet.set_column('K:K', 11,format1)#, format1)#11
worksheet.set_column('L:L', 11,format1)#, format1)#12
worksheet.set_column('M:M', 11,format1)#, format1)#13
worksheet.set_column('N:N', 11,format1)#, format1)#14
worksheet.set_column('O:O', 11,format1)#, format1)#15
worksheet.set_column('P:P', 15,format1)

worksheet.write('A1', 'Security',format_top_summary)#,format5)
worksheet.write('B1', 'Cusip',format_top_summary)#,format5)
worksheet.write('C1', 'Account',format_top_summary)#,format5)
worksheet.write('D1', 'Price',format_top_summary)#,format5)
worksheet.write('E1', 'TW QTY ',format_top_summary)#,format5)
worksheet.write('F1', 'HT QTY',format_top_summary)#,format5)
worksheet.write('G1', 'QTY Discrepancy',format_top_summary)#,format5)
worksheet.write('H1', 'HT QTY Change',format_top_summary)#,format5)
worksheet.write('I1', 'HT New Unreal PNL',format_top_summary)#,format5)
worksheet.write('J1', 'HT Old Unreal PNL',format_top_summary)#,format5)
worksheet.write('K1', 'Real PNL Change',format_top_summary)#,format5)
worksheet.write('L1', 'Adj Unreal PNL Change',format_top_summary)#,format5)
worksheet.write('M1', 'TW PNL',format_top_summary)#,format5)
worksheet.write('N1', 'HT-TW PNL Discrep.',format_top_summary)#,format5)
worksheet.write('O1', 'Req. Change',format_top_summary)#,format5)
worksheet.write('P1', 'Requirement',format_top_summary)#,format5)

worksheet.set_zoom(90)

worksheet.autofilter('A1:O20000')

TW_Detail.reset_index(inplace = True)
TW_Detail.to_excel(writer,sheet_name = 'TW Detail',index = False)
worksheet_TW_Detail = writer.sheets['TW Detail']
worksheet_TW_Detail.write('A1', 'Cusip',format_top_summary)#,format5)
worksheet_TW_Detail.write('B1', 'P&L',format_top_summary)#,format5)
worksheet_TW_Detail.write('C1', 'Security',format_top_summary)#,format5)
worksheet_TW_Detail.write('D1', 'Position',format_top_summary)#,format5)
worksheet_TW_Detail.write('E1', 'Symbol',format_top_summary)#,format5)
worksheet_TW_Detail.write('F1', 'Book',format_top_summary)#,format5)
worksheet_TW_Detail.write('G1', 'MTG Position',format_top_summary)#,format5)
worksheet_TW_Detail.set_column('A:A', 12,format1)#, format7) #2
worksheet_TW_Detail.set_column('B:B', 12,format1)#, format2) #2
worksheet_TW_Detail.set_column('C:C', 25,format1)#, format7) #3
worksheet_TW_Detail.set_column('D:D', 12,format1)#, format1)#4
worksheet_TW_Detail.set_column('E:E', 15,format1)#, format1)#4
worksheet_TW_Detail.set_column('F:F', 12,format1)#, format1)#4
worksheet_TW_Detail.set_column('G:G', 12,format1)#, format1)#4
worksheet_TW_Detail.autofilter('A1:G20000')
worksheet_TW_Detail.set_zoom(90)

workbook.close()


def subject_name(file_text):
    now = datetime.datetime.now()
    weekday = now.weekday()
    if weekday == 0:
        today = now - datetime.timedelta(days=3)
        yesterday = today - datetime.timedelta(days=1)
    elif weekday == 1:
        today = now - datetime.timedelta(days=1)
        yesterday = today - datetime.timedelta(days=3)
    else:
        today = now - datetime.timedelta(days=1)
        yesterday = today - datetime.timedelta(days=1)
    current = today.strftime("%Y-%m-%d")
    current = str(file_text)+str(current)
    PNL_Report_Date = today.strftime("%m.%d.%Y")
    PNL_Report_Date = str(PNL_Report_Date)+'PNL Discrepancy'
    now = datetime.datetime.now()
    PNL_Report_Write_to_Date = now.strftime('%m.%d.%Y')
    PNL_Report_Write_to_Date = str(PNL_Report_Write_to_Date)+'PNL Discrepancy'
    today = file_text + str(today)
    yesterday = yesterday.strftime("%Y-%m-%d")
    yesterday = file_text + str(yesterday)
    return current,yesterday,PNL_Report_Date,PNL_Report_Write_to_Date



file_text = 'Inventory Margin Report for '
x = subject_name(file_text)

"""
Pull and Generate new QTY_DSP_Cleared_Positions file
"""
PNL_Report_Date = x[2]
PNL_Report_File_Most_Recent = 'P:/2. Corps/PNL_Daily_Report/Reports/PNL_Report.xlsx'
Cleared_Yesterday_PNL_Report = pd.read_excel(PNL_Report_File_Most_Recent,sheet_name = 'Quantity Diff')
Cleared_Yesterday_PNL_Report = Cleared_Yesterday_PNL_Report[['Security','Account','Cusip','QTY DSP','Position Notes']]
Cleared_Yesterday_PNL_Report.dropna(inplace = True)
QTY_DSP_Cleared_Positions = 'P:/2. Corps/PNL_Daily_Report/Cleared_Position_File/QTY_DSP_Cleared_Positions.xlsx'
QTY_DSP_Cleared_Positions= pd.read_excel(QTY_DSP_Cleared_Positions,index = False)
QTY_DSP_Cleared_Positions = QTY_DSP_Cleared_Positions.append(Cleared_Yesterday_PNL_Report)

QTY_DSP_Cleared_Positions.drop_duplicates(keep='first',inplace = True)
writer = pd.ExcelWriter('P:/2. Corps/PNL_Daily_Report/Cleared_Position_File/QTY_DSP_Cleared_Positions.xlsx', engine='xlsxwriter')
QTY_DSP_Cleared_Positions.to_excel(writer)
writer.save()






today = x[0]
yesterday = x[1]
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder(6) # "6" refers to the index of a folder - in this case the inbox. You can change that number to reference
messages = inbox.Items
messages.Sort("[ReceivedTime]", True)
message = messages.GetFirst()

i = 0
while i < 20:
    if message.Subject == today:
        try:
            attachments = message.Attachments
            attachment = attachments.Item(1)
            attachment.SaveASFile('P:/2. Corps/PNL_Daily_Report/HT_Files/' + str(today)+'.xlsx') #Saves to the attachment to current folder
            print('HT File Found')
            message = messages.GetNext()

        except:
            message = messages.GetNext()
    i += 1
    message = messages.GetNext()


file_text = 'Report "TW 16 22" is generated.'
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder(6) # "6" refers to the index of a folder - in this case the inbox. You can change that number to reference
messages = inbox.Items
messages.Sort("[ReceivedTime]", True)
message = messages.GetFirst()
i = 0
while i < 10:
   
    if message.Subject == file_text:
        print(message.Subject)
    
        try:
            attachments = message.Attachments
            attachment = attachments.Item(2)
            attachment.SaveASFile('C:\\Users\\ccraig\\Desktop\\PNL Project' + '\\Bloomberg TW'+'.xls') #Saves to the attachment to current folder
            message = messages.GetNext()

        except:
            message = messages.GetNext()
    i += 1
    message = messages.GetNext()
    
Bloomberg_Inventory_16_22 = pd.read_table('C:/Users/ccraig/Desktop/PNL Project/Bloomberg TW.xls')
Bloomberg_Inventory_16_22 = Bloomberg_Inventory_16_22.transpose()
Bloomberg_Inventory_16_22.reset_index(inplace = True)
Bloomberg_Inventory_16_22.drop('index',axis = 1, inplace = True)
Bloomberg_Inventory_16_22.drop(0,axis = 1, inplace = True)
Bloomberg_Inventory_16_22 = Bloomberg_Inventory_16_22.transpose()
Bloomberg_Inventory_16_22 = Bloomberg_Inventory_16_22[:-1]
Bloomberg_Inventory_16_22 = Bloomberg_Inventory_16_22[2:]
Bloomberg_Inventory_16_22 = Bloomberg_Inventory_16_22[0].str.split(',',expand=True)
Bloomberg_Inventory_16_22.rename(columns={0: 'Security'}, inplace=True)
Bloomberg_Inventory_16_22.rename(columns={1: 'P&L'},inplace =True)
Bloomberg_Inventory_16_22.rename(columns={2: 'Cusip'},inplace =True)
Bloomberg_Inventory_16_22.rename(columns={3: 'Position'},inplace =True)
Bloomberg_Inventory_16_22.rename(columns={4: 'Symbol'},inplace =True)
Bloomberg_Inventory_16_22.rename(columns={5: 'Book'},inplace =True)
Bloomberg_Inventory_16_22.rename(columns={6: 'MTG Position'},inplace =True)
Bloomberg_Inventory_16_22 = correct_position_type(Bloomberg_Inventory_16_22)

file_text = 'Report "TW 1 5" is generated.'
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder(6) # "6" refers to the index of a folder - in this case the inbox. You can change that number to reference
messages = inbox.Items
messages.Sort("[ReceivedTime]", True)
message = messages.GetFirst()
i = 0
while i < 10:

    if message.Subject == file_text:
        print(message.Subject)
    
        try:
            attachments = message.Attachments
            attachment = attachments.Item(2)
            attachment.SaveASFile('C:\\Users\\ccraig\\Desktop\\PNL Project' + '\\Bloomberg TW'+'.xls') #Saves to the attachment to current folder
            message = messages.GetNext()

        except:
            message = messages.GetNext()
    i += 1
    message = messages.GetNext()
    
Bloomberg_Inventory_1_5 = pd.read_table('C:/Users/ccraig/Desktop/PNL Project/Bloomberg TW.xls')
Bloomberg_Inventory_1_5 = Bloomberg_Inventory_1_5.transpose()
Bloomberg_Inventory_1_5.reset_index(inplace = True)
Bloomberg_Inventory_1_5.drop('index',axis = 1, inplace = True)
Bloomberg_Inventory_1_5.drop(0,axis = 1, inplace = True)
Bloomberg_Inventory_1_5 = Bloomberg_Inventory_1_5.transpose()
Bloomberg_Inventory_1_5 = Bloomberg_Inventory_1_5[:-1]
Bloomberg_Inventory_1_5 = Bloomberg_Inventory_1_5[2:]
Bloomberg_Inventory_1_5 = Bloomberg_Inventory_1_5[0].str.split(',',expand=True)
Bloomberg_Inventory_1_5.rename(columns={0: 'Security'}, inplace=True)
Bloomberg_Inventory_1_5.rename(columns={1: 'P&L'},inplace =True)
Bloomberg_Inventory_1_5.rename(columns={2: 'Cusip'},inplace =True)
Bloomberg_Inventory_1_5.rename(columns={3: 'Position'},inplace =True)
Bloomberg_Inventory_1_5.rename(columns={4: 'Symbol'},inplace =True)
Bloomberg_Inventory_1_5.rename(columns={5: 'Book'},inplace =True)
Bloomberg_Inventory_1_5.rename(columns={6: 'MTG Position'},inplace =True)
Bloomberg_Inventory_1_5 = correct_position_type(Bloomberg_Inventory_1_5)

file_text = 'Report "TW 6 10" is generated.'
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder(6) # "6" refers to the index of a folder - in this case the inbox. You can change that number to reference
messages = inbox.Items
messages.Sort("[ReceivedTime]", True)
message = messages.GetFirst()
i = 0
while i < 10:
   
    if message.Subject == file_text:
        print(message.Subject)
    
        try:
            attachments = message.Attachments
            attachment = attachments.Item(2)
            attachment.SaveASFile('C:\\Users\\ccraig\\Desktop\\PNL Project' + '\\Bloomberg TW'+'.xls') #Saves to the attachment to current folder
            message = messages.GetNext()

        except:
            message = messages.GetNext()
    i += 1
    message = messages.GetNext()
    
Bloomberg_Inventory_6_10 = pd.read_table('C:/Users/ccraig/Desktop/PNL Project/Bloomberg TW.xls')
Bloomberg_Inventory_6_10 = Bloomberg_Inventory_6_10.transpose()
Bloomberg_Inventory_6_10.reset_index(inplace = True)
Bloomberg_Inventory_6_10.drop('index',axis = 1, inplace = True)
Bloomberg_Inventory_6_10.drop(0,axis = 1, inplace = True)
Bloomberg_Inventory_6_10 = Bloomberg_Inventory_6_10.transpose()
Bloomberg_Inventory_6_10 = Bloomberg_Inventory_6_10[:-1]
Bloomberg_Inventory_6_10 = Bloomberg_Inventory_6_10[2:]
Bloomberg_Inventory_6_10 = Bloomberg_Inventory_6_10[0].str.split(',',expand=True)
Bloomberg_Inventory_6_10.rename(columns={0: 'Security'}, inplace=True)
Bloomberg_Inventory_6_10.rename(columns={1: 'P&L'},inplace =True)
Bloomberg_Inventory_6_10.rename(columns={2: 'Cusip'},inplace =True)
Bloomberg_Inventory_6_10.rename(columns={3: 'Position'},inplace =True)
Bloomberg_Inventory_6_10.rename(columns={4: 'Symbol'},inplace =True)
Bloomberg_Inventory_6_10.rename(columns={5: 'Book'},inplace =True)
Bloomberg_Inventory_6_10.rename(columns={6: 'MTG Position'},inplace =True)
Bloomberg_Inventory_6_10 = correct_position_type(Bloomberg_Inventory_6_10)

file_text = 'Report "TW 11 15" is generated.'
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder(6) # "6" refers to the index of a folder - in this case the inbox. You can change that number to reference
messages = inbox.Items
messages.Sort("[ReceivedTime]", True)
message = messages.GetFirst()
i = 0
while i < 10:
   
    if message.Subject == file_text:
        print(message.Subject)
        try:
            attachments = message.Attachments
            attachment = attachments.Item(2)
            attachment.SaveASFile('C:\\Users\\ccraig\\Desktop\\PNL Project' + '\\Bloomberg TW'+'.xls') #Saves to the attachment to current folder
            message = messages.GetNext()

        except:
            message = messages.GetNext()
    i += 1
    message = messages.GetNext()
    
Bloomberg_Inventory_11_15 = pd.read_table('C:/Users/ccraig/Desktop/PNL Project/Bloomberg TW.xls')
Bloomberg_Inventory_11_15 = Bloomberg_Inventory_11_15.transpose()
Bloomberg_Inventory_11_15.reset_index(inplace = True)
Bloomberg_Inventory_11_15.drop('index',axis = 1, inplace = True)
Bloomberg_Inventory_11_15.drop(0,axis = 1, inplace = True)
Bloomberg_Inventory_11_15 = Bloomberg_Inventory_11_15.transpose()
Bloomberg_Inventory_11_15 = Bloomberg_Inventory_11_15[:-1]
Bloomberg_Inventory_11_15 = Bloomberg_Inventory_11_15[2:]
Bloomberg_Inventory_11_15 = Bloomberg_Inventory_11_15[0].str.split(',',expand=True)
Bloomberg_Inventory_11_15.rename(columns={0: 'Security'}, inplace=True)
Bloomberg_Inventory_11_15.rename(columns={1: 'P&L'},inplace =True)
Bloomberg_Inventory_11_15.rename(columns={2: 'Cusip'},inplace =True)
Bloomberg_Inventory_11_15.rename(columns={3: 'Position'},inplace =True)
Bloomberg_Inventory_11_15.rename(columns={4: 'Symbol'},inplace =True)
Bloomberg_Inventory_11_15.rename(columns={5: 'Book'},inplace =True)
Bloomberg_Inventory_11_15.rename(columns={6: 'MTG Position'},inplace =True)
Bloomberg_Inventory_11_15 = correct_position_type(Bloomberg_Inventory_11_15)

Bloomberg_Inventory = pd.concat([Bloomberg_Inventory_16_22,
                                 Bloomberg_Inventory_1_5,
                                 Bloomberg_Inventory_6_10,
                                 Bloomberg_Inventory_11_15],ignore_index = True)

"""
Read in Excel files from Hilltop and Bloomberg

"""

'P:/2. Corps/PNL_Daily_Report/HT_Files/' + str(today)+'.xlsx'

Bloomberg_Inventory['Cusip'] = Bloomberg_Inventory['Symbol'].str[1:10]
Bloomberg_Inventory = Bloomberg_Inventory[['Cusip', 'P&L', 'Security', 'Position','Symbol','Book','MTG Position']]
Bloomberg_Inventory['Position'] = Bloomberg_Inventory['Position']*1000
Bloomberg_Inventory['MTG Position'] = Bloomberg_Inventory['MTG Position']*1000
# Bloomberg_Inventory['Cusip'] = Bloomberg_Inventory['Cusip'].astype(str)
print(today)
print(yesterday)
Recent = 'P:/2. Corps/PNL_Daily_Report/HT_Files/' + str(today)+'.xlsx'#Recent = 'C:/Users/ccraig/Desktop/PNL Project/'+str(today)+'.xlsx'
Old = 'P:/2. Corps/PNL_Daily_Report/HT_Files/' + str(yesterday)+'.xlsx'
Hilltop_Recent_x = pd.read_excel(io=Recent, sheet_name='Detail')
# Hilltop_Recent_x['Cusip'] = Hilltop_Recent['Cusip'].astype(str)
Hilltop_Old_y = pd.read_excel(io=Old, sheet_name='Detail')
# Hilltop_Old_y['Cusip'] = Hilltop_Old_y['Cusip'].astype(str)
Hilltop_Recent_s = pd.read_excel(io=Recent, sheet_name='Summary')
Hilltop_Old_s = pd.read_excel(io=Old, sheet_name='Summary')
Hilltop_Recent_s = Hilltop_Recent_s.head(10)
Hilltop_Old_s = Hilltop_Old_s.head(10)
Hilltop_Recent_x['Cusip_group_by'] = Hilltop_Recent_x['Cusip']
Hilltop_Recent_x['Cusip_group_by'] = 'C'+ Hilltop_Recent_x['Cusip_group_by']
TW_Detail = Bloomberg_Inventory
Bloomberg_Inventory = Bloomberg_Inventory.groupby(['Cusip']).agg({'P&L':'sum',
                                                                   'Security':'first',
                                                                   'Position':'sum',
                                                                   'Symbol':'first',
                                                                   'Book':'first',
                                                                   'MTG Position':'sum'})
"""
Fix MTG Position
"""
# Bloomberg_Inventory.loc[(Bloomberg_Inventory['Book'] == '8763') | (Bloomberg_Inventory['Book']=='IO'), 'Position'] = 'MTG Position'


Hilltop_Recent = Hilltop_Recent_x.groupby(['Cusip_group_by']).agg({'Quantity':'sum',
                                                                   'Unreal PNL':'sum',
                                                                   'Real PNL':'sum',
                                                                   'Requirement':'sum',
                                                                   'Cusip':'first',
                                                                   'Description':'first',
                                                                   'Price':'mean'})

Hilltop_Old_y['Cusip_group_by'] = Hilltop_Old_y['Cusip']
Hilltop_Old = Hilltop_Old_y.groupby(['Cusip_group_by']).agg({'Quantity':'sum',
                                                             'Unreal PNL':'sum',
                                                             'Real PNL':'sum',
                                                             'Requirement':'sum',
                                                             'Cusip':'first',
                                                             'Description':'first',
                                                             'Price':'mean'})

"""
Merge Hilltop Recent and Old together

"""

Hilltop_Recent = pd.merge(Hilltop_Recent, Hilltop_Old, on='Cusip', how='left')

Hilltop_Recent = pd.merge(Hilltop_Recent, Bloomberg_Inventory, on='Cusip', how='outer')

Hilltop_Recent_x = Hilltop_Recent_x[['Cusip','Account Name']]

Hilltop_Recent = pd.merge(Hilltop_Recent, Hilltop_Recent_x,on='Cusip',how='outer')

Hilltop_Recent = Hilltop_Recent.fillna(0)
Hilltop_Recent.loc[Hilltop_Recent['Security']==0,'Security'] = Hilltop_Recent['Description_x']

"""
Fix Account Names returning 0
"""
Hilltop_Recent.loc[(Hilltop_Recent['Account Name'] == 0) & (Hilltop_Recent['Book']=='8701'), 'Account Name'] = 'K74 Corporates'
Hilltop_Recent.loc[(Hilltop_Recent['Account Name'] == 0) & (Hilltop_Recent['Book']=='CORPSP'), 'Account Name'] = 'P02 Corp SP'
Hilltop_Recent.loc[(Hilltop_Recent['Account Name'] == 0) & (Hilltop_Recent['Book']=='CORPFRN'), 'Account Name'] = 'P01 Corp Floate'
Hilltop_Recent.loc[(Hilltop_Recent['Account Name'] == 0) & (Hilltop_Recent['Book']=='8503'), 'Account Name'] = 'K72 Muni Inv FI'
Hilltop_Recent.loc[(Hilltop_Recent['Account Name'] == 0) & (Hilltop_Recent['Book']=='8763'), 'Account Name'] = 'K76 S P Inv'
Hilltop_Recent.loc[(Hilltop_Recent['Account Name'] == 0) & (Hilltop_Recent['Book']=='8782'), 'Account Name'] = 'K77 CD Inv'
Hilltop_Recent.loc[(Hilltop_Recent['Account Name'] == 0) & (Hilltop_Recent['Book']=='8937'), 'Account Name'] = 'K78 Taxable Mun'
Hilltop_Recent.loc[(Hilltop_Recent['Account Name'] == 0) & (Hilltop_Recent['Book']=='8938'), 'Account Name'] = 'K79 Cali Tax Ex'
Hilltop_Recent.loc[(Hilltop_Recent['Account Name'] == 0) & (Hilltop_Recent['Book']=='8939'), 'Account Name'] = 'K80 Muni Tax Ex'
Hilltop_Recent.loc[(Hilltop_Recent['Account Name'] == 0) & (Hilltop_Recent['Book']=='8940'), 'Account Name'] = 'K81 Muni Tax'
Hilltop_Recent.loc[(Hilltop_Recent['Account Name'] == 0) & (Hilltop_Recent['Book']=='8941'), 'Account Name'] = 'K82 Tax 0 Muni'
Hilltop_Recent.loc[(Hilltop_Recent['Account Name'] == 0) & (Hilltop_Recent['Book']=='CORPIG'), 'Account Name'] = 'N88 Corp Notes'
Hilltop_Recent.loc[(Hilltop_Recent['Account Name'] == 0) & (Hilltop_Recent['Book']=='CORPNOTE'), 'Account Name'] = 'N90 CD'
Hilltop_Recent.loc[(Hilltop_Recent['Account Name'] == 0) & (Hilltop_Recent['Book']=='IO'), 'Account Name'] = 'M64 Sierra MBS'
Hilltop_Recent = Hilltop_Recent[Hilltop_Recent['Book'] != '8659']
Hilltop_Recent = Hilltop_Recent[Hilltop_Recent['Book'] != '8720']
Hilltop_Recent['Position'] = np.where((Hilltop_Recent['Account Name']== 'K76 S P Inv'),Hilltop_Recent['MTG Position'],Hilltop_Recent['Position'])

"""
Calculate nessicary values

"""

Hilltop_Recent['Real_PNL_Change'] = Hilltop_Recent['Real PNL_x'] - Hilltop_Recent['Real PNL_y']
Hilltop_Recent['Real Discrepancy'] = Hilltop_Recent['Real_PNL_Change'] - Hilltop_Recent['P&L']
Hilltop_Recent['Quantity Change'] = Hilltop_Recent['Quantity_x'] - Hilltop_Recent['Quantity_y']

Hilltop_Individual_Book_Summary = Hilltop_Recent

Hilltop_Recent.rename(columns={'Quantity Change': 'HT Quantity Change',
                               'Quantity_x': 'HT New Quantity',
                               'Quantity_y': 'HT Old Quantity',
                               'Quantity_x_y': '-  HT Quantity  =',
                               'Real Discrepancy_y': 'TW vs. HT Real Discrepancy',
                               'Position': 'TW Quantity',
                               'Real PNL_x': 'HT New Real PNL',
                               'Account Name': 'Account',
                               'Quantity_y':'HT Old Quantity',
                               'Unreal PNL_x':'HT New Unreal PNL',
                               'Unreal PNL_y':'HT Old Unreal PNL',
                               'Real PNL_y':'HT Old Real PNL',
                               'P&L':'TW PNL',
                               'Quantity_y':'HT Old Quantity',
                               'Price_x':'Price'}, inplace=True)




# Hilltop_Recent['TW Quantity'] = Hilltop_Recent['TW Quantity'] * 1000

Hilltop_Recent['HT Change in Quantity'] = Hilltop_Recent['HT New Quantity']-Hilltop_Recent['HT Old Quantity']
Hilltop_Recent['TW Quantity'] = pd.to_numeric(Hilltop_Recent['TW Quantity'])
Hilltop_Recent['HT New Quantity'] = pd.to_numeric(Hilltop_Recent['HT New Quantity'])
Hilltop_Recent['TW - HT Quantity Discrepancy'] = Hilltop_Recent['TW Quantity']-Hilltop_Recent['HT New Quantity']
Hilltop_Recent['Real PNL Change'] = Hilltop_Recent['HT New Real PNL']-Hilltop_Recent['HT Old Real PNL']
Hilltop_Recent['HT Real PNL Change'] = Hilltop_Recent['HT New Real PNL']-Hilltop_Recent['HT Old Real PNL']
Hilltop_Recent['Adj Unreal PNL Change'] = Hilltop_Recent['HT New Unreal PNL']-Hilltop_Recent['HT Old Unreal PNL']+Hilltop_Recent['HT Real PNL Change']
Hilltop_Recent['HT-TW PNL Discrepancy'] = Hilltop_Recent['HT Real PNL Change']-Hilltop_Recent['TW PNL']
Hilltop_Recent['Requirement Change'] = Hilltop_Recent['Requirement_x']-Hilltop_Recent['Requirement_y']
Hilltop_Recent['Filter Column'] = Hilltop_Recent['TW - HT Quantity Discrepancy'] + Hilltop_Recent['HT Change in Quantity'] + Hilltop_Recent['Adj Unreal PNL Change'] + Hilltop_Recent['HT-TW PNL Discrepancy'] + Hilltop_Recent['Requirement Change']
# Hilltop_Recent = Hilltop_Recent[(Hilltop_Recent['Filter Column'] != 0)]
Hilltop_Recent = pd.merge(HT_Detail,Hilltop_Recent, on='Cusip', how='left')
Hilltop_Recent = Hilltop_Recent[[
                                 'HT Quantity Change',                         
                                 'Security',                                       #A
                                 'Cusip',                                          #B
                                 'Account',                                        #C
                                 'Price_x',                                          #D
                                 'TW Quantity',                                    #E
                                 'HT New Quantity',                                #F
                                 'TW - HT Quantity Discrepancy',                   #G
                                 'HT Change in Quantity',                          #H
                                 'HT New Unreal PNL',                              #I
                                 'HT Old Unreal PNL',                              #J
                                 'Real PNL Change',                                #K
                                 'Adj Unreal PNL Change',                          #L
                                 'TW PNL',                                         #M
                                 'HT-TW PNL Discrepancy',                          #N
                                 'Requirement Change',
                                 'Requirement_x'
]]
Hilltop_Recent.dropna(thresh = 5,inplace = True)  

"""
# Drop Duplicated Values for Cusip and HT Quantity Change

"""
Hilltop_Recent = Hilltop_Recent.drop_duplicates(['Cusip', 'HT Quantity Change'])


"""
# Set up excel file naming path w/ today's date

"""
time = datetime.datetime.today()
current = time.strftime("%m.%d.%Y")
"""
# Write file to excel

"""
Hilltop_Recent.drop(
    [
        'HT Quantity Change'
    ],
    axis=1, inplace=True)


Hilltop_Individual_Summary = Hilltop_Recent
Hilltop_Recent.sort_values('TW PNL', axis=0, ascending=False, inplace=True)

filepath = 'C:/Users/ccraig/Desktop/New folder/'+ str(current) + ' Inventory Report.xlsx'


writer = pd.ExcelWriter(filepath, engine='xlsxwriter')
"""
            Summary Code

"""
Summary_Recent = pd.read_excel(io=Recent, sheet_name='Summary')
Summary_Old = pd.read_excel(io=Old,sheet_name='Summary')
Summary_Recent = Summary_Recent[12:]
Summary_Old = Summary_Old[12:]
Summary_Recent.rename(columns={'Account Number': 'Inventory Sub Totals',
                                'Total Available Funds':'Account Name',
                                'Unnamed: 2':'Position Type',
                                'Unnamed: 3':'Cost',
                                'Unnamed: 4':'Market Value',
                                'Unnamed: 5':'Unreal PNL',
                                'Unnamed: 6':'Requirement',
                                'Unnamed: 7':'Real PNL'}, inplace=True)
Summary_Recent.reset_index(inplace = True)
Summary_Old.rename(columns={'Account Number': 'Inventory Sub Totals',
                                'Total Available Funds':'Account Name',
                                'Unnamed: 2':'Position Type',
                                'Unnamed: 3':'Cost',
                                'Unnamed: 4':'Market Value',
                                'Unnamed: 5':'Unreal PNL',
                                'Unnamed: 6':'Requirement',
                                'Unnamed: 7':'Real PNL'}, inplace=True)
Summary_Old.reset_index(inplace = True)

Daily_Change=pd.merge(Summary_Recent, Summary_Old, on='index', how='left')

Daily_Change['Cost'] = Daily_Change['Cost_x']-Daily_Change['Cost_y']
Daily_Change['Market Value']=Daily_Change['Market Value_x']-Daily_Change['Market Value_y']
Daily_Change['Requirement']=Daily_Change['Requirement_x']-Daily_Change['Requirement_y']
Daily_Change['Unreal PNL']=Daily_Change['Unreal PNL_x']-Daily_Change['Unreal PNL_y']
Daily_Change['Real PNL']=Daily_Change['Real PNL_x']-Daily_Change['Real PNL_y']
Daily_Change = Daily_Change[['Account Name_x','Position Type_x','Cost','Market Value','Requirement','Unreal PNL','Real PNL']]
Daily_Change.rename(columns={'Account Name_x': 'Account Name',
                                'Position Type_x':'Position Type'
                            }, inplace=True)
Summary_Recent = Summary_Recent[['Account Name','Position Type','Cost','Market Value','Requirement','Unreal PNL','Real PNL']]
Muni_Summary_Recent = Summary_Recent.reindex([0,1,8,9,10,11,12,13,14,15,16,17]) 
Muni_Daily_Change =  Daily_Change.reindex([0,1,8,9,10,11,12,13,14,15,16,17])
Muni_Summary_Recent_Grouped = Muni_Summary_Recent.groupby(['Account Name']).agg({'Position Type':'first',
                                                                                 'Cost':'sum',
                                                                     'Market Value':'sum',
                                                                    'Requirement':'sum',
                                                                    'Unreal PNL':'sum',
                                                                    'Real PNL':'sum'})
Muni_Summary_Recent.reset_index(inplace = True)
Muni_Summary_Recent_Grouped.reset_index(inplace = True)
Muni_Summary_Recent_Short = Muni_Summary_Recent[Muni_Summary_Recent['Position Type'] == 'Short']

Muni_Summary_Recent_Short['Short Total'] = abs(Muni_Summary_Recent_Short['Cost'] + Muni_Summary_Recent_Short['Market Value'] + Muni_Summary_Recent_Short['Requirement'] + Muni_Summary_Recent_Short['Unreal PNL'] + 
                                             Muni_Summary_Recent_Short['Real PNL'])
Muni_Summary_Recent_Short = Muni_Summary_Recent_Short[Muni_Summary_Recent_Short['Position Type'] == 'Short']
Muni_Summary_Recent_Short = Muni_Summary_Recent_Short[Muni_Summary_Recent_Short['Short Total'] > 0]
Muni_Summary_Recent_Short = Muni_Summary_Recent_Short['Account Name'].tolist()


Muni_Daily_Change_Grouped = Muni_Daily_Change.groupby(['Account Name']).agg({'Position Type':'first',
                                                                             'Cost':'sum',
                                                                     'Market Value':'sum',
                                                                    'Requirement':'sum',
                                                                    'Unreal PNL':'sum',
                                                                    'Real PNL':'sum'})
Muni_Daily_Change.reset_index(inplace = True)
Muni_Daily_Change_Grouped.reset_index(inplace = True)
Muni_Daily_Change_Short = Muni_Daily_Change[Muni_Daily_Change['Position Type'] == 'Short']

Muni_Daily_Change_Short['Short Total'] = abs(Muni_Daily_Change_Short['Cost'] + Muni_Daily_Change_Short['Market Value'] + Muni_Daily_Change_Short['Requirement'] + Muni_Daily_Change_Short['Unreal PNL'] + 
                                             Muni_Daily_Change_Short['Real PNL'])
Muni_Daily_Change_Short = Muni_Daily_Change_Short[Muni_Daily_Change_Short['Position Type'] == 'Short']
Muni_Daily_Change_Short = Muni_Daily_Change_Short[Muni_Daily_Change_Short['Short Total'] > 0]
Muni_Daily_Change_Short = Muni_Daily_Change_Short['Account Name'].tolist()

N87_Summary_Recent  = Summary_Recent.reindex([22,23])
N87_Daily_Change  = Daily_Change.reindex([22,23])

N88_Summary_Recent  = Summary_Recent.reindex([24,25])
N88_Daily_Change  = Daily_Change.reindex([24,25])

N90_Summary_Recent  = Summary_Recent.reindex([26,27])
N90_Daily_Change  = Daily_Change.reindex([26,27])

P01_Summary_Recent  = Summary_Recent.reindex([28,29])
P01_Daily_Change  = Daily_Change.reindex([28,29])

K74_Summary_Recent  = Summary_Recent.reindex([2,3])
K74_Daily_Change  = Daily_Change.reindex([2,3])

L81_Summary_Recent  = Summary_Recent.reindex([18,19])
L81_Daily_Change  = Daily_Change.reindex([18,19])

P02_Summary_Recent  = Summary_Recent.reindex([30,31])
P02_Daily_Change  = Daily_Change.reindex([30,31])

P03_Summary_Recent  = Summary_Recent.reindex([32,33])
P03_Daily_Change  = Daily_Change.reindex([32,33])

CD_Summary_Recent  = Summary_Recent.reindex([6])
CD_Daily_Change  = Daily_Change.reindex([6])

CMO_Summary_Recent  = Summary_Recent.reindex([4,20])
CMO_Daily_Change  = Daily_Change.reindex([4,20])





"""
Create and Format Summary sheet

# """
Hilltop_Summary_new = pd.read_excel(io=Recent, sheet_name='Detail')
Hilltop_Summary_new = Hilltop_Summary_new.groupby(['Account Name']).agg({
                                                          'Unreal PNL':'sum',
                                                          'Real PNL':'sum',
                                                          'Requirement':'sum',
                                                          'Cost':'sum',
                                                          'Market Value':'sum'})
Hilltop_Summary_new = Hilltop_Summary_new[['Cost','Market Value','Requirement','Unreal PNL','Real PNL']]

Hilltop_Summary_new.loc['Column_Total'] = Hilltop_Summary_new.sum(numeric_only=True, axis=0)



# Daily_Change.to_excel(writer,sheet_name ='Summary',index=False,startrow=20)

Muni_Summary_Recent_Grouped.to_excel(writer,sheet_name = 'Summary',index = False,startrow =11 )
Muni_Daily_Change_Grouped.to_excel(writer,sheet_name = 'Summary',index = False,startrow =11, startcol = 8 )
N88_Summary_Recent.to_excel(writer,sheet_name = 'Summary',index = False,startrow =19 )
N88_Daily_Change.to_excel(writer,sheet_name = 'Summary',index = False,startrow =19, startcol = 8)
N90_Summary_Recent.to_excel(writer,sheet_name = 'Summary',index = False,startrow =23 )
N90_Daily_Change.to_excel(writer,sheet_name = 'Summary',index = False,startrow =23, startcol = 8 )
P01_Summary_Recent.to_excel(writer,sheet_name = 'Summary',index = False,startrow =27 )
P01_Daily_Change.to_excel(writer,sheet_name = 'Summary',index = False,startrow =27, startcol = 8 )
P02_Summary_Recent.to_excel(writer,sheet_name = 'Summary',index = False,startrow =31 )
P02_Daily_Change.to_excel(writer,sheet_name = 'Summary',index = False,startrow =31, startcol = 8 )
K74_Summary_Recent.to_excel(writer,sheet_name = 'Summary',index = False,startrow =35 )
K74_Daily_Change.to_excel(writer,sheet_name = 'Summary',index = False,startrow =35, startcol = 8 )
L81_Summary_Recent.to_excel(writer,sheet_name = 'Summary',index = False,startrow =39 )
L81_Daily_Change.to_excel(writer,sheet_name = 'Summary',index = False,startrow =39, startcol = 8 )
P03_Summary_Recent.to_excel(writer,sheet_name = 'Summary',index = False,startrow =43 )
P03_Daily_Change.to_excel(writer,sheet_name = 'Summary',index = False,startrow =43, startcol = 8 )
N87_Summary_Recent.to_excel(writer,sheet_name = 'Summary',index = False,startrow =47 )
N87_Daily_Change.to_excel(writer,sheet_name = 'Summary',index = False,startrow =47, startcol = 8 )

CD_Summary_Recent.to_excel(writer,sheet_name = 'Summary',index = False,startrow = 52)
CD_Daily_Change.to_excel(writer,sheet_name = 'Summary',index = False,startrow = 52, startcol = 8)
CMO_Summary_Recent.to_excel(writer,sheet_name = 'Summary',index = False,startrow = 54)
CMO_Daily_Change.to_excel(writer,sheet_name = 'Summary',index = False,startrow = 54, startcol = 8)

"""
Calculate Column Totals
"""
#       MUNI
Muni_Cost_Summary_Recent = Muni_Summary_Recent['Cost'].sum()
Muni_Market_Value_Summary_Recent = Muni_Summary_Recent['Market Value'].sum()
Muni_Requirement_Summary_Recent = Muni_Summary_Recent['Requirement'].sum()
Muni_Unreal_PNL_Summary_Recent = Muni_Summary_Recent['Unreal PNL'].sum()
Muni_Real_PNL_Summary_Recent = Muni_Summary_Recent['Real PNL'].sum()

Muni_Cost_Daily_Change = Muni_Daily_Change['Cost'].sum()
Muni_Market_Value_Daily_Change = Muni_Daily_Change['Market Value'].sum()
Muni_Requirement_Daily_Change = Muni_Daily_Change['Requirement'].sum()
Muni_Unreal_PNL_Daily_Change = Muni_Daily_Change['Unreal PNL'].sum()
Muni_Real_PNL_Daily_Change = Muni_Daily_Change['Real PNL'].sum()

#      Corp
Corp_N88_Cost_Summary_Recent = N88_Summary_Recent['Cost'].sum()
Corp_N88_Market_Value_Summary_Recent = N88_Summary_Recent['Market Value'].sum()
Corp_N88_Requirement_Summary_Recent = N88_Summary_Recent['Requirement'].sum()
Corp_N88_Unreal_PNL_Summary_Recent = N88_Summary_Recent['Unreal PNL'].sum()
Corp_N88_Real_PNL_Summary_Recent = N88_Summary_Recent['Real PNL'].sum()

Corp_N88_Cost_Daily_Change = N88_Daily_Change['Cost'].sum()
Corp_N88_Market_Value_Daily_Change = N88_Daily_Change['Market Value'].sum()
Corp_N88_Requirement_Daily_Change = N88_Daily_Change['Requirement'].sum()
Corp_N88_Unreal_PNL_Daily_Change = N88_Daily_Change['Unreal PNL'].sum()
Corp_N88_Real_PNL_Daily_Change = N88_Daily_Change['Real PNL'].sum()


Corp_N90_Cost_Summary_Recent = N90_Summary_Recent['Cost'].sum()
Corp_N90_Market_Value_Summary_Recent = N90_Summary_Recent['Market Value'].sum()
Corp_N90_Requirement_Summary_Recent = N90_Summary_Recent['Requirement'].sum()
Corp_N90_Unreal_PNL_Summary_Recent = N90_Summary_Recent['Unreal PNL'].sum()
Corp_N90_Real_PNL_Summary_Recent = N90_Summary_Recent['Real PNL'].sum()

Corp_N90_Cost_Daily_Change = N90_Daily_Change['Cost'].sum()
Corp_N90_Market_Value_Daily_Change = N90_Daily_Change['Market Value'].sum()
Corp_N90_Requirement_Daily_Change = N90_Daily_Change['Requirement'].sum()
Corp_N90_Unreal_PNL_Daily_Change = N90_Daily_Change['Unreal PNL'].sum()
Corp_N90_Real_PNL_Daily_Change = N90_Daily_Change['Real PNL'].sum()


Corp_P01_Cost_Summary_Recent = P01_Summary_Recent['Cost'].sum()
Corp_P01_Market_Value_Summary_Recent = P01_Summary_Recent['Market Value'].sum()
Corp_P01_Requirement_Summary_Recent = P01_Summary_Recent['Requirement'].sum()
Corp_P01_Unreal_PNL_Summary_Recent = P01_Summary_Recent['Unreal PNL'].sum()
Corp_P01_Real_PNL_Summary_Recent = P01_Summary_Recent['Real PNL'].sum()

Corp_P01_Cost_Daily_Change = P01_Daily_Change['Cost'].sum()
Corp_P01_Market_Value_Daily_Change = P01_Daily_Change['Market Value'].sum()
Corp_P01_Requirement_Daily_Change = P01_Daily_Change['Requirement'].sum()
Corp_P01_Unreal_PNL_Daily_Change = P01_Daily_Change['Unreal PNL'].sum()
Corp_P01_Real_PNL_Daily_Change = P01_Daily_Change['Real PNL'].sum()


Corp_P02_Cost_Summary_Recent = P02_Summary_Recent['Cost'].sum()
Corp_P02_Market_Value_Summary_Recent = P02_Summary_Recent['Market Value'].sum()
Corp_P02_Requirement_Summary_Recent = P02_Summary_Recent['Requirement'].sum()
Corp_P02_Unreal_PNL_Summary_Recent = P02_Summary_Recent['Unreal PNL'].sum()
Corp_P02_Real_PNL_Summary_Recent = P02_Summary_Recent['Real PNL'].sum()

Corp_P02_Cost_Daily_Change = P02_Daily_Change['Cost'].sum()
Corp_P02_Market_Value_Daily_Change = P02_Daily_Change['Market Value'].sum()
Corp_P02_Requirement_Daily_Change = P02_Daily_Change['Requirement'].sum()
Corp_P02_Unreal_PNL_Daily_Change = P02_Daily_Change['Unreal PNL'].sum()
Corp_P02_Real_PNL_Daily_Change = P02_Daily_Change['Real PNL'].sum()


Corp_K74_Cost_Summary_Recent = K74_Summary_Recent['Cost'].sum()
Corp_K74_Market_Value_Summary_Recent = K74_Summary_Recent['Market Value'].sum()
Corp_K74_Requirement_Summary_Recent = K74_Summary_Recent['Requirement'].sum()
Corp_K74_Unreal_PNL_Summary_Recent = K74_Summary_Recent['Unreal PNL'].sum()
Corp_K74_Real_PNL_Summary_Recent = K74_Summary_Recent['Real PNL'].sum()

Corp_K74_Cost_Daily_Change = K74_Daily_Change['Cost'].sum()
Corp_K74_Market_Value_Daily_Change = K74_Daily_Change['Market Value'].sum()
Corp_K74_Requirement_Daily_Change = K74_Daily_Change['Requirement'].sum()
Corp_K74_Unreal_PNL_Daily_Change = K74_Daily_Change['Unreal PNL'].sum()
Corp_K74_Real_PNL_Daily_Change = K74_Daily_Change['Real PNL'].sum()


Corp_L81_Cost_Summary_Recent = L81_Summary_Recent['Cost'].sum()
Corp_L81_Market_Value_Summary_Recent = L81_Summary_Recent['Market Value'].sum()
Corp_L81_Requirement_Summary_Recent = L81_Summary_Recent['Requirement'].sum()
Corp_L81_Unreal_PNL_Summary_Recent = L81_Summary_Recent['Unreal PNL'].sum()
Corp_L81_Real_PNL_Summary_Recent = L81_Summary_Recent['Real PNL'].sum()

Corp_L81_Cost_Daily_Change = L81_Daily_Change['Cost'].sum()
Corp_L81_Market_Value_Daily_Change = L81_Daily_Change['Market Value'].sum()
Corp_L81_Requirement_Daily_Change = L81_Daily_Change['Requirement'].sum()
Corp_L81_Unreal_PNL_Daily_Change = L81_Daily_Change['Unreal PNL'].sum()
Corp_L81_Real_PNL_Daily_Change = L81_Daily_Change['Real PNL'].sum()

Corp_P03_Cost_Summary_Recent = P03_Summary_Recent['Cost'].sum()
Corp_P03_Market_Value_Summary_Recent = P03_Summary_Recent['Market Value'].sum()
Corp_P03_Requirement_Summary_Recent = P03_Summary_Recent['Requirement'].sum()
Corp_P03_Unreal_PNL_Summary_Recent = P03_Summary_Recent['Unreal PNL'].sum()
Corp_P03_Real_PNL_Summary_Recent = P03_Summary_Recent['Real PNL'].sum()

Corp_P03_Cost_Daily_Change = P03_Daily_Change['Cost'].sum()
Corp_P03_Market_Value_Daily_Change = P03_Daily_Change['Market Value'].sum()
Corp_P03_Requirement_Daily_Change = P03_Daily_Change['Requirement'].sum()
Corp_P03_Unreal_PNL_Daily_Change = P03_Daily_Change['Unreal PNL'].sum()
Corp_P03_Real_PNL_Daily_Change = P03_Daily_Change['Real PNL'].sum()

Corp_N87_Cost_Summary_Recent = N87_Summary_Recent['Cost'].sum()
Corp_N87_Market_Value_Summary_Recent = N87_Summary_Recent['Market Value'].sum()
Corp_N87_Requirement_Summary_Recent = N87_Summary_Recent['Requirement'].sum()
Corp_N87_Unreal_PNL_Summary_Recent =N87_Summary_Recent['Unreal PNL'].sum()
Corp_N87_Real_PNL_Summary_Recent = N87_Summary_Recent['Real PNL'].sum()

Corp_N87_Cost_Daily_Change =N87_Daily_Change['Cost'].sum()
Corp_N87_Market_Value_Daily_Change = N87_Daily_Change['Market Value'].sum()
Corp_N87_Requirement_Daily_Change = N87_Daily_Change['Requirement'].sum()
Corp_N87_Unreal_PNL_Daily_Change = N87_Daily_Change['Unreal PNL'].sum()
Corp_N87_Real_PNL_Daily_Change = N87_Daily_Change['Real PNL'].sum()


# Overall Totals
Corp_Total_Cost_Summary = (Corp_N88_Cost_Summary_Recent,
                              Corp_N90_Cost_Summary_Recent,
                              Corp_P01_Cost_Summary_Recent,
                              Corp_P02_Cost_Summary_Recent,
                              Corp_K74_Cost_Summary_Recent,
                              Corp_L81_Cost_Summary_Recent,
                              Corp_P03_Cost_Summary_Recent,
                              Corp_N87_Cost_Summary_Recent)
Corp_Total_Cost_Summary = sum(Corp_Total_Cost_Summary)


Corp_Total_Cost_Daily = (Corp_N88_Cost_Daily_Change,
                              Corp_N90_Cost_Daily_Change,
                              Corp_P01_Cost_Daily_Change,
                              Corp_P02_Cost_Daily_Change,
                              Corp_K74_Cost_Daily_Change,
                              Corp_L81_Cost_Daily_Change,
                        Corp_P03_Cost_Daily_Change,
                        Corp_N87_Cost_Daily_Change)
Corp_Total_Cost_Daily = sum(Corp_Total_Cost_Daily)

Corp_Total_Market_Value_Summary = (Corp_N88_Market_Value_Summary_Recent,
                                      Corp_N90_Market_Value_Summary_Recent,
                                      Corp_P01_Market_Value_Summary_Recent,
                                      Corp_P02_Market_Value_Summary_Recent,
                                      Corp_K74_Market_Value_Summary_Recent,
                                      Corp_L81_Market_Value_Summary_Recent,
                                  Corp_P03_Market_Value_Summary_Recent,
                                  Corp_N87_Market_Value_Summary_Recent)
Corp_Total_Market_Value_Summary = sum(Corp_Total_Market_Value_Summary)

Corp_Total_Market_Value_Daily = (Corp_N88_Market_Value_Daily_Change,
                                    Corp_N90_Market_Value_Daily_Change,
                                    Corp_P01_Market_Value_Daily_Change,
                                    Corp_P02_Market_Value_Daily_Change,
                                    Corp_K74_Market_Value_Daily_Change,
                                    Corp_L81_Market_Value_Daily_Change,
                                Corp_P03_Market_Value_Daily_Change,
                                Corp_N87_Market_Value_Daily_Change)
Corp_Total_Market_Value_Daily = sum(Corp_Total_Market_Value_Daily)

Corp_Total_Requirement_Summary = (Corp_N88_Requirement_Summary_Recent,
                              Corp_N90_Requirement_Summary_Recent,
                              Corp_P01_Requirement_Summary_Recent,
                              Corp_P02_Requirement_Summary_Recent,
                              Corp_K74_Requirement_Summary_Recent,
                              Corp_L81_Requirement_Summary_Recent,
                                 Corp_P03_Requirement_Summary_Recent,
                                 Corp_N87_Requirement_Summary_Recent)
Corp_Total_Requirement_Summary = sum(Corp_Total_Requirement_Summary)

Corp_Total_Requirement_Daily = (Corp_N88_Requirement_Daily_Change,
                              Corp_N90_Requirement_Daily_Change,
                              Corp_P01_Requirement_Daily_Change,
                              Corp_P02_Requirement_Daily_Change,
                              Corp_K74_Requirement_Daily_Change,
                              Corp_L81_Requirement_Daily_Change,
                              Corp_P03_Requirement_Daily_Change,
                              Corp_N87_Requirement_Daily_Change)
Corp_Total_Requirement_Daily = sum(Corp_Total_Requirement_Daily)

Corp_Total_Unreal_PNL_Summary = (Corp_N88_Unreal_PNL_Summary_Recent,
                              Corp_N90_Unreal_PNL_Summary_Recent,
                              Corp_P01_Unreal_PNL_Summary_Recent,
                              Corp_P02_Unreal_PNL_Summary_Recent,
                              Corp_K74_Unreal_PNL_Summary_Recent,
                              Corp_L81_Unreal_PNL_Summary_Recent,
                              Corp_P03_Unreal_PNL_Summary_Recent,
                              Corp_N87_Unreal_PNL_Summary_Recent)
Corp_Total_Unreal_PNL_Summary = sum(Corp_Total_Unreal_PNL_Summary)

Corp_Total_Unreal_PNL_Daily = (Corp_N88_Unreal_PNL_Daily_Change,
                              Corp_N90_Unreal_PNL_Daily_Change,
                              Corp_P01_Unreal_PNL_Daily_Change,
                              Corp_P02_Unreal_PNL_Daily_Change,
                              Corp_K74_Unreal_PNL_Daily_Change,
                              Corp_L81_Unreal_PNL_Daily_Change,
                              Corp_P03_Unreal_PNL_Daily_Change,
                              Corp_N87_Unreal_PNL_Daily_Change)
Corp_Total_Unreal_PNL_Daily = sum(Corp_Total_Unreal_PNL_Daily)


Corp_Total_Real_PNL_Summary = (Corp_N88_Real_PNL_Summary_Recent,
                              Corp_N90_Real_PNL_Summary_Recent,
                              Corp_P01_Real_PNL_Summary_Recent,
                              Corp_P02_Real_PNL_Summary_Recent,
                              Corp_K74_Real_PNL_Summary_Recent,
                              Corp_L81_Real_PNL_Summary_Recent,
                              Corp_P03_Real_PNL_Summary_Recent,
                              Corp_N87_Real_PNL_Summary_Recent)
Corp_Total_Real_PNL_Summary = sum(Corp_Total_Real_PNL_Summary )

Corp_Total_Real_PNL_Daily = (Corp_N88_Real_PNL_Daily_Change,
                              Corp_N90_Real_PNL_Daily_Change,
                              Corp_P01_Real_PNL_Daily_Change,
                              Corp_P02_Real_PNL_Daily_Change,
                              Corp_K74_Real_PNL_Daily_Change,
                              Corp_L81_Real_PNL_Daily_Change,
                              Corp_P03_Real_PNL_Daily_Change,
                              Corp_N87_Real_PNL_Daily_Change)
Corp_Total_Real_PNL_Daily = sum(Corp_Total_Real_PNL_Daily)


# CD

CD_Cost_Summary_Recent = CD_Summary_Recent['Cost'].sum()
CD_Market_Value_Summary_Recent = CD_Summary_Recent['Market Value'].sum()
CD_Requirement_Summary_Recent = CD_Summary_Recent['Requirement'].sum()
CD_Unreal_PNL_Summary_Recent = CD_Summary_Recent['Unreal PNL'].sum()
CD_Real_PNL_Summary_Recent = CD_Summary_Recent['Real PNL'].sum()

CD_Cost_Daily_Change = CD_Daily_Change['Cost'].sum()
CD_Market_Value_Daily_Change = CD_Daily_Change['Market Value'].sum()
CD_Requirement_Daily_Change = CD_Daily_Change['Requirement'].sum()
CD_Unreal_PNL_Daily_Change = CD_Daily_Change['Unreal PNL'].sum()
CD_Real_PNL_Daily_Change = CD_Daily_Change['Real PNL'].sum()


#CMO
CMO_Cost_Summary_Recent = CMO_Summary_Recent['Cost'].sum()
CMO_Market_Value_Summary_Recent = CMO_Summary_Recent['Market Value'].sum()
CMO_Requirement_Summary_Recent = CMO_Summary_Recent['Requirement'].sum()
CMO_Unreal_PNL_Summary_Recent = CMO_Summary_Recent['Unreal PNL'].sum()
CMO_Real_PNL_Summary_Recent = CMO_Summary_Recent['Real PNL'].sum()

CMO_Cost_Daily_Change = CMO_Daily_Change['Cost'].sum()
CMO_Market_Value_Daily_Change = CMO_Daily_Change['Market Value'].sum()
CMO_Requirement_Daily_Change = CMO_Daily_Change['Requirement'].sum()
CMO_Unreal_PNL_Daily_Change = CMO_Daily_Change['Unreal PNL'].sum()
CMO_Real_PNL_Daily_Change = CMO_Daily_Change['Real PNL'].sum()




Firm_Cost_Summary_Total = (Muni_Cost_Summary_Recent,Corp_Total_Cost_Summary,CD_Cost_Summary_Recent,CMO_Cost_Summary_Recent)
Firm_Market_Value_Summary_Total = (Muni_Market_Value_Summary_Recent,Corp_Total_Market_Value_Summary,CD_Market_Value_Summary_Recent,CMO_Market_Value_Summary_Recent )
Firm_Requirement_Summary_Total = (Muni_Requirement_Summary_Recent,Corp_Total_Requirement_Summary,CD_Requirement_Summary_Recent,CMO_Requirement_Summary_Recent)
Firm_Unreal_PNL_Summary_Total = (Muni_Unreal_PNL_Summary_Recent,Corp_Total_Unreal_PNL_Summary,CD_Unreal_PNL_Summary_Recent,CMO_Unreal_PNL_Summary_Recent)
Firm_Real_PNL_Summary_Total = (Muni_Real_PNL_Summary_Recent,Corp_Total_Real_PNL_Summary,CD_Real_PNL_Summary_Recent,CMO_Real_PNL_Summary_Recent)

Firm_Cost_Daily_Total = (Muni_Cost_Daily_Change,Corp_Total_Cost_Daily,CD_Cost_Daily_Change,CMO_Cost_Daily_Change)
Firm_Market_Value_Daily_Total = (Muni_Market_Value_Daily_Change,Corp_Total_Market_Value_Daily,CD_Market_Value_Daily_Change,CMO_Market_Value_Daily_Change)
Firm_Requirement_Daily_Total = (Muni_Requirement_Daily_Change,Corp_Total_Requirement_Daily,CD_Requirement_Daily_Change,CMO_Requirement_Daily_Change)
Firm_Unreal_PNL_Daily_Total = (Muni_Unreal_PNL_Daily_Change,Corp_Total_Unreal_PNL_Daily,CD_Unreal_PNL_Daily_Change,CMO_Unreal_PNL_Daily_Change)
Firm_Real_PNL_Daily_Total = (Muni_Real_PNL_Daily_Change,Corp_Total_Real_PNL_Daily,CD_Real_PNL_Daily_Change,CMO_Real_PNL_Daily_Change)



Firm_Cost_Summary_Total = sum(Firm_Cost_Summary_Total)
Firm_Market_Value_Summary_Total = sum(Firm_Market_Value_Summary_Total)
Firm_Requirement_Summary_Total = sum(Firm_Requirement_Summary_Total)
Firm_Unreal_PNL_Summary_Total = sum(Firm_Unreal_PNL_Summary_Total)
Firm_Real_PNL_Summary_Total = sum(Firm_Real_PNL_Summary_Total)
Firm_Cost_Daily_Total = sum(Firm_Cost_Daily_Total)
Firm_Market_Value_Daily_Total = sum(Firm_Market_Value_Daily_Total)
Firm_Requirement_Daily_Total = sum(Firm_Requirement_Daily_Total)
Firm_Unreal_PNL_Daily_Total = sum(Firm_Unreal_PNL_Daily_Total)
Firm_Real_PNL_Daily_Total = sum(Firm_Real_PNL_Daily_Total)







# Hilltop_Summary_new.to_excel(writer,sheet_name ='Summary',index=True,startrow=1)

Hilltop_Recent_s['Change'] = Hilltop_Recent_s['Total Available Funds']-Hilltop_Old_s['Total Available Funds']
Hilltop_Recent_s = Hilltop_Recent_s[3:]
Hilltop_Recent_s = Hilltop_Recent_s[['Account Number','Total Available Funds','Change']]
Hilltop_Recent_s.to_excel(writer,sheet_name ='Summary',index=False,startrow=1,startcol=30)
workbook = writer.book


format_mini_total = workbook.add_format({'num_format': '#,##0',
                                         'font_size':'8',
                                         'bold': True,
                                         'top':1})
format_general = workbook.add_format({'num_format': '#,##0',
                                         'font_size':'8'})
format_blank_blue = workbook.add_format({'bg_color':'#4267b8',
                                         'font_size':'8',
                                         'font_color':'white'})
format_top_summary = workbook.add_format({'bg_color':'#000e6b',
                                          'font_size':'10',
                                          'font_color':'white'})

format_grey_columnhead = workbook.add_format({'bg_color':'#d4d4d4',
                                              'font_size':'8',
                                              'bottom':1})

format_subtotal = workbook.add_format({'num_format': '#,##0',
                                       'bold':True,
                                       'font_size':'10',
                                       'bottom':2,
                                       'top':1})
format_general_row = workbook.add_format({'font_size':'8',
                                          'num_format': '#,##0'})
format_general_row_green = workbook.add_format({'font_size':'8',
                                                'num_format': '#,##0',
                                                'font_color':'green'})
format_general_row_red = workbook.add_format({'font_size':'8',
                                              'num_format': '#,##0',
                                              'font_color':'red'})
format_subtotal_row = workbook.add_format({'font_size':'10',
                                          'num_format': '#,##0',
                                          'bottom':1,
                                          'top':1})
format_group_total = workbook.add_format({'font_size':'10',
                                          'num_format':'#,##0',
                                          'bold': True})
format_column = workbook.add_format({'bottom':0,
                                     'top':0,
                                     'border_color':'white'})
format_url_links = workbook.add_format({'font_size':'10',
                                       'font_color':'blue',
                                       'underline': 1})

merge_format = workbook.add_format({
    'font_size':'10',
    'bold': 1,
    'border': 0,
    'align': 'left',
    'valign': 'vcenter',
    'fg_color':'#000e6b',
    'font_color':'white'})

worksheet_summary = writer.sheets['Summary']

worksheet_summary.write('D1', 'Summary',format_top_summary) 
worksheet_summary.write('E1', ' ',format_top_summary)
worksheet_summary.write('F1', ' ',format_top_summary)
worksheet_summary.write('D2', 'Item',format_grey_columnhead)
worksheet_summary.write('E2', 'Available Funds',format_grey_columnhead)
worksheet_summary.write('F2', 'Change',format_grey_columnhead)

worksheet_summary.write_formula('D3','=AE3',format_general)
worksheet_summary.write_formula('D4','=AE4',format_general)
worksheet_summary.write_formula('D5','=AE5',format_mini_total)
worksheet_summary.write_formula('D6','=AE6',format_general)
worksheet_summary.write_formula('D7','=AE7',format_general)
worksheet_summary.write_formula('D8','=AE8',format_general)
worksheet_summary.write_formula('D9','=AE9',format_mini_total)
worksheet_summary.write_formula('E3','=AF3',format_general)
worksheet_summary.write_formula('E4','=AF4',format_general)
worksheet_summary.write_formula('E5','=AF5',format_mini_total)
worksheet_summary.write_formula('E6','=AF6',format_general)
worksheet_summary.write_formula('E7','=AF7',format_general)
worksheet_summary.write_formula('E8','=AF8',format_general)
worksheet_summary.write_formula('E9','=AF9',format_mini_total)
worksheet_summary.write_formula('F3','=AG3',format_general)
worksheet_summary.write_formula('F4','=AG4',format_general)
worksheet_summary.write_formula('F5','=AG5',format_mini_total)
worksheet_summary.write_formula('F6','=AG6',format_general)
worksheet_summary.write_formula('F7','=AG7',format_general)
worksheet_summary.write_formula('F8','=AG8',format_general)
worksheet_summary.write_formula('F9','=AG9',format_mini_total)

worksheet_summary.write('A19', 'Muni Total',format_subtotal)
worksheet_summary.write('B19', ' ',format_subtotal)
worksheet_summary.write('I19', 'Muni Total',format_subtotal)
worksheet_summary.write('J19', ' ',format_subtotal)
worksheet_summary.write('C19', Muni_Cost_Summary_Recent,format_subtotal)
worksheet_summary.write('D19', Muni_Market_Value_Summary_Recent,format_subtotal)
worksheet_summary.write('E19', Muni_Requirement_Summary_Recent,format_subtotal)
worksheet_summary.write('F19', Muni_Unreal_PNL_Summary_Recent,format_subtotal)
worksheet_summary.write('G19', Muni_Real_PNL_Summary_Recent,format_subtotal)
worksheet_summary.write('K19', Muni_Cost_Daily_Change,format_subtotal)
worksheet_summary.write('L19', Muni_Market_Value_Daily_Change,format_subtotal)
worksheet_summary.write('M19', Muni_Requirement_Daily_Change,format_subtotal)
worksheet_summary.write('N19', Muni_Unreal_PNL_Daily_Change,format_subtotal)
worksheet_summary.write('O19', Muni_Real_PNL_Daily_Change,format_subtotal)

worksheet_summary.write('A12', 'Account Name',format_grey_columnhead)
worksheet_summary.write('B12', 'Position Type',format_grey_columnhead)
worksheet_summary.write('C12', 'Cost',format_grey_columnhead)
worksheet_summary.write('D12', 'Market Value',format_grey_columnhead)
worksheet_summary.write('E12', 'Requirement',format_grey_columnhead)
worksheet_summary.write('F12', 'Unreal PNL',format_grey_columnhead)
worksheet_summary.write('G12', 'Real PNL',format_grey_columnhead)
worksheet_summary.write('I12', 'Account Name',format_grey_columnhead)
worksheet_summary.write('J12', 'Position Type',format_grey_columnhead)
worksheet_summary.write('K12', 'Cost',format_grey_columnhead)
worksheet_summary.write('L12', 'Market Value',format_grey_columnhead)
worksheet_summary.write('M12', 'Requirement',format_grey_columnhead)
worksheet_summary.write('N12', 'UnrealPNL',format_grey_columnhead)
worksheet_summary.write('O12', 'Real PNL',format_grey_columnhead)

if 'K72 Muni Inv Fl' in Muni_Summary_Recent_Short:
    worksheet_summary.write('B13', '*short*')
else:
    worksheet_summary.write('B13', '     -')
if 'K78 Taxable Mun' in Muni_Summary_Recent_Short:
    worksheet_summary.write('B14', '*short*')
else:
    worksheet_summary.write('B14', '     -')
if 'K79 Cali Tax Ex' in Muni_Summary_Recent_Short:
    worksheet_summary.write('B15', '*short*')
else:
    worksheet_summary.write('B15', '     -')
if 'K80 Muni Tax Ex' in Muni_Summary_Recent_Short:
    worksheet_summary.write('B16', '*short*')
else:
    worksheet_summary.write('B16', '     -')
if 'K81 Muni Tax' in Muni_Summary_Recent_Short:
    worksheet_summary.write('B17', '*short*')
else:
    worksheet_summary.write('B17', '     -')
if 'K82 Tax 0 Muni' in Muni_Summary_Recent_Short:
    worksheet_summary.write('B18', '*short*')
else:
    worksheet_summary.write('B18', '     -')


if 'K72 Muni Inv Fl' in Muni_Daily_Change_Short:
    worksheet_summary.write('J13', '*short*')
else:
    worksheet_summary.write('J13', '     -')
if 'K78 Taxable Mun' in Muni_Daily_Change_Short:
    worksheet_summary.write('J14', '*short*')
else:
    worksheet_summary.write('J14', '     -')
if 'K79 Cali Tax Ex' in Muni_Daily_Change_Short:
    worksheet_summary.write('J15', '*short*')
else:
    worksheet_summary.write('J15', '     -')
if 'K80 Muni Tax Ex' in Muni_Daily_Change_Short:
    worksheet_summary.write('J16', '*short*')
else:
    worksheet_summary.write('J16', '     -')
if 'K81 Muni Tax' in Muni_Daily_Change_Short:
    worksheet_summary.write('J17', '*short*')
else:
    worksheet_summary.write('J17', '     -')
if 'K82 Tax 0 Muni' in Muni_Daily_Change_Short:
    worksheet_summary.write('J18', '*short*')
else:
    worksheet_summary.write('J18', '     -')



worksheet_summary.write('A20', ' ')
worksheet_summary.write('B20', ' ')
worksheet_summary.write('C20', ' ')
worksheet_summary.write('D20', ' ')
worksheet_summary.write('E20', ' ')
worksheet_summary.write('F20', ' ')
worksheet_summary.write('G20', ' ')
worksheet_summary.write('I20', ' ')
worksheet_summary.write('J20', ' ')
worksheet_summary.write('K20', ' ')
worksheet_summary.write('L20', ' ')
worksheet_summary.write('M20', ' ')
worksheet_summary.write('N20', ' ')
worksheet_summary.write('O20', ' ')

worksheet_summary.write('A13', 'K72 MUNI',format_general)
worksheet_summary.write('A14', 'K78 MUNTAX',format_general)
worksheet_summary.write('A15', 'K79 MUNCC',format_general)
worksheet_summary.write('A16', 'K80 MUNBT',format_general)
worksheet_summary.write('A17', 'K81 MUNCCTAX',format_general)
worksheet_summary.write('A18', 'K82 MUNBTTAX',format_general)
worksheet_summary.write('A21', 'N88 CORPIG',format_general)
worksheet_summary.write('A25', 'N90 CORPNOTE',format_general)
worksheet_summary.write('A29', 'P01 CORPFRN',format_general)
worksheet_summary.write('A33', 'P02 CORPSP',format_general)
worksheet_summary.write('A37', 'K74 CORPHY',format_general)
worksheet_summary.write('A41', 'L81 Corp Other',format_general)
worksheet_summary.write('A45', 'P03 CORPDIST',format_general)
worksheet_summary.write('A49', 'N87 CORPXOVR',format_general)
worksheet_summary.write('A56', 'K76 CMO',format_general)
worksheet_summary.write('A57', 'M64 IO',format_general)

worksheet_summary.write('I13', 'K72 MUNI',format_general)
worksheet_summary.write('I14', 'K78 MUNTAX',format_general)
worksheet_summary.write('I15', 'K79 MUNCC',format_general)
worksheet_summary.write('I16', 'K80 MUNBT',format_general)
worksheet_summary.write('I17', 'K81 MUNCCTAX',format_general)
worksheet_summary.write('I18', 'K82 MUNBTTAX',format_general)
worksheet_summary.write('I21', 'N88 CORPIG',format_general)
worksheet_summary.write('I25', 'N90 CORPNOTE',format_general)
worksheet_summary.write('I29', 'P01 CORPFRN',format_general)
worksheet_summary.write('I33', 'P02 CORPSP',format_general)
worksheet_summary.write('I37', 'K74 CORPHY',format_general)
worksheet_summary.write('I41', 'L81 Corp Other',format_general)
worksheet_summary.write('I45', 'P03 CORPDIST',format_general)
worksheet_summary.write('I49', 'N87 CORPXOVR',format_general)
worksheet_summary.write('I56', 'K76 CMO',format_general)
worksheet_summary.write('I57', 'M64 IO',format_general)

worksheet_summary.write('A22', ' ')
worksheet_summary.write('I22', ' ')
worksheet_summary.write('B23', 'Total',format_mini_total)
worksheet_summary.write('J23', 'Total',format_mini_total)
worksheet_summary.write('C23', Corp_N88_Cost_Summary_Recent,format_mini_total)
worksheet_summary.write('D23', Corp_N88_Market_Value_Summary_Recent,format_mini_total)
worksheet_summary.write('E23', Corp_N88_Requirement_Summary_Recent,format_mini_total)
worksheet_summary.write('F23', Corp_N88_Unreal_PNL_Summary_Recent,format_mini_total)
worksheet_summary.write('G23', Corp_N88_Real_PNL_Summary_Recent,format_mini_total)
worksheet_summary.write('I23', ' ')
worksheet_summary.write('K23', Corp_N88_Cost_Daily_Change,format_mini_total)
worksheet_summary.write('L23', Corp_N88_Market_Value_Daily_Change,format_mini_total)
worksheet_summary.write('M23', Corp_N88_Requirement_Daily_Change ,format_mini_total)
worksheet_summary.write('N23', Corp_N88_Unreal_PNL_Daily_Change,format_mini_total)
worksheet_summary.write('O23', Corp_N88_Real_PNL_Daily_Change,format_mini_total)


worksheet_summary.write('A24', ' ')
worksheet_summary.write('B24', ' ')
worksheet_summary.write('C24', ' ')
worksheet_summary.write('D24', ' ')
worksheet_summary.write('E24', ' ')
worksheet_summary.write('F24', ' ')
worksheet_summary.write('G24', ' ')
worksheet_summary.write('I24', ' ')
worksheet_summary.write('J24', ' ')
worksheet_summary.write('K24', ' ')
worksheet_summary.write('L24', ' ')
worksheet_summary.write('M24', ' ')
worksheet_summary.write('N24', ' ')
worksheet_summary.write('O24', ' ')

worksheet_summary.write('I26', ' ')
worksheet_summary.write('A26', ' ')
worksheet_summary.write('B27', 'Total',format_mini_total)
worksheet_summary.write('J27', 'Total',format_mini_total)
worksheet_summary.write('C27', Corp_N90_Cost_Summary_Recent,format_mini_total)
worksheet_summary.write('D27', Corp_N90_Market_Value_Summary_Recent,format_mini_total)
worksheet_summary.write('E27', Corp_N90_Requirement_Summary_Recent,format_mini_total)
worksheet_summary.write('F27', Corp_N90_Unreal_PNL_Summary_Recent,format_mini_total)
worksheet_summary.write('G27', Corp_N90_Real_PNL_Summary_Recent,format_mini_total)
worksheet_summary.write('I27', ' ')
worksheet_summary.write('K27', Corp_N90_Cost_Daily_Change,format_mini_total)
worksheet_summary.write('L27', Corp_N90_Market_Value_Daily_Change,format_mini_total)
worksheet_summary.write('M27', Corp_N90_Requirement_Daily_Change,format_mini_total)
worksheet_summary.write('N27', Corp_N90_Unreal_PNL_Daily_Change,format_mini_total)
worksheet_summary.write('O27', Corp_N90_Real_PNL_Daily_Change,format_mini_total)

worksheet_summary.write('A28', ' ')
worksheet_summary.write('B28', ' ')
worksheet_summary.write('C28', ' ')
worksheet_summary.write('D28', ' ')
worksheet_summary.write('E28', ' ')
worksheet_summary.write('F28', ' ')
worksheet_summary.write('G28', ' ')
worksheet_summary.write('I28', ' ')
worksheet_summary.write('J28', ' ')
worksheet_summary.write('K28', ' ')
worksheet_summary.write('L28', ' ')
worksheet_summary.write('M28', ' ')
worksheet_summary.write('N28', ' ')
worksheet_summary.write('O28', ' ')

worksheet_summary.write('A30', ' ')
worksheet_summary.write('I30', ' ')
worksheet_summary.write('B31', 'Total',format_mini_total)
worksheet_summary.write('J31', 'Total',format_mini_total)
worksheet_summary.write('C31', Corp_P01_Cost_Summary_Recent ,format_mini_total)
worksheet_summary.write('D31', Corp_P01_Market_Value_Summary_Recent,format_mini_total)
worksheet_summary.write('E31', Corp_P01_Requirement_Summary_Recent,format_mini_total)
worksheet_summary.write('F31', Corp_P01_Unreal_PNL_Summary_Recent,format_mini_total)
worksheet_summary.write('G31', Corp_P01_Real_PNL_Summary_Recent,format_mini_total)
worksheet_summary.write('I31', ' ')
worksheet_summary.write('K31', Corp_P01_Cost_Daily_Change,format_mini_total)
worksheet_summary.write('L31', Corp_P01_Market_Value_Daily_Change,format_mini_total)
worksheet_summary.write('M31', Corp_P01_Requirement_Daily_Change,format_mini_total)
worksheet_summary.write('N31', Corp_P01_Unreal_PNL_Daily_Change,format_mini_total)
worksheet_summary.write('O31', Corp_P01_Real_PNL_Daily_Change,format_mini_total)

worksheet_summary.write('A32', ' ')
worksheet_summary.write('B32', ' ')
worksheet_summary.write('C32', ' ')
worksheet_summary.write('D32', ' ')
worksheet_summary.write('E32', ' ')
worksheet_summary.write('F32', ' ')
worksheet_summary.write('G32', ' ')
worksheet_summary.write('I32', ' ')
worksheet_summary.write('J32', ' ')
worksheet_summary.write('K32', ' ')
worksheet_summary.write('L32', ' ')
worksheet_summary.write('M32', ' ')
worksheet_summary.write('N32', ' ')
worksheet_summary.write('O32', ' ')


worksheet_summary.write('A34', ' ')
worksheet_summary.write('I34', ' ')
worksheet_summary.write('B35', 'Total',format_mini_total)
worksheet_summary.write('J35', 'Total',format_mini_total)
worksheet_summary.write('C35', Corp_P02_Cost_Summary_Recent ,format_mini_total)
worksheet_summary.write('D35', Corp_P02_Market_Value_Summary_Recent,format_mini_total)
worksheet_summary.write('E35', Corp_P02_Requirement_Summary_Recent,format_mini_total)
worksheet_summary.write('F35', Corp_P02_Unreal_PNL_Summary_Recent,format_mini_total)
worksheet_summary.write('G35', Corp_P02_Real_PNL_Summary_Recent,format_mini_total)
worksheet_summary.write('I35', ' ')
worksheet_summary.write('K35', Corp_P02_Cost_Daily_Change,format_mini_total)
worksheet_summary.write('L35', Corp_P02_Market_Value_Daily_Change,format_mini_total)
worksheet_summary.write('M35', Corp_P02_Requirement_Daily_Change,format_mini_total)
worksheet_summary.write('N35', Corp_P02_Unreal_PNL_Daily_Change,format_mini_total)
worksheet_summary.write('O35', Corp_P02_Real_PNL_Daily_Change,format_mini_total)


Corp_P02_Cost_Summary_Recent = P02_Summary_Recent['Cost'].sum()
Corp_P02_Market_Value_Summary_Recent = P02_Summary_Recent['Market Value'].sum()
Corp_P02_Requirement_Summary_Recent = P02_Summary_Recent['Requirement'].sum()
Corp_P02_Unreal_PNL_Summary_Recent = P02_Summary_Recent['Unreal PNL'].sum()
Corp_P02_Real_PNL_Summary_Recent = P02_Summary_Recent['Real PNL'].sum()

Corp_P02_Cost_Daily_Change = P02_Daily_Change['Cost'].sum()
Corp_P02_Market_Value_Daily_Change = P02_Daily_Change['Market Value'].sum()
Corp_P02_Requirement_Daily_Change = P02_Daily_Change['Requirement'].sum()
Corp_P02_Unreal_PNL_Daily_Change = P02_Daily_Change['Unreal PNL'].sum()
Corp_P02_Real_PNL_Daily_Change = P02_Daily_Change['Real PNL'].sum()


worksheet_summary.write('A36', ' ')
worksheet_summary.write('B36', ' ')
worksheet_summary.write('C36', ' ')
worksheet_summary.write('D36', ' ')
worksheet_summary.write('E36', ' ')
worksheet_summary.write('F36', ' ')
worksheet_summary.write('G36', ' ')
worksheet_summary.write('I36', ' ')
worksheet_summary.write('J36', ' ')
worksheet_summary.write('K36', ' ')
worksheet_summary.write('L36', ' ')
worksheet_summary.write('M36', ' ')
worksheet_summary.write('N36', ' ')
worksheet_summary.write('O36', ' ')

worksheet_summary.write('A38', ' ')
worksheet_summary.write('I38', ' ')
worksheet_summary.write('B39', 'Total',format_mini_total)
worksheet_summary.write('J39', 'Total',format_mini_total)
worksheet_summary.write('C39', Corp_K74_Cost_Summary_Recent ,format_mini_total)
worksheet_summary.write('D39', Corp_K74_Market_Value_Summary_Recent,format_mini_total)
worksheet_summary.write('E39', Corp_K74_Requirement_Summary_Recent,format_mini_total)
worksheet_summary.write('F39', Corp_K74_Unreal_PNL_Summary_Recent,format_mini_total)
worksheet_summary.write('G39', Corp_K74_Real_PNL_Summary_Recent,format_mini_total)
worksheet_summary.write('I39', ' ')
worksheet_summary.write('K39', Corp_K74_Cost_Daily_Change,format_mini_total)
worksheet_summary.write('L39', Corp_K74_Market_Value_Daily_Change,format_mini_total)
worksheet_summary.write('M39', Corp_K74_Requirement_Daily_Change,format_mini_total)
worksheet_summary.write('N39', Corp_K74_Unreal_PNL_Daily_Change,format_mini_total)
worksheet_summary.write('O39', Corp_K74_Real_PNL_Daily_Change,format_mini_total)

worksheet_summary.write('A42', ' ')
worksheet_summary.write('I42', ' ')
worksheet_summary.write('B43', 'Total',format_mini_total)
worksheet_summary.write('J43', 'Total',format_mini_total)
worksheet_summary.write('C43', Corp_L81_Cost_Summary_Recent ,format_mini_total)
worksheet_summary.write('D43', Corp_L81_Market_Value_Summary_Recent,format_mini_total)
worksheet_summary.write('E43', Corp_L81_Requirement_Summary_Recent,format_mini_total)
worksheet_summary.write('F43', Corp_L81_Unreal_PNL_Summary_Recent,format_mini_total)
worksheet_summary.write('G43', Corp_L81_Real_PNL_Summary_Recent,format_mini_total)
worksheet_summary.write('I43', ' ')
worksheet_summary.write('K43', Corp_L81_Cost_Daily_Change,format_mini_total)
worksheet_summary.write('L43', Corp_L81_Market_Value_Daily_Change,format_mini_total)
worksheet_summary.write('M43', Corp_L81_Requirement_Daily_Change,format_mini_total)
worksheet_summary.write('N43', Corp_L81_Unreal_PNL_Daily_Change,format_mini_total)
worksheet_summary.write('O43', Corp_L81_Real_PNL_Daily_Change,format_mini_total)

worksheet_summary.write('A44', ' ')
worksheet_summary.write('B44', ' ')
worksheet_summary.write('C44', ' ')
worksheet_summary.write('D44', ' ')
worksheet_summary.write('E44', ' ')
worksheet_summary.write('F44', ' ')
worksheet_summary.write('G44', ' ')
worksheet_summary.write('I44', ' ')
worksheet_summary.write('J44', ' ')
worksheet_summary.write('K44', ' ')
worksheet_summary.write('L44', ' ')
worksheet_summary.write('M44', ' ')
worksheet_summary.write('N44', ' ')
worksheet_summary.write('O44', ' ')


worksheet_summary.write('A48', ' ')
worksheet_summary.write('B48', ' ')
worksheet_summary.write('C48', ' ')
worksheet_summary.write('D48', ' ')
worksheet_summary.write('E48', ' ')
worksheet_summary.write('F48', ' ')
worksheet_summary.write('G48', ' ')
worksheet_summary.write('I48', ' ')
worksheet_summary.write('J48', ' ')
worksheet_summary.write('K48', ' ')
worksheet_summary.write('L48', ' ')
worksheet_summary.write('M48', ' ')
worksheet_summary.write('N48', ' ')
worksheet_summary.write('O48', ' ')

worksheet_summary.write('A50', ' ')
worksheet_summary.write('I50', ' ')
worksheet_summary.write('B51', 'Total',format_mini_total)
worksheet_summary.write('J51', 'Total',format_mini_total)
worksheet_summary.write('C51', Corp_N87_Cost_Summary_Recent ,format_mini_total)
worksheet_summary.write('D51', Corp_N87_Market_Value_Summary_Recent,format_mini_total)
worksheet_summary.write('E51', Corp_N87_Requirement_Summary_Recent,format_mini_total)
worksheet_summary.write('F51', Corp_N87_Unreal_PNL_Summary_Recent,format_mini_total)
worksheet_summary.write('G51', Corp_N87_Real_PNL_Summary_Recent,format_mini_total)
worksheet_summary.write('I51', ' ')
worksheet_summary.write('K51', Corp_N87_Cost_Daily_Change,format_mini_total)
worksheet_summary.write('L51', Corp_N87_Market_Value_Daily_Change,format_mini_total)
worksheet_summary.write('M51', Corp_N87_Requirement_Daily_Change,format_mini_total)
worksheet_summary.write('N51', Corp_N87_Unreal_PNL_Daily_Change,format_mini_total)
worksheet_summary.write('O51', Corp_N87_Real_PNL_Daily_Change,format_mini_total)


worksheet_summary.write('A40', ' ')
worksheet_summary.write('B40', ' ')
worksheet_summary.write('C40', ' ')
worksheet_summary.write('D40', ' ')
worksheet_summary.write('E40', ' ')
worksheet_summary.write('F40', ' ')
worksheet_summary.write('G40', ' ')
worksheet_summary.write('I40', ' ')
worksheet_summary.write('J40', ' ')
worksheet_summary.write('K40', ' ')
worksheet_summary.write('L40', ' ')
worksheet_summary.write('M40', ' ')
worksheet_summary.write('N40', ' ')
worksheet_summary.write('O40', ' ')

worksheet_summary.write('A46', ' ')
worksheet_summary.write('I46', ' ')
worksheet_summary.write('B47', 'Total',format_mini_total)
worksheet_summary.write('J47', 'Total',format_mini_total)
worksheet_summary.write('C47', Corp_P03_Cost_Summary_Recent ,format_mini_total)
worksheet_summary.write('D47', Corp_P03_Market_Value_Summary_Recent,format_mini_total)
worksheet_summary.write('E47', Corp_P03_Requirement_Summary_Recent,format_mini_total)
worksheet_summary.write('F47', Corp_P03_Unreal_PNL_Summary_Recent,format_mini_total)
worksheet_summary.write('G47', Corp_P03_Real_PNL_Summary_Recent,format_mini_total)
worksheet_summary.write('I47', ' ')
worksheet_summary.write('K47', Corp_P03_Cost_Daily_Change,format_mini_total)
worksheet_summary.write('L47', Corp_P03_Market_Value_Daily_Change,format_mini_total)
worksheet_summary.write('M47', Corp_P03_Requirement_Daily_Change,format_mini_total)
worksheet_summary.write('N47', Corp_P03_Unreal_PNL_Daily_Change,format_mini_total)
worksheet_summary.write('O47', Corp_P03_Real_PNL_Daily_Change,format_mini_total)


worksheet_summary.write('A52', 'Corp Total',format_subtotal)
worksheet_summary.write('B52', ' ',format_subtotal)
worksheet_summary.write('I52', 'Corp Total',format_subtotal)
worksheet_summary.write('J52', ' ',format_subtotal)
worksheet_summary.write('C52', Corp_Total_Cost_Summary,format_subtotal)
worksheet_summary.write('D52', Corp_Total_Market_Value_Summary,format_subtotal)
worksheet_summary.write('E52', Corp_Total_Requirement_Summary,format_subtotal)
worksheet_summary.write('F52', Corp_Total_Unreal_PNL_Summary,format_subtotal)
worksheet_summary.write('G52', Corp_Total_Real_PNL_Summary ,format_subtotal)
worksheet_summary.write('K52', Corp_Total_Cost_Daily,format_subtotal)
worksheet_summary.write('L52', Corp_Total_Market_Value_Daily ,format_subtotal)
worksheet_summary.write('M52', Corp_Total_Requirement_Daily,format_subtotal)
worksheet_summary.write('N52', Corp_Total_Unreal_PNL_Daily,format_subtotal)
worksheet_summary.write('O52', Corp_Total_Real_PNL_Daily,format_subtotal)

worksheet_summary.write('A53', ' ')
worksheet_summary.write('B53', ' ')
worksheet_summary.write('C53', ' ')
worksheet_summary.write('D53', ' ')
worksheet_summary.write('E53', ' ')
worksheet_summary.write('F53', ' ')
worksheet_summary.write('G53', ' ')
worksheet_summary.write('I53', ' ')
worksheet_summary.write('J53', ' ')
worksheet_summary.write('K53', ' ')
worksheet_summary.write('L53', ' ')
worksheet_summary.write('M53', ' ')
worksheet_summary.write('N53', ' ')
worksheet_summary.write('O53', ' ')


worksheet_summary.write('A54', 'CD Total',format_subtotal)
worksheet_summary.write('B54', ' ',format_subtotal)
worksheet_summary.write('I54', 'CD Total',format_subtotal)
worksheet_summary.write('J54', ' ',format_subtotal)
worksheet_summary.write('C54', CD_Cost_Summary_Recent ,format_subtotal)
worksheet_summary.write('D54', CD_Market_Value_Summary_Recent,format_subtotal)
worksheet_summary.write('E54', CD_Requirement_Summary_Recent,format_subtotal)
worksheet_summary.write('F54', CD_Unreal_PNL_Summary_Recent,format_subtotal)
worksheet_summary.write('G54', CD_Real_PNL_Summary_Recent,format_subtotal)
worksheet_summary.write('K54', CD_Cost_Daily_Change,format_subtotal)
worksheet_summary.write('L54', CD_Market_Value_Daily_Change,format_subtotal)
worksheet_summary.write('M54', CD_Requirement_Daily_Change ,format_subtotal)
worksheet_summary.write('N54', CD_Unreal_PNL_Daily_Change,format_subtotal)
worksheet_summary.write('O54', CD_Real_PNL_Daily_Change,format_subtotal)


worksheet_summary.write('A55', ' ')
worksheet_summary.write('B55', ' ')
worksheet_summary.write('C55', ' ')
worksheet_summary.write('D55', ' ')
worksheet_summary.write('E55', ' ')
worksheet_summary.write('F55', ' ')
worksheet_summary.write('G55', ' ')
worksheet_summary.write('I55', ' ')
worksheet_summary.write('J55', ' ')
worksheet_summary.write('K55', ' ')
worksheet_summary.write('L55', ' ')
worksheet_summary.write('M55', ' ')
worksheet_summary.write('N55', ' ')
worksheet_summary.write('O55', ' ')

worksheet_summary.write('B56', '     -')
worksheet_summary.write('B57', '     -')
worksheet_summary.write('J56', '     -')
worksheet_summary.write('J57', '     -')

worksheet_summary.write('A58', 'CMO Total',format_subtotal)
worksheet_summary.write('B58', ' ',format_subtotal)
worksheet_summary.write('I58', 'CMO Total',format_subtotal)
worksheet_summary.write('J58', ' ',format_subtotal)
worksheet_summary.write('C58', CMO_Cost_Summary_Recent,format_subtotal)
worksheet_summary.write('D58', CMO_Market_Value_Summary_Recent,format_subtotal)
worksheet_summary.write('E58', CMO_Requirement_Summary_Recent,format_subtotal)
worksheet_summary.write('F58', CMO_Unreal_PNL_Summary_Recent,format_subtotal)
worksheet_summary.write('G58', CMO_Real_PNL_Summary_Recent,format_subtotal)
worksheet_summary.write('K58', CMO_Cost_Daily_Change,format_subtotal)
worksheet_summary.write('L58', CMO_Market_Value_Daily_Change,format_subtotal)
worksheet_summary.write('M58', CMO_Requirement_Daily_Change,format_subtotal)
worksheet_summary.write('N58', CMO_Unreal_PNL_Daily_Change,format_subtotal)
worksheet_summary.write('O58', CMO_Real_PNL_Daily_Change,format_subtotal)


worksheet_summary.write('A60', 'Firm Total',format_subtotal)
worksheet_summary.write('B60', ' ',format_subtotal)
worksheet_summary.write('D60', Firm_Market_Value_Summary_Total,format_subtotal)
worksheet_summary.write('E60', Firm_Requirement_Summary_Total,format_subtotal)
worksheet_summary.write('F60', Firm_Unreal_PNL_Summary_Total,format_subtotal)
worksheet_summary.write('G60', Firm_Real_PNL_Summary_Total,format_subtotal)

worksheet_summary.write('I60', 'Firm Total',format_subtotal)
worksheet_summary.write('J60', ' ',format_subtotal)
worksheet_summary.write('L60', Firm_Market_Value_Daily_Total,format_subtotal)
worksheet_summary.write('M60', Firm_Requirement_Daily_Total,format_subtotal)
worksheet_summary.write('N60', Firm_Unreal_PNL_Daily_Total,format_subtotal)
worksheet_summary.write('O60', Firm_Real_PNL_Daily_Total,format_subtotal)


                                      
worksheet_summary.merge_range('A11:G11', 'Month Summary',merge_format)

worksheet_summary.merge_range('I11:O11', 'Daily Change',merge_format)
worksheet_summary.set_row(1,None,format_general_row) 
worksheet_summary.set_row(2,None,format_general_row) 
worksheet_summary.set_row(3,None,format_general_row) 
worksheet_summary.set_row(4,None,format_general_row) 
worksheet_summary.set_row(5,None,format_general_row) 
worksheet_summary.set_row(6,None,format_general_row) 
worksheet_summary.set_row(7,None,format_general_row) 
worksheet_summary.set_row(8,None,format_general_row) 

worksheet_summary.set_row(9,None,format_general_row)
worksheet_summary.set_row(10,None,format_general_row)
worksheet_summary.set_row(12,None,format_general_row)  
worksheet_summary.set_row(13,None,format_general_row)  
worksheet_summary.set_row(14,None,format_general_row)  
worksheet_summary.set_row(15,None,format_general_row) 
worksheet_summary.set_row(16,None,format_general_row) 
worksheet_summary.set_row(17,None,format_general_row) 
worksheet_summary.set_row(18,None,format_general_row) 
worksheet_summary.set_row(19,3,format_general_row) 
worksheet_summary.set_row(20,None,format_general_row) 
worksheet_summary.set_row(21,None,format_general_row) 
worksheet_summary.set_row(22,None,format_general_row) 
worksheet_summary.set_row(23,3,format_general_row) 
worksheet_summary.set_row(24,None,format_general_row) 
worksheet_summary.set_row(25,None,format_general_row) 
worksheet_summary.set_row(26,None,format_general_row) 
worksheet_summary.set_row(27,3,format_general_row) 
worksheet_summary.set_row(28,None,format_general_row) 
worksheet_summary.set_row(29,None,format_general_row) 
worksheet_summary.set_row(30,None,format_general_row) 
worksheet_summary.set_row(31,3,format_general_row) 
worksheet_summary.set_row(32,None,format_general_row) 
worksheet_summary.set_row(33,None,format_general_row) 
worksheet_summary.set_row(34,None,format_general_row) 
worksheet_summary.set_row(35,3,format_general_row) 
worksheet_summary.set_row(36,None,format_general_row) 
worksheet_summary.set_row(37,None,format_general_row) 
worksheet_summary.set_row(38,None,format_general_row) 
worksheet_summary.set_row(39,3,format_general_row) 
worksheet_summary.set_row(40,None,format_general_row) 
worksheet_summary.set_row(41,None,format_general_row) 
worksheet_summary.set_row(42,None,format_general_row) 
worksheet_summary.set_row(43,3,format_general_row) 
worksheet_summary.set_row(44,None,format_general_row) 
worksheet_summary.set_row(45,None,format_general_row) 
worksheet_summary.set_row(46,None,format_general_row) 
worksheet_summary.set_row(47,3,format_general_row) 
worksheet_summary.set_row(48,None,format_general_row) 
worksheet_summary.set_row(49,None,format_general_row) 
worksheet_summary.set_row(50,None,format_general_row) 
worksheet_summary.set_row(51,None,format_general_row) 
worksheet_summary.set_row(52,3) 
worksheet_summary.set_row(53,None,format_general_row) 
worksheet_summary.set_row(54,3) 
worksheet_summary.set_row(55,None,format_general_row)
worksheet_summary.set_row(56,None,format_general_row)
worksheet_summary.set_row(57,None,format_general_row)
worksheet_summary.set_row(58,3)
worksheet_summary.set_row(59,None,format_general_row)
# worksheet_summary.set_row(60,None,format_subtotal_row)

worksheet_summary.set_column('A:A',15,None)
worksheet_summary.set_column('B:B',13,None)
worksheet_summary.set_column('C:C',13,None,{'hidden':True})
worksheet_summary.set_column('D:D',15,None)
worksheet_summary.set_column('E:E',13,None)
worksheet_summary.set_column('F:F',13,None)
worksheet_summary.set_column('G:G',13,None)
worksheet_summary.set_column('H:H',2,None)
worksheet_summary.set_column('I:I',15,None)
worksheet_summary.set_column('J:J',13,None)
worksheet_summary.set_column('K:K',13,None,{'hidden':True})
worksheet_summary.set_column('L:L',13,None)
worksheet_summary.set_column('M:M',13,None)
worksheet_summary.set_column('N:N',13,None)
worksheet_summary.set_column('O:O',13,None)

worksheet_summary.write('A1', 'Page Links',format_top_summary)

worksheet_summary.write_url('A2',"internal:'Quantity Diff'!A1",format_url_links,string = '1. Quantity Diff')
worksheet_summary.write_url('A3',"internal:'PNL Diff'!A1",format_url_links,string = '2. PNL Diff')
worksheet_summary.write_url('A4',"internal:'Adj Unrealized PNL Change'!A1",format_url_links,string = '3. Adj Unrealized PNL Change')
worksheet_summary.write_url('A5',"internal:'Requirement Change'!A1",format_url_links,string = '4. Requirement Change')
worksheet_summary.write_url('A6',"internal:'HT Detail'!A1",format_url_links,string = '5. HT Detail')
worksheet_summary.write_url('A7',"internal:'TW Detail'!A1",format_url_links ,string = '6. TW Detail')

# Conditional Formating
worksheet_summary.conditional_format('M13:M60', {'type':'cell',
                                        'criteria': '<',
                                        'value':    0,
                                        'format':   format_general_row_green})
worksheet_summary.conditional_format('M13:M60', {'type':'cell',
                                        'criteria': '>',
                                        'value':    0,
                                        'format':   format_general_row_red})
worksheet_summary.conditional_format('N13:O60', {'type':'cell',
                                        'criteria': '<',
                                        'value':    0,
                                        'format':   format_general_row_red})
worksheet_summary.conditional_format('N13:O60', {'type':'cell',
                                        'criteria': '>',
                                        'value':    0,
                                        'format':   format_general_row_green})
"""
Create and format Detail sheet
"""



format1 = workbook.add_format({'num_format': '#,##0',
                               'font_size':'9'})
format2 = workbook.add_format({'num_format': '#,##0.00',
                               'font_size':'9'})

worksheet_summary.insert_image('M1', 'P:/1. Individual Folders/Chad/Python Scripts/PNL Report/Logo.png',{'x_scale':.7,'y_scale':.7})

# Format each colum to fit and display data correclty


Hilltop_x = Hilltop_Individual_Summary
Hilltop_y = Hilltop_x


def Summary_Individual_Sheets(Hilltop_Individual_Summary):
    column_summary = ('TW - HT Quantity Discrepancy',
                      'HT Change in Quantity',
                      'Adj Unreal PNL Change',
                      'HT-TW PNL Discrepancy',
                      'Requirement Change')
    for item in column_summary:
        Hilltop_Individual_Summary[item] = Hilltop_Individual_Summary[item]
    Hilltop_QTY_DSP = Hilltop_Individual_Summary[(Hilltop_Individual_Summary['TW - HT Quantity Discrepancy'] != 0)]
    Hilltop_HT_QTY_Change = Hilltop_Individual_Summary[(Hilltop_Individual_Summary['HT Change in Quantity'] != 0)]
    Hilltop_Adj_Unreal_PNL_Change = Hilltop_Individual_Summary[(Hilltop_Individual_Summary['Adj Unreal PNL Change'] != 0)]
    Hilltop_HT_TW_PNL_DSP = Hilltop_Individual_Summary[(Hilltop_Individual_Summary['HT-TW PNL Discrepancy'] != 0)]
    Hilltop_Requirement_Change = Hilltop_Individual_Summary[(Hilltop_Individual_Summary['Requirement Change'] != 0)]
    return Hilltop_QTY_DSP,Hilltop_HT_QTY_Change,Hilltop_Adj_Unreal_PNL_Change, Hilltop_HT_TW_PNL_DSP,Hilltop_Requirement_Change

Individual_Sheets = Summary_Individual_Sheets(Hilltop_Individual_Summary)


Hilltop_QTY_DSP  = Individual_Sheets[0]
Hilltop_QTY_DSP = pd.merge(Hilltop_QTY_DSP, Hilltop_x, on='Cusip', how='left')
Hilltop_QTY_DSP['TW - HT Quantity Discrepancy_x'] = Hilltop_QTY_DSP['TW - HT Quantity Discrepancy_x'].abs()
Hilltop_QTY_DSP.sort_values('TW - HT Quantity Discrepancy_x', axis=0, ascending=False, inplace=True)
Hilltop_QTY_DSP = Hilltop_QTY_DSP[['Security_x','Cusip','Account_x','TW - HT Quantity Discrepancy_y']]
Hilltop_QTY_DSP.rename(columns={'Security_x': 'Security','Account_x':'Account','TW - HT Quantity Discrepancy_y':"QTY DSP"}, inplace=True)
QTY_DSP_Cleared_Positions_Drop = QTY_DSP_Cleared_Positions.drop('Position Notes',axis = 1)
Hilltop_QTY_DSP = Hilltop_QTY_DSP.append(QTY_DSP_Cleared_Positions_Drop)
Hilltop_QTY_DSP.drop_duplicates(subset ='Cusip',keep = False, inplace = True)
Hilltop_QTY_DSP = Hilltop_QTY_DSP[['Security','Account','Cusip','QTY DSP']]
QTY_DSP_Cleared_Positions= QTY_DSP_Cleared_Positions[['Security','Account','Cusip','QTY DSP','Position Notes']]


Hilltop_HT_TW_PNL_DSP = Individual_Sheets[3]
Hilltop_HT_TW_PNL_DSP.sort_values('HT-TW PNL Discrepancy', axis=0, ascending=False, inplace=True)
Hilltop_HT_TW_PNL_DSP = pd.merge(Hilltop_HT_TW_PNL_DSP, Hilltop_x, on='Cusip', how='left')
Hilltop_HT_TW_PNL_DSP['HT-TW PNL Discrepancy_x'] = Hilltop_HT_TW_PNL_DSP['HT-TW PNL Discrepancy_x'].abs()
Hilltop_HT_TW_PNL_DSP.sort_values('HT-TW PNL Discrepancy_x', axis=0, ascending=False, inplace=True)
Hilltop_HT_TW_PNL_DSP = Hilltop_HT_TW_PNL_DSP[['Security_x','Cusip','Account_x','HT-TW PNL Discrepancy_y']]
Hilltop_HT_TW_PNL_DSP_Lower = Hilltop_HT_TW_PNL_DSP[(Hilltop_HT_TW_PNL_DSP['HT-TW PNL Discrepancy_y'] < -10)]
Hilltop_HT_TW_PNL_DSP_Upper = Hilltop_HT_TW_PNL_DSP[(Hilltop_HT_TW_PNL_DSP['HT-TW PNL Discrepancy_y'] > 10)]
Hilltop_Chunks = [Hilltop_HT_TW_PNL_DSP_Upper,Hilltop_HT_TW_PNL_DSP_Lower]
Hilltop_HT_TW_PNL_DSP = pd.concat(Hilltop_Chunks)
Hilltop_HT_TW_PNL_DSP.sort_values(by = 'HT-TW PNL Discrepancy_y',ascending = False)
x = subject_name(file_text)
y = x[3]
PNL_DSP_Date = y[:10]
Hilltop_HT_TW_PNL_DSP['Date'] = PNL_DSP_Date
Hilltop_HT_TW_PNL_DSP = Hilltop_HT_TW_PNL_DSP[['Date','Security_x','Account_x','Cusip','HT-TW PNL Discrepancy_y',   ]]


"""
Pull and Generate new PNL DIFF Position File
"""
# PNL_Report_Date = x[2]
# PNL_Report_File_Most_Recent = 'P:/2. Corps/PNL_Daily_Report/Reports/PNL_Report.xlsx'
# PNL_DSP_Yesterday = pd.read_excel(PNL_Report_File_Most_Recent,sheet_name = 'PNL Diff')  #DSP Items from previous report
# Additions_to_Running_PNL_DSP = PNL_DSP_Yesterday[['Date','Security','Account','Cusip','PNL DSP']]  #DSP Items from previous report sorted
# Additions_to_Running_PNL_DSP.dropna(inplace = True)
# Current_Running_PNL_DSP_Filepath = 'P:/2. Corps/PNL_Daily_Report/Cleared_Position_File/PNL_DSP_History.xlsx' 
# Current_Running_PNL_DSP = pd.read_excel(Current_Running_PNL_DSP_Filepath)# reads running PNL DSP file
# Current_Running_PNL_DSP = Current_Running_PNL_DSP[Current_Running_PNL_DSP['Previous PNL DSP'] > 5] # filters out 'closed' Positions
# Additions_to_Running_PNL_DSP['Previous PNL DSP'] = Additions_to_Running_PNL_DSP['PNL DSP']
# Additions_to_Running_PNL_DSP = Additions_to_Running_PNL_DSP[['Account','Cusip','Date','Previous PNL DSP','Security']]
# Current_Running_PNL_DSP_List = [Current_Running_PNL_DSP,Additions_to_Running_PNL_DSP]                                  # creates a list to concat the dataframes together
# Complete_Running_PNL_DSP = pd.concat(Current_Running_PNL_DSP_List)     # Concatinate the running PNL DSP and the Most recent DSP
# Complete_Running_PNL_DSP.drop_duplicates(subset = 'Cusip', keep = 'first', inplace = True)
# Complete_Running_PNL_DSP_with_Detail = pd.merge(Complete_Running_PNL_DSP,Hilltop_Recent, on = 'Cusip', how = 'left')
# Complete_Running_PNL_DSP_with_Detail['Net PNL DSP'] = Complete_Running_PNL_DSP_with_Detail['Previous PNL DSP'] + Complete_Running_PNL_DSP_with_Detail['HT-TW PNL Discrepancy']
# Complete_Running_PNL_DSP_with_Detail = Complete_Running_PNL_DSP_with_Detail[['Date','Security_x','Account_x','Cusip','Previous PNL DSP','HT-TW PNL Discrepancy','Net PNL DSP']]

# Complete_Running_PNL_DSP_with_Detail.rename(columns={'Date':'Date',
#                                         'Security_x':'Security',
#                                         'Account_x':'Account',
#                                         'Cusip':'Cusip',
#                                         'Previous PNL DSP':'Previous PNL DSP',
#                                          'HT-TW PNL Discrepancy':'Current PNL DSP'
#                                         },inplace = True)

Hilltop_Adj_Unreal_PNL_Change = Individual_Sheets[2]
Hilltop_Adj_Unreal_PNL_Change = pd.merge(Hilltop_Adj_Unreal_PNL_Change, Hilltop_x, on='Cusip', how='left')
Hilltop_Adj_Unreal_PNL_Change['Adj Unreal PNL Change_x'] = Hilltop_Adj_Unreal_PNL_Change['Adj Unreal PNL Change_x'].abs()
Hilltop_Adj_Unreal_PNL_Change.sort_values('Adj Unreal PNL Change_x', axis=0, ascending=False, inplace=True)
Hilltop_Adj_Unreal_PNL_Change = Hilltop_Adj_Unreal_PNL_Change[['Security_x','Cusip','Account_x','Adj Unreal PNL Change_y']]



Hilltop_Requirement_Change = Individual_Sheets[4]
Hilltop_Requirement_Change = pd.merge(Hilltop_Requirement_Change, Hilltop_x, on='Cusip', how='left')
Hilltop_Requirement_Change = Hilltop_Requirement_Change[['Cusip','Security_x','Account_x','Requirement Change_x','Requirement Change_y']]
Hilltop_Requirement_Change['Requirement Change_x'] = Hilltop_Requirement_Change['Requirement Change_x'].abs()
Hilltop_Requirement_Change.sort_values('Requirement Change_x', axis=0, ascending=False, inplace=True)
Hilltop_Requirement_Change = Hilltop_Requirement_Change[['Security_x','Cusip','Account_x','Requirement Change_y']]


"""
Write to Excel

"""

QTY_DSP_Cleared_Positions.to_excel(writer,sheet_name ='Quantity Diff',index=False,startrow=1,startcol=6)
Hilltop_QTY_DSP.to_excel(writer, sheet_name = 'Quantity Diff', index=False)
worksheet_Hilltop_QTY_DSP = writer.sheets['Quantity Diff']

Hilltop_HT_TW_PNL_DSP.to_excel(writer, sheet_name = 'PNL Diff', index=False)
Complete_Running_PNL_DSP_with_Detail.to_excel(writer, sheet_name = 'PNL Diff', index=False,startrow = 1,startcol=7)
worksheet_Hilltop_HT_TW_PNL_DSP = writer.sheets['PNL Diff']

Hilltop_Adj_Unreal_PNL_Change.to_excel(writer, sheet_name = 'Adj Unrealized PNL Change', index=False)
worksheet_Hilltop_Adj_Unreal_PNL_Change = writer.sheets['Adj Unrealized PNL Change']

Hilltop_Requirement_Change.to_excel(writer, sheet_name = 'Requirement Change', index=False)
worksheet_Hilltop_Requirement_Change = writer.sheets['Requirement Change']


Hilltop_Recent.to_excel(writer, sheet_name='HT Detail', index=False)
worksheet = writer.sheets['HT Detail']

worksheet_Hilltop_QTY_DSP.set_column('A:A', 35,format1)#, format7) #2
worksheet_Hilltop_QTY_DSP.set_column('B:B', 15,format1)#, format2) #2
worksheet_Hilltop_QTY_DSP.set_column('C:C', 12,format1)#, format7) #3
worksheet_Hilltop_QTY_DSP.set_column('D:D', 12,format1)#, format1)#4
worksheet_Hilltop_QTY_DSP.set_column('D:D', 12,format1)#, format1)#4
worksheet_Hilltop_QTY_DSP.set_column('E:E', 30,format1)#, format1)

worksheet_Hilltop_QTY_DSP.write('A1', 'Security',format_top_summary)#,format5)
worksheet_Hilltop_QTY_DSP.write('B1', 'Account',format_top_summary)#,format5)
worksheet_Hilltop_QTY_DSP.write('C1', 'Cusip',format_top_summary)#,format5)
worksheet_Hilltop_QTY_DSP.write('D1', 'QTY DSP',format_top_summary)#,format5)
worksheet_Hilltop_QTY_DSP.write('E1', 'Position Notes',format_top_summary)#,format5)

worksheet_Hilltop_QTY_DSP.write('G2', 'Security',format_top_summary)#,format5)
worksheet_Hilltop_QTY_DSP.write('H2', 'Account',format_top_summary)#,format5)
worksheet_Hilltop_QTY_DSP.write('I2', 'Cusip',format_top_summary)#,format5)
worksheet_Hilltop_QTY_DSP.write('J2', 'QTY DSP',format_top_summary)#,format5)
worksheet_Hilltop_QTY_DSP.write('K2', 'Position Notes',format_top_summary)#,format5)

worksheet_Hilltop_QTY_DSP.merge_range('G1:K1', 'Cleared QTY DSP',merge_format)

worksheet_Hilltop_QTY_DSP.set_column('G:G', 20,format1)#, format1) #2
worksheet_Hilltop_QTY_DSP.set_column('H:H', 12,format1)#, format1) #2
worksheet_Hilltop_QTY_DSP.set_column('I:I', 12,format1)#, format1) #3
worksheet_Hilltop_QTY_DSP.set_column('J:J', 15,format1)#, format1)#4
worksheet_Hilltop_QTY_DSP.set_column('K:K', 30,format1)#, format1)#4

# worksheet_Hilltop_QTY_DSP.freeze_panes(1, 1)
worksheet_Hilltop_QTY_DSP.autofilter('A1:V20000')
# worksheet_Hilltop_QTY_DSP.hide_gridlines(2)
# worksheet_Hilltop_QTY_DSP.protect('welcome123')
worksheet_Hilltop_QTY_DSP.set_zoom(90)
"""
"""
worksheet_Hilltop_Adj_Unreal_PNL_Change.set_column('A:A', 35,format1)#, format7) #2
worksheet_Hilltop_Adj_Unreal_PNL_Change.set_column('B:B', 12,format1)#, format2) #2
worksheet_Hilltop_Adj_Unreal_PNL_Change.set_column('C:C', 15,format1)#, format7) #3
worksheet_Hilltop_Adj_Unreal_PNL_Change.set_column('D:D', 12,format1)#, format1)#4

worksheet_Hilltop_Adj_Unreal_PNL_Change.write('A1', 'Security',format_top_summary)#,format5)
worksheet_Hilltop_Adj_Unreal_PNL_Change.write('B1', 'Cusip',format_top_summary)#,format5)
worksheet_Hilltop_Adj_Unreal_PNL_Change.write('C1', 'Account',format_top_summary)#,format5)
worksheet_Hilltop_Adj_Unreal_PNL_Change.write('D1', 'Adj Unreal PNL Change',format_top_summary)#,format5)
worksheet_Hilltop_Adj_Unreal_PNL_Change.set_zoom(90)
"""
"""
worksheet_Hilltop_HT_TW_PNL_DSP.set_column('A:A', 11,format1)#, format7) #2
worksheet_Hilltop_HT_TW_PNL_DSP.set_column('B:B', 25,format1)#, format2) #2
worksheet_Hilltop_HT_TW_PNL_DSP.set_column('C:C', 13,format1)#, format7) #3
worksheet_Hilltop_HT_TW_PNL_DSP.set_column('D:D', 11,format1)#, format1)#4
worksheet_Hilltop_HT_TW_PNL_DSP.set_column('E:E', 11,format1)#, format1)#4
worksheet_Hilltop_HT_TW_PNL_DSP.set_column('F:F', 15,format1)#, format1)#4

worksheet_Hilltop_HT_TW_PNL_DSP.write('A1', 'Date',format_top_summary)#,format5)
worksheet_Hilltop_HT_TW_PNL_DSP.write('B1', 'Security',format_top_summary)#,format5)
worksheet_Hilltop_HT_TW_PNL_DSP.write('C1', 'Account',format_top_summary)#,format5)
worksheet_Hilltop_HT_TW_PNL_DSP.write('D1', 'Cusip',format_top_summary)#,format5)
worksheet_Hilltop_HT_TW_PNL_DSP.write('E1', 'PNL DSP',format_top_summary)#,format5)
# worksheet_Hilltop_HT_TW_PNL_DSP.write('F1', 'Position Notes',format_top_summary)#,format5)


worksheet_Hilltop_HT_TW_PNL_DSP.write('H2', 'Date',format_top_summary)#,format5)
worksheet_Hilltop_HT_TW_PNL_DSP.write('I2', 'Security',format_top_summary)#,format5)
worksheet_Hilltop_HT_TW_PNL_DSP.write('J2', 'Account',format_top_summary)#,format5)
worksheet_Hilltop_HT_TW_PNL_DSP.write('K2', 'Cusip',format_top_summary)#,format5)
worksheet_Hilltop_HT_TW_PNL_DSP.write('L2', 'Previous PNL',format_top_summary)#,format5)
worksheet_Hilltop_HT_TW_PNL_DSP.write('M2', 'Current PNL',format_top_summary)#,format5)
worksheet_Hilltop_HT_TW_PNL_DSP.write('N2', 'Net PNL',format_top_summary)#,format5)

worksheet_Hilltop_HT_TW_PNL_DSP.merge_range('H1:N1', 'Unresolved PNL DSP',merge_format)

worksheet_Hilltop_HT_TW_PNL_DSP.set_column('G:G', 3,format1)#, format1) #2
worksheet_Hilltop_HT_TW_PNL_DSP.set_column('H:H', 12,format1)#, format1) #2
worksheet_Hilltop_HT_TW_PNL_DSP.set_column('I:I', 12,format1)#, format1) #3
worksheet_Hilltop_HT_TW_PNL_DSP.set_column('J:J', 15,format1)#, format1)#4
worksheet_Hilltop_HT_TW_PNL_DSP.set_column('K:K', 15,format1)#, format1)#4
worksheet_Hilltop_HT_TW_PNL_DSP.set_column('L:L', 15,format1)#, format1)#4
worksheet_Hilltop_HT_TW_PNL_DSP.set_column('M:M', 15,format1)#, format1)#4
worksheet_Hilltop_HT_TW_PNL_DSP.set_column('N:N', 15,format1)#, format1)#4
worksheet_Hilltop_HT_TW_PNL_DSP.set_zoom(90)

"""
"""
worksheet_Hilltop_Requirement_Change.set_column('A:A', 35,format1)#, format7) #2
worksheet_Hilltop_Requirement_Change.set_column('B:B', 15,format1)#, format2) #2
worksheet_Hilltop_Requirement_Change.set_column('C:C', 15,format1)#, format7) #3
worksheet_Hilltop_Requirement_Change.set_column('D:D', 15,format1)#, format1)#4

worksheet_Hilltop_Requirement_Change.write('A1', 'Security',format_top_summary)#,format5)
worksheet_Hilltop_Requirement_Change.write('B1', 'Cusip',format_top_summary)#,format5)
worksheet_Hilltop_Requirement_Change.write('C1', 'Account',format_top_summary)#,format5)
worksheet_Hilltop_Requirement_Change.write('D1', 'Requirement Change',format_top_summary)#,format5)
worksheet_Hilltop_Requirement_Change.set_zoom(90)


#
# worksheet_Hilltop_QTY_DSP.freeze_panes(1, 1)
worksheet_Hilltop_QTY_DSP.autofilter('A1:E20000')
# worksheet_Hilltop_QTY_DSP.hide_gridlines(2)

# worksheet_Hilltop_HT_TW_PNL_DSP.freeze_panes(1, 1)
worksheet_Hilltop_HT_TW_PNL_DSP.autofilter('A1:E20000')
# worksheet_Hilltop_HT_TW_PNL_DSP.hide_gridlines(2)


# worksheet_Hilltop_Adj_Unreal_PNL_Change.freeze_panes(1, 1)
worksheet_Hilltop_Adj_Unreal_PNL_Change.autofilter('A1:D20000')
# worksheet_Hilltop_Adj_Unreal_PNL_Change.hide_gridlines(2)

# worksheet_Hilltop_Requirement_Change.freeze_panes(1, 1)
worksheet_Hilltop_Requirement_Change.autofilter('A1:D20000')
# worksheet_Hilltop_Requirement_Change.hide_gridlines(2)

worksheet_summary.hide_gridlines(2)
# worksheet_TW_Detail = writer.sheets['TW Detail']


worksheet.set_column('A:A', 34,format1)#, format7) #2
worksheet.set_column('B:B', 13,format1)#, format2) #2
worksheet.set_column('C:C', 14,format1)#, format2) #3
worksheet.set_column('D:D', 11,format2)#, format1)#4
worksheet.set_column('E:E', 11,format1)#, format1)#5
worksheet.set_column('F:F', 11,format1)#, format1)#6
worksheet.set_column('G:G', 11,format1)#, format1)#7
worksheet.set_column('H:H', 11,format1)#, format1)#8
worksheet.set_column('I:I', 11,format1)#, format1)#9
worksheet.set_column('J:J', 11,format1)#, format1)#10
worksheet.set_column('K:K', 11,format1)#, format1)#11
worksheet.set_column('L:L', 11,format1)#, format1)#12
worksheet.set_column('M:M', 11,format1)#, format1)#13
worksheet.set_column('N:N', 11,format1)#, format1)#14
worksheet.set_column('O:O', 11,format1)#, format1)#15
worksheet.set_column('P:P', 15,format1)

worksheet.write('A1', 'Security',format_top_summary)#,format5)
worksheet.write('B1', 'Cusip',format_top_summary)#,format5)
worksheet.write('C1', 'Account',format_top_summary)#,format5)
worksheet.write('D1', 'Price',format_top_summary)#,format5)
worksheet.write('E1', 'TW QTY ',format_top_summary)#,format5)
worksheet.write('F1', 'HT QTY',format_top_summary)#,format5)
worksheet.write('G1', 'QTY Discrepancy',format_top_summary)#,format5)
worksheet.write('H1', 'HT QTY Change',format_top_summary)#,format5)
worksheet.write('I1', 'HT New Unreal PNL',format_top_summary)#,format5)
worksheet.write('J1', 'HT Old Unreal PNL',format_top_summary)#,format5)
worksheet.write('K1', 'Real PNL Change',format_top_summary)#,format5)
worksheet.write('L1', 'Adj Unreal PNL Change',format_top_summary)#,format5)
worksheet.write('M1', 'TW PNL',format_top_summary)#,format5)
worksheet.write('N1', 'HT-TW PNL Discrep.',format_top_summary)#,format5)
worksheet.write('O1', 'Req. Change',format_top_summary)#,format5)
worksheet.write('P1', 'Requirement',format_top_summary)#,format5)

worksheet.set_zoom(90)

worksheet.autofilter('A1:O20000')

TW_Detail.reset_index(inplace = True)
TW_Detail.to_excel(writer,sheet_name = 'TW Detail',index = False)
worksheet_TW_Detail = writer.sheets['TW Detail']
worksheet_TW_Detail.write('A1', 'Cusip',format_top_summary)#,format5)
worksheet_TW_Detail.write('B1', 'P&L',format_top_summary)#,format5)
worksheet_TW_Detail.write('C1', 'Security',format_top_summary)#,format5)
worksheet_TW_Detail.write('D1', 'Position',format_top_summary)#,format5)
worksheet_TW_Detail.write('E1', 'Symbol',format_top_summary)#,format5)
worksheet_TW_Detail.write('F1', 'Book',format_top_summary)#,format5)
worksheet_TW_Detail.write('G1', 'MTG Position',format_top_summary)#,format5)
worksheet_TW_Detail.set_column('A:A', 12,format1)#, format7) #2
worksheet_TW_Detail.set_column('B:B', 12,format1)#, format2) #2
worksheet_TW_Detail.set_column('C:C', 25,format1)#, format7) #3
worksheet_TW_Detail.set_column('D:D', 12,format1)#, format1)#4
worksheet_TW_Detail.set_column('E:E', 15,format1)#, format1)#4
worksheet_TW_Detail.set_column('F:F', 12,format1)#, format1)#4
worksheet_TW_Detail.set_column('G:G', 12,format1)#, format1)#4
worksheet_TW_Detail.autofilter('A1:G20000')
worksheet_TW_Detail.set_zoom(90)

workbook.close()





"""






Email Ready Version







"""



def subject_name(file_text):
    now = datetime.datetime.now()
    weekday = now.weekday()
    if weekday == 0:
        today = now - datetime.timedelta(days=3)
        yesterday = today - datetime.timedelta(days=1)
    elif weekday == 1:
        today = now - datetime.timedelta(days=1)
        yesterday = today - datetime.timedelta(days=3)
    else:
        today = now - datetime.timedelta(days=1)
        yesterday = today - datetime.timedelta(days=1)
    current = today.strftime("%Y-%m-%d")
    current = str(file_text)+str(current)
    PNL_Report_Date = today.strftime("%m.%d.%Y")
    PNL_Report_Date = str(PNL_Report_Date)+'PNL Discrepancy'
    now = datetime.datetime.now()
    PNL_Report_Write_to_Date = now.strftime('%m.%d.%Y')
    PNL_Report_Write_to_Date = str(PNL_Report_Write_to_Date)+'PNL Discrepancy'
    today = file_text + str(today)
    yesterday = yesterday.strftime("%Y-%m-%d")
    yesterday = file_text + str(yesterday)
    return current,yesterday,PNL_Report_Date,PNL_Report_Write_to_Date



file_text = 'Inventory Margin Report for '
x = subject_name(file_text)

"""
Pull and Generate new QTY_DSP_Cleared_Positions file
"""
PNL_Report_Date = x[2]
PNL_Report_File_Most_Recent = 'P:/2. Corps/PNL_Daily_Report/Reports/PNL_Report.xlsx'
Cleared_Yesterday_PNL_Report = pd.read_excel(PNL_Report_File_Most_Recent,sheet_name = 'Quantity Diff')
Cleared_Yesterday_PNL_Report = Cleared_Yesterday_PNL_Report[['Security','Account','Cusip','QTY DSP','Position Notes']]
Cleared_Yesterday_PNL_Report.dropna(inplace = True)
QTY_DSP_Cleared_Positions = 'P:/2. Corps/PNL_Daily_Report/Cleared_Position_File/QTY_DSP_Cleared_Positions.xlsx'
QTY_DSP_Cleared_Positions= pd.read_excel(QTY_DSP_Cleared_Positions,index = False)
QTY_DSP_Cleared_Positions = QTY_DSP_Cleared_Positions.append(Cleared_Yesterday_PNL_Report)

QTY_DSP_Cleared_Positions.drop_duplicates(keep='first',inplace = True)
writer = pd.ExcelWriter('P:/2. Corps/PNL_Daily_Report/Cleared_Position_File/QTY_DSP_Cleared_Positions.xlsx', engine='xlsxwriter')
QTY_DSP_Cleared_Positions.to_excel(writer)
writer.save()


today = x[0]
yesterday = x[1]
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder(6) # "6" refers to the index of a folder - in this case the inbox. You can change that number to reference
messages = inbox.Items
messages.Sort("[ReceivedTime]", True)
message = messages.GetFirst()

i = 0
while i < 20:
    if message.Subject == today:
        try:
            attachments = message.Attachments
            attachment = attachments.Item(1)
            attachment.SaveASFile('P:/2. Corps/PNL_Daily_Report/HT_Files/' + str(today)+'.xlsx') #Saves to the attachment to current folder
            print('HT File Found')
            message = messages.GetNext()

        except:
            message = messages.GetNext()
    i += 1
    message = messages.GetNext()


file_text = 'Report "TW 16 22" is generated.'
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder(6) # "6" refers to the index of a folder - in this case the inbox. You can change that number to reference
messages = inbox.Items
messages.Sort("[ReceivedTime]", True)
message = messages.GetFirst()
i = 0
while i < 10:
   
    if message.Subject == file_text:
        print(message.Subject)
    
        try:
            attachments = message.Attachments
            attachment = attachments.Item(2)
            attachment.SaveASFile('C:\\Users\\ccraig\\Desktop\\PNL Project' + '\\Bloomberg TW'+'.xls') #Saves to the attachment to current folder
            message = messages.GetNext()

        except:
            message = messages.GetNext()
    i += 1
    message = messages.GetNext()
    
Bloomberg_Inventory_16_22 = pd.read_table('C:/Users/ccraig/Desktop/PNL Project/Bloomberg TW.xls')
Bloomberg_Inventory_16_22 = Bloomberg_Inventory_16_22.transpose()
Bloomberg_Inventory_16_22.reset_index(inplace = True)
Bloomberg_Inventory_16_22.drop('index',axis = 1, inplace = True)
Bloomberg_Inventory_16_22.drop(0,axis = 1, inplace = True)
Bloomberg_Inventory_16_22 = Bloomberg_Inventory_16_22.transpose()
Bloomberg_Inventory_16_22 = Bloomberg_Inventory_16_22[:-1]
Bloomberg_Inventory_16_22 = Bloomberg_Inventory_16_22[2:]
Bloomberg_Inventory_16_22 = Bloomberg_Inventory_16_22[0].str.split(',',expand=True)
Bloomberg_Inventory_16_22.rename(columns={0: 'Security'}, inplace=True)
Bloomberg_Inventory_16_22.rename(columns={1: 'P&L'},inplace =True)
Bloomberg_Inventory_16_22.rename(columns={2: 'Cusip'},inplace =True)
Bloomberg_Inventory_16_22.rename(columns={3: 'Position'},inplace =True)
Bloomberg_Inventory_16_22.rename(columns={4: 'Symbol'},inplace =True)
Bloomberg_Inventory_16_22.rename(columns={5: 'Book'},inplace =True)
Bloomberg_Inventory_16_22.rename(columns={6: 'MTG Position'},inplace =True)
Bloomberg_Inventory_16_22 = correct_position_type(Bloomberg_Inventory_16_22)

file_text = 'Report "TW 1 5" is generated.'
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder(6) # "6" refers to the index of a folder - in this case the inbox. You can change that number to reference
messages = inbox.Items
messages.Sort("[ReceivedTime]", True)
message = messages.GetFirst()
i = 0
while i < 10:

    if message.Subject == file_text:
        print(message.Subject)
    
        try:
            attachments = message.Attachments
            attachment = attachments.Item(2)
            attachment.SaveASFile('C:\\Users\\ccraig\\Desktop\\PNL Project' + '\\Bloomberg TW'+'.xls') #Saves to the attachment to current folder
            message = messages.GetNext()

        except:
            message = messages.GetNext()
    i += 1
    message = messages.GetNext()
    
Bloomberg_Inventory_1_5 = pd.read_table('C:/Users/ccraig/Desktop/PNL Project/Bloomberg TW.xls')
Bloomberg_Inventory_1_5 = Bloomberg_Inventory_1_5.transpose()
Bloomberg_Inventory_1_5.reset_index(inplace = True)
Bloomberg_Inventory_1_5.drop('index',axis = 1, inplace = True)
Bloomberg_Inventory_1_5.drop(0,axis = 1, inplace = True)
Bloomberg_Inventory_1_5 = Bloomberg_Inventory_1_5.transpose()
Bloomberg_Inventory_1_5 = Bloomberg_Inventory_1_5[:-1]
Bloomberg_Inventory_1_5 = Bloomberg_Inventory_1_5[2:]
Bloomberg_Inventory_1_5 = Bloomberg_Inventory_1_5[0].str.split(',',expand=True)
Bloomberg_Inventory_1_5.rename(columns={0: 'Security'}, inplace=True)
Bloomberg_Inventory_1_5.rename(columns={1: 'P&L'},inplace =True)
Bloomberg_Inventory_1_5.rename(columns={2: 'Cusip'},inplace =True)
Bloomberg_Inventory_1_5.rename(columns={3: 'Position'},inplace =True)
Bloomberg_Inventory_1_5.rename(columns={4: 'Symbol'},inplace =True)
Bloomberg_Inventory_1_5.rename(columns={5: 'Book'},inplace =True)
Bloomberg_Inventory_1_5.rename(columns={6: 'MTG Position'},inplace =True)
Bloomberg_Inventory_1_5 = correct_position_type(Bloomberg_Inventory_1_5)

file_text = 'Report "TW 6 10" is generated.'
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder(6) # "6" refers to the index of a folder - in this case the inbox. You can change that number to reference
messages = inbox.Items
messages.Sort("[ReceivedTime]", True)
message = messages.GetFirst()
i = 0
while i < 10:
   
    if message.Subject == file_text:
        print(message.Subject)
    
        try:
            attachments = message.Attachments
            attachment = attachments.Item(2)
            attachment.SaveASFile('C:\\Users\\ccraig\\Desktop\\PNL Project' + '\\Bloomberg TW'+'.xls') #Saves to the attachment to current folder
            message = messages.GetNext()

        except:
            message = messages.GetNext()
    i += 1
    message = messages.GetNext()
    
Bloomberg_Inventory_6_10 = pd.read_table('C:/Users/ccraig/Desktop/PNL Project/Bloomberg TW.xls')
Bloomberg_Inventory_6_10 = Bloomberg_Inventory_6_10.transpose()
Bloomberg_Inventory_6_10.reset_index(inplace = True)
Bloomberg_Inventory_6_10.drop('index',axis = 1, inplace = True)
Bloomberg_Inventory_6_10.drop(0,axis = 1, inplace = True)
Bloomberg_Inventory_6_10 = Bloomberg_Inventory_6_10.transpose()
Bloomberg_Inventory_6_10 = Bloomberg_Inventory_6_10[:-1]
Bloomberg_Inventory_6_10 = Bloomberg_Inventory_6_10[2:]
Bloomberg_Inventory_6_10 = Bloomberg_Inventory_6_10[0].str.split(',',expand=True)
Bloomberg_Inventory_6_10.rename(columns={0: 'Security'}, inplace=True)
Bloomberg_Inventory_6_10.rename(columns={1: 'P&L'},inplace =True)
Bloomberg_Inventory_6_10.rename(columns={2: 'Cusip'},inplace =True)
Bloomberg_Inventory_6_10.rename(columns={3: 'Position'},inplace =True)
Bloomberg_Inventory_6_10.rename(columns={4: 'Symbol'},inplace =True)
Bloomberg_Inventory_6_10.rename(columns={5: 'Book'},inplace =True)
Bloomberg_Inventory_6_10.rename(columns={6: 'MTG Position'},inplace =True)
Bloomberg_Inventory_6_10 = correct_position_type(Bloomberg_Inventory_6_10)

file_text = 'Report "TW 11 15" is generated.'
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder(6) # "6" refers to the index of a folder - in this case the inbox. You can change that number to reference
messages = inbox.Items
messages.Sort("[ReceivedTime]", True)
message = messages.GetFirst()
i = 0
while i < 10:
   
    if message.Subject == file_text:
        print(message.Subject)
        try:
            attachments = message.Attachments
            attachment = attachments.Item(2)
            attachment.SaveASFile('C:\\Users\\ccraig\\Desktop\\PNL Project' + '\\Bloomberg TW'+'.xls') #Saves to the attachment to current folder
            message = messages.GetNext()

        except:
            message = messages.GetNext()
    i += 1
    message = messages.GetNext()
    
Bloomberg_Inventory_11_15 = pd.read_table('C:/Users/ccraig/Desktop/PNL Project/Bloomberg TW.xls')
Bloomberg_Inventory_11_15 = Bloomberg_Inventory_11_15.transpose()
Bloomberg_Inventory_11_15.reset_index(inplace = True)
Bloomberg_Inventory_11_15.drop('index',axis = 1, inplace = True)
Bloomberg_Inventory_11_15.drop(0,axis = 1, inplace = True)
Bloomberg_Inventory_11_15 = Bloomberg_Inventory_11_15.transpose()
Bloomberg_Inventory_11_15 = Bloomberg_Inventory_11_15[:-1]
Bloomberg_Inventory_11_15 = Bloomberg_Inventory_11_15[2:]
Bloomberg_Inventory_11_15 = Bloomberg_Inventory_11_15[0].str.split(',',expand=True)
Bloomberg_Inventory_11_15.rename(columns={0: 'Security'}, inplace=True)
Bloomberg_Inventory_11_15.rename(columns={1: 'P&L'},inplace =True)
Bloomberg_Inventory_11_15.rename(columns={2: 'Cusip'},inplace =True)
Bloomberg_Inventory_11_15.rename(columns={3: 'Position'},inplace =True)
Bloomberg_Inventory_11_15.rename(columns={4: 'Symbol'},inplace =True)
Bloomberg_Inventory_11_15.rename(columns={5: 'Book'},inplace =True)
Bloomberg_Inventory_11_15.rename(columns={6: 'MTG Position'},inplace =True)
Bloomberg_Inventory_11_15 = correct_position_type(Bloomberg_Inventory_11_15)

Bloomberg_Inventory = pd.concat([Bloomberg_Inventory_16_22,
                                 Bloomberg_Inventory_1_5,
                                 Bloomberg_Inventory_6_10,
                                 Bloomberg_Inventory_11_15],ignore_index = True)

"""
Read in Excel files from Hilltop and Bloomberg

"""

'P:/2. Corps/PNL_Daily_Report/HT_Files/' + str(today)+'.xlsx'

Bloomberg_Inventory['Cusip'] = Bloomberg_Inventory['Symbol'].str[1:10]
Bloomberg_Inventory = Bloomberg_Inventory[['Cusip', 'P&L', 'Security', 'Position','Symbol','Book','MTG Position']]
Bloomberg_Inventory['Position'] = Bloomberg_Inventory['Position']*1000
Bloomberg_Inventory['MTG Position'] = Bloomberg_Inventory['MTG Position']*1000
# Bloomberg_Inventory['Cusip'] = Bloomberg_Inventory['Cusip'].astype(str)
print(today)
print(yesterday)
Recent = 'P:/2. Corps/PNL_Daily_Report/HT_Files/' + str(today)+'.xlsx'#Recent = 'C:/Users/ccraig/Desktop/PNL Project/'+str(today)+'.xlsx'
Old = 'P:/2. Corps/PNL_Daily_Report/HT_Files/' + str(yesterday)+'.xlsx'
Hilltop_Recent_x = pd.read_excel(io=Recent, sheet_name='Detail')
# Hilltop_Recent_x['Cusip'] = Hilltop_Recent['Cusip'].astype(str)
Hilltop_Old_y = pd.read_excel(io=Old, sheet_name='Detail')
# Hilltop_Old_y['Cusip'] = Hilltop_Old_y['Cusip'].astype(str)
Hilltop_Recent_s = pd.read_excel(io=Recent, sheet_name='Summary')
Hilltop_Old_s = pd.read_excel(io=Old, sheet_name='Summary')
Hilltop_Recent_s = Hilltop_Recent_s.head(10)
Hilltop_Old_s = Hilltop_Old_s.head(10)
Hilltop_Recent_x['Cusip_group_by'] = Hilltop_Recent_x['Cusip']
Hilltop_Recent_x['Cusip_group_by'] = 'C'+ Hilltop_Recent_x['Cusip_group_by']
TW_Detail = Bloomberg_Inventory
Bloomberg_Inventory = Bloomberg_Inventory.groupby(['Cusip']).agg({'P&L':'sum',
                                                                   'Security':'first',
                                                                   'Position':'sum',
                                                                   'Symbol':'first',
                                                                   'Book':'first',
                                                                   'MTG Position':'sum'})
"""
Fix MTG Position
"""
# Bloomberg_Inventory.loc[(Bloomberg_Inventory['Book'] == '8763') | (Bloomberg_Inventory['Book']=='IO'), 'Position'] = 'MTG Position'


Hilltop_Recent = Hilltop_Recent_x.groupby(['Cusip_group_by']).agg({'Quantity':'sum',
                                                                   'Unreal PNL':'sum',
                                                                   'Real PNL':'sum',
                                                                   'Requirement':'sum',
                                                                   'Cusip':'first',
                                                                   'Description':'first',
                                                                   'Price':'mean'})

Hilltop_Old_y['Cusip_group_by'] = Hilltop_Old_y['Cusip']
Hilltop_Old = Hilltop_Old_y.groupby(['Cusip_group_by']).agg({'Quantity':'sum',
                                                             'Unreal PNL':'sum',
                                                             'Real PNL':'sum',
                                                             'Requirement':'sum',
                                                             'Cusip':'first',
                                                             'Description':'first',
                                                             'Price':'mean'})

"""
Merge Hilltop Recent and Old together

"""

Hilltop_Recent = pd.merge(Hilltop_Recent, Hilltop_Old, on='Cusip', how='left')

Hilltop_Recent = pd.merge(Hilltop_Recent, Bloomberg_Inventory, on='Cusip', how='outer')

Hilltop_Recent_x = Hilltop_Recent_x[['Cusip','Account Name']]

Hilltop_Recent = pd.merge(Hilltop_Recent, Hilltop_Recent_x,on='Cusip',how='outer')

Hilltop_Recent = Hilltop_Recent.fillna(0)
Hilltop_Recent.loc[Hilltop_Recent['Security']==0,'Security'] = Hilltop_Recent['Description_x']

"""
Fix Account Names returning 0
"""
Hilltop_Recent.loc[(Hilltop_Recent['Account Name'] == 0) & (Hilltop_Recent['Book']=='8701'), 'Account Name'] = 'K74 Corporates'
Hilltop_Recent.loc[(Hilltop_Recent['Account Name'] == 0) & (Hilltop_Recent['Book']=='CORPSP'), 'Account Name'] = 'P02 Corp SP'
Hilltop_Recent.loc[(Hilltop_Recent['Account Name'] == 0) & (Hilltop_Recent['Book']=='CORPFRN'), 'Account Name'] = 'P01 Corp Floate'
Hilltop_Recent.loc[(Hilltop_Recent['Account Name'] == 0) & (Hilltop_Recent['Book']=='8503'), 'Account Name'] = 'K72 Muni Inv FI'
Hilltop_Recent.loc[(Hilltop_Recent['Account Name'] == 0) & (Hilltop_Recent['Book']=='8763'), 'Account Name'] = 'K76 S P Inv'
Hilltop_Recent.loc[(Hilltop_Recent['Account Name'] == 0) & (Hilltop_Recent['Book']=='8782'), 'Account Name'] = 'K77 CD Inv'
Hilltop_Recent.loc[(Hilltop_Recent['Account Name'] == 0) & (Hilltop_Recent['Book']=='8937'), 'Account Name'] = 'K78 Taxable Mun'
Hilltop_Recent.loc[(Hilltop_Recent['Account Name'] == 0) & (Hilltop_Recent['Book']=='8938'), 'Account Name'] = 'K79 Cali Tax Ex'
Hilltop_Recent.loc[(Hilltop_Recent['Account Name'] == 0) & (Hilltop_Recent['Book']=='8939'), 'Account Name'] = 'K80 Muni Tax Ex'
Hilltop_Recent.loc[(Hilltop_Recent['Account Name'] == 0) & (Hilltop_Recent['Book']=='8940'), 'Account Name'] = 'K81 Muni Tax'
Hilltop_Recent.loc[(Hilltop_Recent['Account Name'] == 0) & (Hilltop_Recent['Book']=='8941'), 'Account Name'] = 'K82 Tax 0 Muni'
Hilltop_Recent.loc[(Hilltop_Recent['Account Name'] == 0) & (Hilltop_Recent['Book']=='CORPIG'), 'Account Name'] = 'N88 Corp Notes'
Hilltop_Recent.loc[(Hilltop_Recent['Account Name'] == 0) & (Hilltop_Recent['Book']=='CORPNOTE'), 'Account Name'] = 'N90 CD'
Hilltop_Recent.loc[(Hilltop_Recent['Account Name'] == 0) & (Hilltop_Recent['Book']=='IO'), 'Account Name'] = 'M64 Sierra MBS'
Hilltop_Recent = Hilltop_Recent[Hilltop_Recent['Book'] != '8659']
Hilltop_Recent = Hilltop_Recent[Hilltop_Recent['Book'] != '8720']
Hilltop_Recent['Position'] = np.where((Hilltop_Recent['Account Name']== 'K76 S P Inv'),Hilltop_Recent['MTG Position'],Hilltop_Recent['Position'])

"""
Calculate nessicary values

"""

Hilltop_Recent['Real_PNL_Change'] = Hilltop_Recent['Real PNL_x'] - Hilltop_Recent['Real PNL_y']
Hilltop_Recent['Real Discrepancy'] = Hilltop_Recent['Real_PNL_Change'] - Hilltop_Recent['P&L']
Hilltop_Recent['Quantity Change'] = Hilltop_Recent['Quantity_x'] - Hilltop_Recent['Quantity_y']

Hilltop_Individual_Book_Summary = Hilltop_Recent

Hilltop_Recent.rename(columns={'Quantity Change': 'HT Quantity Change',
                               'Quantity_x': 'HT New Quantity',
                               'Quantity_y': 'HT Old Quantity',
                               'Quantity_x_y': '-  HT Quantity  =',
                               'Real Discrepancy_y': 'TW vs. HT Real Discrepancy',
                               'Position': 'TW Quantity',
                               'Real PNL_x': 'HT New Real PNL',
                               'Account Name': 'Account',
                               'Quantity_y':'HT Old Quantity',
                               'Unreal PNL_x':'HT New Unreal PNL',
                               'Unreal PNL_y':'HT Old Unreal PNL',
                               'Real PNL_y':'HT Old Real PNL',
                               'P&L':'TW PNL',
                               'Quantity_y':'HT Old Quantity',
                               'Price_x':'Price'}, inplace=True)




# Hilltop_Recent['TW Quantity'] = Hilltop_Recent['TW Quantity'] * 1000

Hilltop_Recent['HT Change in Quantity'] = Hilltop_Recent['HT New Quantity']-Hilltop_Recent['HT Old Quantity']
Hilltop_Recent['TW Quantity'] = pd.to_numeric(Hilltop_Recent['TW Quantity'])
Hilltop_Recent['HT New Quantity'] = pd.to_numeric(Hilltop_Recent['HT New Quantity'])
Hilltop_Recent['TW - HT Quantity Discrepancy'] = Hilltop_Recent['TW Quantity']-Hilltop_Recent['HT New Quantity']
Hilltop_Recent['Real PNL Change'] = Hilltop_Recent['HT New Real PNL']-Hilltop_Recent['HT Old Real PNL']
Hilltop_Recent['HT Real PNL Change'] = Hilltop_Recent['HT New Real PNL']-Hilltop_Recent['HT Old Real PNL']
Hilltop_Recent['Adj Unreal PNL Change'] = Hilltop_Recent['HT New Unreal PNL']-Hilltop_Recent['HT Old Unreal PNL']+Hilltop_Recent['HT Real PNL Change']
Hilltop_Recent['HT-TW PNL Discrepancy'] = Hilltop_Recent['HT Real PNL Change']-Hilltop_Recent['TW PNL']
Hilltop_Recent['Requirement Change'] = Hilltop_Recent['Requirement_x']-Hilltop_Recent['Requirement_y']
Hilltop_Recent['Filter Column'] = Hilltop_Recent['TW - HT Quantity Discrepancy'] + Hilltop_Recent['HT Change in Quantity'] + Hilltop_Recent['Adj Unreal PNL Change'] + Hilltop_Recent['HT-TW PNL Discrepancy'] + Hilltop_Recent['Requirement Change']
# Hilltop_Recent = Hilltop_Recent[(Hilltop_Recent['Filter Column'] != 0)]
Hilltop_Recent = pd.merge(HT_Detail,Hilltop_Recent, on='Cusip', how='left')
Hilltop_Recent = Hilltop_Recent[[
                                 'HT Quantity Change',                         
                                 'Security',                                       #A
                                 'Cusip',                                          #B
                                 'Account',                                        #C
                                 'Price_x',                                          #D
                                 'TW Quantity',                                    #E
                                 'HT New Quantity',                                #F
                                 'TW - HT Quantity Discrepancy',                   #G
                                 'HT Change in Quantity',                          #H
                                 'HT New Unreal PNL',                              #I
                                 'HT Old Unreal PNL',                              #J
                                 'Real PNL Change',                                #K
                                 'Adj Unreal PNL Change',                          #L
                                 'TW PNL',                                         #M
                                 'HT-TW PNL Discrepancy',                          #N
                                 'Requirement Change',
                                 'Requirement_x'
]]
Hilltop_Recent.dropna(thresh = 5,inplace = True)  

"""
# Drop Duplicated Values for Cusip and HT Quantity Change

"""
Hilltop_Recent = Hilltop_Recent.drop_duplicates(['Cusip', 'HT Quantity Change'])


"""
# Set up excel file naming path w/ today's date

"""
time = datetime.datetime.today()
current = time.strftime("%m.%d.%Y")
"""
# Write file to excel

"""
Hilltop_Recent.drop(
    [
        'HT Quantity Change'
    ],
    axis=1, inplace=True)


Hilltop_Individual_Summary = Hilltop_Recent
Hilltop_Recent.sort_values('TW PNL', axis=0, ascending=False, inplace=True)

filepath = 'P:/2. Corps/PNL_Daily_Report/Reports/'+ str(current) + ' Inventory Report.xlsx'


writer = pd.ExcelWriter(filepath, engine='xlsxwriter')
"""
            Summary Code

"""
Daily_Change_x = pd.read_excel(io=Recent, sheet_name='Summary')
Daily_Change_y = pd.read_excel(io=Old,sheet_name='Summary')
Daily_Change_x = Daily_Change_x[12:]
Daily_Change_y = Daily_Change_y[12:]
Daily_Change_x.rename(columns={'Account Number': 'Inventory Sub Totals',
                                'Total Available Funds':'Account Name',
                                'Unnamed: 2':'Position Type',
                                'Unnamed: 3':'Cost',
                                'Unnamed: 4':'Market Value',
                                'Unnamed: 5':'Unreal PNL',
                                'Unnamed: 6':'Requirement',
                                'Unnamed: 7':'Real PNL'}, inplace=True)
Daily_Change_x.reset_index(inplace = True)
Daily_Change_y.rename(columns={'Account Number': 'Inventory Sub Totals',
                                'Total Available Funds':'Account Name',
                                'Unnamed: 2':'Position Type',
                                'Unnamed: 3':'Cost',
                                'Unnamed: 4':'Market Value',
                                'Unnamed: 5':'Unreal PNL',
                                'Unnamed: 6':'Requirement',
                                'Unnamed: 7':'Real PNL'}, inplace=True)

Daily_Change_y.reset_index(inplace = True)
Daily_Change_x=pd.merge(Daily_Change_x, Daily_Change_y, on='index', how='left')
Daily_Change_x['Cost'] = Daily_Change_x['Cost_x']-Daily_Change_x['Cost_y']
Daily_Change_x['Market Value']=Daily_Change_x['Market Value_x']-Daily_Change_x['Market Value_y']
Daily_Change_x['Requirement']=Daily_Change_x['Requirement_x']-Daily_Change_x['Requirement_y']
Daily_Change_x['Unreal PNL']=Daily_Change_x['Unreal PNL_x']-Daily_Change_x['Unreal PNL_y']
Daily_Change_x['Real PNL']=Daily_Change_x['Real PNL_x']-Daily_Change_x['Real PNL_y']
Daily_Change_x = Daily_Change_x[['index','Account Name_x','Position Type_x','Cost','Market Value','Requirement','Unreal PNL','Real PNL']]
Daily_Change = Daily_Change_x.groupby('Account Name_x')['Cost','Market Value','Requirement', 'Unreal PNL','Real PNL'].sum()
Daily_Change['Account Name_x']=['K72 Muni Inv','K74 Corporates ','K76 S P Inv',
                                'K77 CD Inv','K78 Taxable Mun',
                                'K79 Cali Tax Ex','K80 Muni Tax Ex',
                                'K81 Muni Tax','K82 Tax 0 Muni','L81 Sierra Comp',
                                'M64 Sierra MBS','N90 CD','P01 Corp Floate','P02 Corp Sp',
                                'N88 Corp Notes','N87 Corp HY','P03 New Corp 01']
Daily_Change = Daily_Change_x.append(Daily_Change, ignore_index = True,sort = False)
Daily_Change['Account Name_x'] = Daily_Change['Account Name_x'].replace({'K72 Muni Inv':'K72 Muni Inv Fl'})
Daily_Change.sort_values(['Account Name_x','Position Type_x'],inplace = True)
Daily_Change.rename(columns={'Account Name_x':'Account Name','Position Type_x':'Position Type'},inplace = True)

"""
Create and Format Summary sheet

# """

# instructions = 'Instructions:\n1.
"""
            Summary Code

"""
Summary_Recent = pd.read_excel(io=Recent, sheet_name='Summary')
Summary_Old = pd.read_excel(io=Old,sheet_name='Summary')
Summary_Recent = Summary_Recent[12:]
Summary_Old = Summary_Old[12:]
Summary_Recent.rename(columns={'Account Number': 'Inventory Sub Totals',
                                'Total Available Funds':'Account Name',
                                'Unnamed: 2':'Position Type',
                                'Unnamed: 3':'Cost',
                                'Unnamed: 4':'Market Value',
                                'Unnamed: 5':'Unreal PNL',
                                'Unnamed: 6':'Requirement',
                                'Unnamed: 7':'Real PNL'}, inplace=True)
Summary_Recent.reset_index(inplace = True)
Summary_Old.rename(columns={'Account Number': 'Inventory Sub Totals',
                                'Total Available Funds':'Account Name',
                                'Unnamed: 2':'Position Type',
                                'Unnamed: 3':'Cost',
                                'Unnamed: 4':'Market Value',
                                'Unnamed: 5':'Unreal PNL',
                                'Unnamed: 6':'Requirement',
                                'Unnamed: 7':'Real PNL'}, inplace=True)
Summary_Old.reset_index(inplace = True)

Daily_Change=pd.merge(Summary_Recent, Summary_Old, on='index', how='left')

Daily_Change['Cost'] = Daily_Change['Cost_x']-Daily_Change['Cost_y']
Daily_Change['Market Value']=Daily_Change['Market Value_x']-Daily_Change['Market Value_y']
Daily_Change['Requirement']=Daily_Change['Requirement_x']-Daily_Change['Requirement_y']
Daily_Change['Unreal PNL']=Daily_Change['Unreal PNL_x']-Daily_Change['Unreal PNL_y']
Daily_Change['Real PNL']=Daily_Change['Real PNL_x']-Daily_Change['Real PNL_y']
Daily_Change = Daily_Change[['Account Name_x','Position Type_x','Cost','Market Value','Requirement','Unreal PNL','Real PNL']]
Daily_Change.rename(columns={'Account Name_x': 'Account Name',
                                'Position Type_x':'Position Type'
                            }, inplace=True)
Summary_Recent = Summary_Recent[['Account Name','Position Type','Cost','Market Value','Requirement','Unreal PNL','Real PNL']]
Muni_Summary_Recent = Summary_Recent.reindex([0,1,8,9,10,11,12,13,14,15,16,17]) 
Muni_Daily_Change =  Daily_Change.reindex([0,1,8,9,10,11,12,13,14,15,16,17])
Muni_Summary_Recent_Grouped = Muni_Summary_Recent.groupby(['Account Name']).agg({'Position Type':'first',
                                                                                 'Cost':'sum',
                                                                     'Market Value':'sum',
                                                                    'Requirement':'sum',
                                                                    'Unreal PNL':'sum',
                                                                    'Real PNL':'sum'})
Muni_Summary_Recent.reset_index(inplace = True)
Muni_Summary_Recent_Grouped.reset_index(inplace = True)
Muni_Summary_Recent_Short = Muni_Summary_Recent[Muni_Summary_Recent['Position Type'] == 'Short']

Muni_Summary_Recent_Short['Short Total'] = abs(Muni_Summary_Recent_Short['Cost'] + Muni_Summary_Recent_Short['Market Value'] + Muni_Summary_Recent_Short['Requirement'] + Muni_Summary_Recent_Short['Unreal PNL'] + 
                                             Muni_Summary_Recent_Short['Real PNL'])
Muni_Summary_Recent_Short = Muni_Summary_Recent_Short[Muni_Summary_Recent_Short['Position Type'] == 'Short']
Muni_Summary_Recent_Short = Muni_Summary_Recent_Short[Muni_Summary_Recent_Short['Short Total'] > 0]
Muni_Summary_Recent_Short = Muni_Summary_Recent_Short['Account Name'].tolist()


Muni_Daily_Change_Grouped = Muni_Daily_Change.groupby(['Account Name']).agg({'Position Type':'first',
                                                                             'Cost':'sum',
                                                                     'Market Value':'sum',
                                                                    'Requirement':'sum',
                                                                    'Unreal PNL':'sum',
                                                                    'Real PNL':'sum'})
Muni_Daily_Change.reset_index(inplace = True)
Muni_Daily_Change_Grouped.reset_index(inplace = True)
Muni_Daily_Change_Short = Muni_Daily_Change[Muni_Daily_Change['Position Type'] == 'Short']

Muni_Daily_Change_Short['Short Total'] = abs(Muni_Daily_Change_Short['Cost'] + Muni_Daily_Change_Short['Market Value'] + Muni_Daily_Change_Short['Requirement'] + Muni_Daily_Change_Short['Unreal PNL'] + 
                                             Muni_Daily_Change_Short['Real PNL'])
Muni_Daily_Change_Short = Muni_Daily_Change_Short[Muni_Daily_Change_Short['Position Type'] == 'Short']
Muni_Daily_Change_Short = Muni_Daily_Change_Short[Muni_Daily_Change_Short['Short Total'] > 0]
Muni_Daily_Change_Short = Muni_Daily_Change_Short['Account Name'].tolist()

N87_Summary_Recent  = Summary_Recent.reindex([22,23])
N87_Daily_Change  = Daily_Change.reindex([22,23])

N88_Summary_Recent  = Summary_Recent.reindex([24,25])
N88_Daily_Change  = Daily_Change.reindex([24,25])

N90_Summary_Recent  = Summary_Recent.reindex([26,27])
N90_Daily_Change  = Daily_Change.reindex([26,27])

P01_Summary_Recent  = Summary_Recent.reindex([28,29])
P01_Daily_Change  = Daily_Change.reindex([28,29])

K74_Summary_Recent  = Summary_Recent.reindex([2,3])
K74_Daily_Change  = Daily_Change.reindex([2,3])

L81_Summary_Recent  = Summary_Recent.reindex([18,19])
L81_Daily_Change  = Daily_Change.reindex([18,19])

P02_Summary_Recent  = Summary_Recent.reindex([30,31])
P02_Daily_Change  = Daily_Change.reindex([30,31])

P03_Summary_Recent  = Summary_Recent.reindex([32,33])
P03_Daily_Change  = Daily_Change.reindex([32,33])

CD_Summary_Recent  = Summary_Recent.reindex([6])
CD_Daily_Change  = Daily_Change.reindex([6])

CMO_Summary_Recent  = Summary_Recent.reindex([4,20])
CMO_Daily_Change  = Daily_Change.reindex([4,20])





"""
Create and Format Summary sheet

# """
Hilltop_Summary_new = pd.read_excel(io=Recent, sheet_name='Detail')
Hilltop_Summary_new = Hilltop_Summary_new.groupby(['Account Name']).agg({
                                                          'Unreal PNL':'sum',
                                                          'Real PNL':'sum',
                                                          'Requirement':'sum',
                                                          'Cost':'sum',
                                                          'Market Value':'sum'})
Hilltop_Summary_new = Hilltop_Summary_new[['Cost','Market Value','Requirement','Unreal PNL','Real PNL']]

Hilltop_Summary_new.loc['Column_Total'] = Hilltop_Summary_new.sum(numeric_only=True, axis=0)



# Daily_Change.to_excel(writer,sheet_name ='Summary',index=False,startrow=20)

Muni_Summary_Recent_Grouped.to_excel(writer,sheet_name = 'Summary',index = False,startrow =11 )
Muni_Daily_Change_Grouped.to_excel(writer,sheet_name = 'Summary',index = False,startrow =11, startcol = 8 )
N88_Summary_Recent.to_excel(writer,sheet_name = 'Summary',index = False,startrow =19 )
N88_Daily_Change.to_excel(writer,sheet_name = 'Summary',index = False,startrow =19, startcol = 8)
N90_Summary_Recent.to_excel(writer,sheet_name = 'Summary',index = False,startrow =23 )
N90_Daily_Change.to_excel(writer,sheet_name = 'Summary',index = False,startrow =23, startcol = 8 )
P01_Summary_Recent.to_excel(writer,sheet_name = 'Summary',index = False,startrow =27 )
P01_Daily_Change.to_excel(writer,sheet_name = 'Summary',index = False,startrow =27, startcol = 8 )
P02_Summary_Recent.to_excel(writer,sheet_name = 'Summary',index = False,startrow =31 )
P02_Daily_Change.to_excel(writer,sheet_name = 'Summary',index = False,startrow =31, startcol = 8 )
K74_Summary_Recent.to_excel(writer,sheet_name = 'Summary',index = False,startrow =35 )
K74_Daily_Change.to_excel(writer,sheet_name = 'Summary',index = False,startrow =35, startcol = 8 )
L81_Summary_Recent.to_excel(writer,sheet_name = 'Summary',index = False,startrow =39 )
L81_Daily_Change.to_excel(writer,sheet_name = 'Summary',index = False,startrow =39, startcol = 8 )
P03_Summary_Recent.to_excel(writer,sheet_name = 'Summary',index = False,startrow =43 )
P03_Daily_Change.to_excel(writer,sheet_name = 'Summary',index = False,startrow =43, startcol = 8 )
N87_Summary_Recent.to_excel(writer,sheet_name = 'Summary',index = False,startrow =47 )
N87_Daily_Change.to_excel(writer,sheet_name = 'Summary',index = False,startrow =47, startcol = 8 )

CD_Summary_Recent.to_excel(writer,sheet_name = 'Summary',index = False,startrow = 52)
CD_Daily_Change.to_excel(writer,sheet_name = 'Summary',index = False,startrow = 52, startcol = 8)
CMO_Summary_Recent.to_excel(writer,sheet_name = 'Summary',index = False,startrow = 54)
CMO_Daily_Change.to_excel(writer,sheet_name = 'Summary',index = False,startrow = 54, startcol = 8)

"""
Calculate Column Totals
"""
#       MUNI
Muni_Cost_Summary_Recent = Muni_Summary_Recent['Cost'].sum()
Muni_Market_Value_Summary_Recent = Muni_Summary_Recent['Market Value'].sum()
Muni_Requirement_Summary_Recent = Muni_Summary_Recent['Requirement'].sum()
Muni_Unreal_PNL_Summary_Recent = Muni_Summary_Recent['Unreal PNL'].sum()
Muni_Real_PNL_Summary_Recent = Muni_Summary_Recent['Real PNL'].sum()

Muni_Cost_Daily_Change = Muni_Daily_Change['Cost'].sum()
Muni_Market_Value_Daily_Change = Muni_Daily_Change['Market Value'].sum()
Muni_Requirement_Daily_Change = Muni_Daily_Change['Requirement'].sum()
Muni_Unreal_PNL_Daily_Change = Muni_Daily_Change['Unreal PNL'].sum()
Muni_Real_PNL_Daily_Change = Muni_Daily_Change['Real PNL'].sum()

#      Corp
Corp_N88_Cost_Summary_Recent = N88_Summary_Recent['Cost'].sum()
Corp_N88_Market_Value_Summary_Recent = N88_Summary_Recent['Market Value'].sum()
Corp_N88_Requirement_Summary_Recent = N88_Summary_Recent['Requirement'].sum()
Corp_N88_Unreal_PNL_Summary_Recent = N88_Summary_Recent['Unreal PNL'].sum()
Corp_N88_Real_PNL_Summary_Recent = N88_Summary_Recent['Real PNL'].sum()

Corp_N88_Cost_Daily_Change = N88_Daily_Change['Cost'].sum()
Corp_N88_Market_Value_Daily_Change = N88_Daily_Change['Market Value'].sum()
Corp_N88_Requirement_Daily_Change = N88_Daily_Change['Requirement'].sum()
Corp_N88_Unreal_PNL_Daily_Change = N88_Daily_Change['Unreal PNL'].sum()
Corp_N88_Real_PNL_Daily_Change = N88_Daily_Change['Real PNL'].sum()


Corp_N90_Cost_Summary_Recent = N90_Summary_Recent['Cost'].sum()
Corp_N90_Market_Value_Summary_Recent = N90_Summary_Recent['Market Value'].sum()
Corp_N90_Requirement_Summary_Recent = N90_Summary_Recent['Requirement'].sum()
Corp_N90_Unreal_PNL_Summary_Recent = N90_Summary_Recent['Unreal PNL'].sum()
Corp_N90_Real_PNL_Summary_Recent = N90_Summary_Recent['Real PNL'].sum()

Corp_N90_Cost_Daily_Change = N90_Daily_Change['Cost'].sum()
Corp_N90_Market_Value_Daily_Change = N90_Daily_Change['Market Value'].sum()
Corp_N90_Requirement_Daily_Change = N90_Daily_Change['Requirement'].sum()
Corp_N90_Unreal_PNL_Daily_Change = N90_Daily_Change['Unreal PNL'].sum()
Corp_N90_Real_PNL_Daily_Change = N90_Daily_Change['Real PNL'].sum()


Corp_P01_Cost_Summary_Recent = P01_Summary_Recent['Cost'].sum()
Corp_P01_Market_Value_Summary_Recent = P01_Summary_Recent['Market Value'].sum()
Corp_P01_Requirement_Summary_Recent = P01_Summary_Recent['Requirement'].sum()
Corp_P01_Unreal_PNL_Summary_Recent = P01_Summary_Recent['Unreal PNL'].sum()
Corp_P01_Real_PNL_Summary_Recent = P01_Summary_Recent['Real PNL'].sum()

Corp_P01_Cost_Daily_Change = P01_Daily_Change['Cost'].sum()
Corp_P01_Market_Value_Daily_Change = P01_Daily_Change['Market Value'].sum()
Corp_P01_Requirement_Daily_Change = P01_Daily_Change['Requirement'].sum()
Corp_P01_Unreal_PNL_Daily_Change = P01_Daily_Change['Unreal PNL'].sum()
Corp_P01_Real_PNL_Daily_Change = P01_Daily_Change['Real PNL'].sum()


Corp_P02_Cost_Summary_Recent = P02_Summary_Recent['Cost'].sum()
Corp_P02_Market_Value_Summary_Recent = P02_Summary_Recent['Market Value'].sum()
Corp_P02_Requirement_Summary_Recent = P02_Summary_Recent['Requirement'].sum()
Corp_P02_Unreal_PNL_Summary_Recent = P02_Summary_Recent['Unreal PNL'].sum()
Corp_P02_Real_PNL_Summary_Recent = P02_Summary_Recent['Real PNL'].sum()

Corp_P02_Cost_Daily_Change = P02_Daily_Change['Cost'].sum()
Corp_P02_Market_Value_Daily_Change = P02_Daily_Change['Market Value'].sum()
Corp_P02_Requirement_Daily_Change = P02_Daily_Change['Requirement'].sum()
Corp_P02_Unreal_PNL_Daily_Change = P02_Daily_Change['Unreal PNL'].sum()
Corp_P02_Real_PNL_Daily_Change = P02_Daily_Change['Real PNL'].sum()


Corp_K74_Cost_Summary_Recent = K74_Summary_Recent['Cost'].sum()
Corp_K74_Market_Value_Summary_Recent = K74_Summary_Recent['Market Value'].sum()
Corp_K74_Requirement_Summary_Recent = K74_Summary_Recent['Requirement'].sum()
Corp_K74_Unreal_PNL_Summary_Recent = K74_Summary_Recent['Unreal PNL'].sum()
Corp_K74_Real_PNL_Summary_Recent = K74_Summary_Recent['Real PNL'].sum()

Corp_K74_Cost_Daily_Change = K74_Daily_Change['Cost'].sum()
Corp_K74_Market_Value_Daily_Change = K74_Daily_Change['Market Value'].sum()
Corp_K74_Requirement_Daily_Change = K74_Daily_Change['Requirement'].sum()
Corp_K74_Unreal_PNL_Daily_Change = K74_Daily_Change['Unreal PNL'].sum()
Corp_K74_Real_PNL_Daily_Change = K74_Daily_Change['Real PNL'].sum()


Corp_L81_Cost_Summary_Recent = L81_Summary_Recent['Cost'].sum()
Corp_L81_Market_Value_Summary_Recent = L81_Summary_Recent['Market Value'].sum()
Corp_L81_Requirement_Summary_Recent = L81_Summary_Recent['Requirement'].sum()
Corp_L81_Unreal_PNL_Summary_Recent = L81_Summary_Recent['Unreal PNL'].sum()
Corp_L81_Real_PNL_Summary_Recent = L81_Summary_Recent['Real PNL'].sum()

Corp_L81_Cost_Daily_Change = L81_Daily_Change['Cost'].sum()
Corp_L81_Market_Value_Daily_Change = L81_Daily_Change['Market Value'].sum()
Corp_L81_Requirement_Daily_Change = L81_Daily_Change['Requirement'].sum()
Corp_L81_Unreal_PNL_Daily_Change = L81_Daily_Change['Unreal PNL'].sum()
Corp_L81_Real_PNL_Daily_Change = L81_Daily_Change['Real PNL'].sum()

Corp_P03_Cost_Summary_Recent = P03_Summary_Recent['Cost'].sum()
Corp_P03_Market_Value_Summary_Recent = P03_Summary_Recent['Market Value'].sum()
Corp_P03_Requirement_Summary_Recent = P03_Summary_Recent['Requirement'].sum()
Corp_P03_Unreal_PNL_Summary_Recent = P03_Summary_Recent['Unreal PNL'].sum()
Corp_P03_Real_PNL_Summary_Recent = P03_Summary_Recent['Real PNL'].sum()

Corp_P03_Cost_Daily_Change = P03_Daily_Change['Cost'].sum()
Corp_P03_Market_Value_Daily_Change = P03_Daily_Change['Market Value'].sum()
Corp_P03_Requirement_Daily_Change = P03_Daily_Change['Requirement'].sum()
Corp_P03_Unreal_PNL_Daily_Change = P03_Daily_Change['Unreal PNL'].sum()
Corp_P03_Real_PNL_Daily_Change = P03_Daily_Change['Real PNL'].sum()

Corp_N87_Cost_Summary_Recent = N87_Summary_Recent['Cost'].sum()
Corp_N87_Market_Value_Summary_Recent = N87_Summary_Recent['Market Value'].sum()
Corp_N87_Requirement_Summary_Recent = N87_Summary_Recent['Requirement'].sum()
Corp_N87_Unreal_PNL_Summary_Recent =N87_Summary_Recent['Unreal PNL'].sum()
Corp_N87_Real_PNL_Summary_Recent = N87_Summary_Recent['Real PNL'].sum()

Corp_N87_Cost_Daily_Change =N87_Daily_Change['Cost'].sum()
Corp_N87_Market_Value_Daily_Change = N87_Daily_Change['Market Value'].sum()
Corp_N87_Requirement_Daily_Change = N87_Daily_Change['Requirement'].sum()
Corp_N87_Unreal_PNL_Daily_Change = N87_Daily_Change['Unreal PNL'].sum()
Corp_N87_Real_PNL_Daily_Change = N87_Daily_Change['Real PNL'].sum()


# Overall Totals
Corp_Total_Cost_Summary = (Corp_N88_Cost_Summary_Recent,
                              Corp_N90_Cost_Summary_Recent,
                              Corp_P01_Cost_Summary_Recent,
                              Corp_P02_Cost_Summary_Recent,
                              Corp_K74_Cost_Summary_Recent,
                              Corp_L81_Cost_Summary_Recent,
                              Corp_P03_Cost_Summary_Recent,
                              Corp_N87_Cost_Summary_Recent)
Corp_Total_Cost_Summary = sum(Corp_Total_Cost_Summary)


Corp_Total_Cost_Daily = (Corp_N88_Cost_Daily_Change,
                              Corp_N90_Cost_Daily_Change,
                              Corp_P01_Cost_Daily_Change,
                              Corp_P02_Cost_Daily_Change,
                              Corp_K74_Cost_Daily_Change,
                              Corp_L81_Cost_Daily_Change,
                        Corp_P03_Cost_Daily_Change,
                        Corp_N87_Cost_Daily_Change)
Corp_Total_Cost_Daily = sum(Corp_Total_Cost_Daily)

Corp_Total_Market_Value_Summary = (Corp_N88_Market_Value_Summary_Recent,
                                      Corp_N90_Market_Value_Summary_Recent,
                                      Corp_P01_Market_Value_Summary_Recent,
                                      Corp_P02_Market_Value_Summary_Recent,
                                      Corp_K74_Market_Value_Summary_Recent,
                                      Corp_L81_Market_Value_Summary_Recent,
                                  Corp_P03_Market_Value_Summary_Recent,
                                  Corp_N87_Market_Value_Summary_Recent)
Corp_Total_Market_Value_Summary = sum(Corp_Total_Market_Value_Summary)

Corp_Total_Market_Value_Daily = (Corp_N88_Market_Value_Daily_Change,
                                    Corp_N90_Market_Value_Daily_Change,
                                    Corp_P01_Market_Value_Daily_Change,
                                    Corp_P02_Market_Value_Daily_Change,
                                    Corp_K74_Market_Value_Daily_Change,
                                    Corp_L81_Market_Value_Daily_Change,
                                Corp_P03_Market_Value_Daily_Change,
                                Corp_N87_Market_Value_Daily_Change)
Corp_Total_Market_Value_Daily = sum(Corp_Total_Market_Value_Daily)

Corp_Total_Requirement_Summary = (Corp_N88_Requirement_Summary_Recent,
                              Corp_N90_Requirement_Summary_Recent,
                              Corp_P01_Requirement_Summary_Recent,
                              Corp_P02_Requirement_Summary_Recent,
                              Corp_K74_Requirement_Summary_Recent,
                              Corp_L81_Requirement_Summary_Recent,
                                 Corp_P03_Requirement_Summary_Recent,
                                 Corp_N87_Requirement_Summary_Recent)
Corp_Total_Requirement_Summary = sum(Corp_Total_Requirement_Summary)

Corp_Total_Requirement_Daily = (Corp_N88_Requirement_Daily_Change,
                              Corp_N90_Requirement_Daily_Change,
                              Corp_P01_Requirement_Daily_Change,
                              Corp_P02_Requirement_Daily_Change,
                              Corp_K74_Requirement_Daily_Change,
                              Corp_L81_Requirement_Daily_Change,
                              Corp_P03_Requirement_Daily_Change,
                              Corp_N87_Requirement_Daily_Change)
Corp_Total_Requirement_Daily = sum(Corp_Total_Requirement_Daily)

Corp_Total_Unreal_PNL_Summary = (Corp_N88_Unreal_PNL_Summary_Recent,
                              Corp_N90_Unreal_PNL_Summary_Recent,
                              Corp_P01_Unreal_PNL_Summary_Recent,
                              Corp_P02_Unreal_PNL_Summary_Recent,
                              Corp_K74_Unreal_PNL_Summary_Recent,
                              Corp_L81_Unreal_PNL_Summary_Recent,
                              Corp_P03_Unreal_PNL_Summary_Recent,
                              Corp_N87_Unreal_PNL_Summary_Recent)
Corp_Total_Unreal_PNL_Summary = sum(Corp_Total_Unreal_PNL_Summary)

Corp_Total_Unreal_PNL_Daily = (Corp_N88_Unreal_PNL_Daily_Change,
                              Corp_N90_Unreal_PNL_Daily_Change,
                              Corp_P01_Unreal_PNL_Daily_Change,
                              Corp_P02_Unreal_PNL_Daily_Change,
                              Corp_K74_Unreal_PNL_Daily_Change,
                              Corp_L81_Unreal_PNL_Daily_Change,
                              Corp_P03_Unreal_PNL_Daily_Change,
                              Corp_N87_Unreal_PNL_Daily_Change)
Corp_Total_Unreal_PNL_Daily = sum(Corp_Total_Unreal_PNL_Daily)


Corp_Total_Real_PNL_Summary = (Corp_N88_Real_PNL_Summary_Recent,
                              Corp_N90_Real_PNL_Summary_Recent,
                              Corp_P01_Real_PNL_Summary_Recent,
                              Corp_P02_Real_PNL_Summary_Recent,
                              Corp_K74_Real_PNL_Summary_Recent,
                              Corp_L81_Real_PNL_Summary_Recent,
                              Corp_P03_Real_PNL_Summary_Recent,
                              Corp_N87_Real_PNL_Summary_Recent)
Corp_Total_Real_PNL_Summary = sum(Corp_Total_Real_PNL_Summary )

Corp_Total_Real_PNL_Daily = (Corp_N88_Real_PNL_Daily_Change,
                              Corp_N90_Real_PNL_Daily_Change,
                              Corp_P01_Real_PNL_Daily_Change,
                              Corp_P02_Real_PNL_Daily_Change,
                              Corp_K74_Real_PNL_Daily_Change,
                              Corp_L81_Real_PNL_Daily_Change,
                              Corp_P03_Real_PNL_Daily_Change,
                              Corp_N87_Real_PNL_Daily_Change)
Corp_Total_Real_PNL_Daily = sum(Corp_Total_Real_PNL_Daily)


# CD

CD_Cost_Summary_Recent = CD_Summary_Recent['Cost'].sum()
CD_Market_Value_Summary_Recent = CD_Summary_Recent['Market Value'].sum()
CD_Requirement_Summary_Recent = CD_Summary_Recent['Requirement'].sum()
CD_Unreal_PNL_Summary_Recent = CD_Summary_Recent['Unreal PNL'].sum()
CD_Real_PNL_Summary_Recent = CD_Summary_Recent['Real PNL'].sum()

CD_Cost_Daily_Change = CD_Daily_Change['Cost'].sum()
CD_Market_Value_Daily_Change = CD_Daily_Change['Market Value'].sum()
CD_Requirement_Daily_Change = CD_Daily_Change['Requirement'].sum()
CD_Unreal_PNL_Daily_Change = CD_Daily_Change['Unreal PNL'].sum()
CD_Real_PNL_Daily_Change = CD_Daily_Change['Real PNL'].sum()


#CMO
CMO_Cost_Summary_Recent = CMO_Summary_Recent['Cost'].sum()
CMO_Market_Value_Summary_Recent = CMO_Summary_Recent['Market Value'].sum()
CMO_Requirement_Summary_Recent = CMO_Summary_Recent['Requirement'].sum()
CMO_Unreal_PNL_Summary_Recent = CMO_Summary_Recent['Unreal PNL'].sum()
CMO_Real_PNL_Summary_Recent = CMO_Summary_Recent['Real PNL'].sum()

CMO_Cost_Daily_Change = CMO_Daily_Change['Cost'].sum()
CMO_Market_Value_Daily_Change = CMO_Daily_Change['Market Value'].sum()
CMO_Requirement_Daily_Change = CMO_Daily_Change['Requirement'].sum()
CMO_Unreal_PNL_Daily_Change = CMO_Daily_Change['Unreal PNL'].sum()
CMO_Real_PNL_Daily_Change = CMO_Daily_Change['Real PNL'].sum()




Firm_Cost_Summary_Total = (Muni_Cost_Summary_Recent,Corp_Total_Cost_Summary,CD_Cost_Summary_Recent,CMO_Cost_Summary_Recent)
Firm_Market_Value_Summary_Total = (Muni_Market_Value_Summary_Recent,Corp_Total_Market_Value_Summary,CD_Market_Value_Summary_Recent,CMO_Market_Value_Summary_Recent )
Firm_Requirement_Summary_Total = (Muni_Requirement_Summary_Recent,Corp_Total_Requirement_Summary,CD_Requirement_Summary_Recent,CMO_Requirement_Summary_Recent)
Firm_Unreal_PNL_Summary_Total = (Muni_Unreal_PNL_Summary_Recent,Corp_Total_Unreal_PNL_Summary,CD_Unreal_PNL_Summary_Recent,CMO_Unreal_PNL_Summary_Recent)
Firm_Real_PNL_Summary_Total = (Muni_Real_PNL_Summary_Recent,Corp_Total_Real_PNL_Summary,CD_Real_PNL_Summary_Recent,CMO_Real_PNL_Summary_Recent)

Firm_Cost_Daily_Total = (Muni_Cost_Daily_Change,Corp_Total_Cost_Daily,CD_Cost_Daily_Change,CMO_Cost_Daily_Change)
Firm_Market_Value_Daily_Total = (Muni_Market_Value_Daily_Change,Corp_Total_Market_Value_Daily,CD_Market_Value_Daily_Change,CMO_Market_Value_Daily_Change)
Firm_Requirement_Daily_Total = (Muni_Requirement_Daily_Change,Corp_Total_Requirement_Daily,CD_Requirement_Daily_Change,CMO_Requirement_Daily_Change)
Firm_Unreal_PNL_Daily_Total = (Muni_Unreal_PNL_Daily_Change,Corp_Total_Unreal_PNL_Daily,CD_Unreal_PNL_Daily_Change,CMO_Unreal_PNL_Daily_Change)
Firm_Real_PNL_Daily_Total = (Muni_Real_PNL_Daily_Change,Corp_Total_Real_PNL_Daily,CD_Real_PNL_Daily_Change,CMO_Real_PNL_Daily_Change)



Firm_Cost_Summary_Total = sum(Firm_Cost_Summary_Total)
Firm_Market_Value_Summary_Total = sum(Firm_Market_Value_Summary_Total)
Firm_Requirement_Summary_Total = sum(Firm_Requirement_Summary_Total)
Firm_Unreal_PNL_Summary_Total = sum(Firm_Unreal_PNL_Summary_Total)
Firm_Real_PNL_Summary_Total = sum(Firm_Real_PNL_Summary_Total)
Firm_Cost_Daily_Total = sum(Firm_Cost_Daily_Total)
Firm_Market_Value_Daily_Total = sum(Firm_Market_Value_Daily_Total)
Firm_Requirement_Daily_Total = sum(Firm_Requirement_Daily_Total)
Firm_Unreal_PNL_Daily_Total = sum(Firm_Unreal_PNL_Daily_Total)
Firm_Real_PNL_Daily_Total = sum(Firm_Real_PNL_Daily_Total)







# Hilltop_Summary_new.to_excel(writer,sheet_name ='Summary',index=True,startrow=1)

Hilltop_Recent_s['Change'] = Hilltop_Recent_s['Total Available Funds']-Hilltop_Old_s['Total Available Funds']
Hilltop_Recent_s = Hilltop_Recent_s[3:]
Hilltop_Recent_s = Hilltop_Recent_s[['Account Number','Total Available Funds','Change']]
Hilltop_Recent_s.to_excel(writer,sheet_name ='Summary',index=False,startrow=1,startcol=30)
workbook = writer.book


format_mini_total = workbook.add_format({'num_format': '#,##0',
                                         'font_size':'8',
                                         'bold': True,
                                         'top':1})
format_general = workbook.add_format({'num_format': '#,##0',
                                         'font_size':'8'})
format_blank_blue = workbook.add_format({'bg_color':'#4267b8',
                                         'font_size':'8',
                                         'font_color':'white'})
format_top_summary = workbook.add_format({'bg_color':'#000e6b',
                                          'font_size':'10',
                                          'font_color':'white'})

format_grey_columnhead = workbook.add_format({'bg_color':'#d4d4d4',
                                              'font_size':'8',
                                              'bottom':1})

format_subtotal = workbook.add_format({'num_format': '#,##0',
                                       'bold':True,
                                       'font_size':'10',
                                       'bottom':2,
                                       'top':1})
format_general_row = workbook.add_format({'font_size':'8',
                                          'num_format': '#,##0'})
format_general_row_green = workbook.add_format({'font_size':'8',
                                                'num_format': '#,##0',
                                                'font_color':'green'})
format_general_row_red = workbook.add_format({'font_size':'8',
                                              'num_format': '#,##0',
                                              'font_color':'red'})
format_subtotal_row = workbook.add_format({'font_size':'10',
                                          'num_format': '#,##0',
                                          'bottom':1,
                                          'top':1})
format_group_total = workbook.add_format({'font_size':'10',
                                          'num_format':'#,##0',
                                          'bold': True})
format_column = workbook.add_format({'bottom':0,
                                     'top':0,
                                     'border_color':'white'})
format_url_links = workbook.add_format({'font_size':'10',
                                       'font_color':'blue',
                                       'underline': 1})

merge_format = workbook.add_format({
    'font_size':'10',
    'bold': 1,
    'border': 0,
    'align': 'left',
    'valign': 'vcenter',
    'fg_color':'#000e6b',
    'font_color':'white'})

worksheet_summary = writer.sheets['Summary']

worksheet_summary.write('D1', 'Summary',format_top_summary) 
worksheet_summary.write('E1', ' ',format_top_summary)
worksheet_summary.write('F1', ' ',format_top_summary)
worksheet_summary.write('D2', 'Item',format_grey_columnhead)
worksheet_summary.write('E2', 'Available Funds',format_grey_columnhead)
worksheet_summary.write('F2', 'Change',format_grey_columnhead)

worksheet_summary.write_formula('D3','=AE3',format_general)
worksheet_summary.write_formula('D4','=AE4',format_general)
worksheet_summary.write_formula('D5','=AE5',format_mini_total)
worksheet_summary.write_formula('D6','=AE6',format_general)
worksheet_summary.write_formula('D7','=AE7',format_general)
worksheet_summary.write_formula('D8','=AE8',format_general)
worksheet_summary.write_formula('D9','=AE9',format_mini_total)
worksheet_summary.write_formula('E3','=AF3',format_general)
worksheet_summary.write_formula('E4','=AF4',format_general)
worksheet_summary.write_formula('E5','=AF5',format_mini_total)
worksheet_summary.write_formula('E6','=AF6',format_general)
worksheet_summary.write_formula('E7','=AF7',format_general)
worksheet_summary.write_formula('E8','=AF8',format_general)
worksheet_summary.write_formula('E9','=AF9',format_mini_total)
worksheet_summary.write_formula('F3','=AG3',format_general)
worksheet_summary.write_formula('F4','=AG4',format_general)
worksheet_summary.write_formula('F5','=AG5',format_mini_total)
worksheet_summary.write_formula('F6','=AG6',format_general)
worksheet_summary.write_formula('F7','=AG7',format_general)
worksheet_summary.write_formula('F8','=AG8',format_general)
worksheet_summary.write_formula('F9','=AG9',format_mini_total)

worksheet_summary.write('A19', 'Muni Total',format_subtotal)
worksheet_summary.write('B19', ' ',format_subtotal)
worksheet_summary.write('I19', 'Muni Total',format_subtotal)
worksheet_summary.write('J19', ' ',format_subtotal)
worksheet_summary.write('C19', Muni_Cost_Summary_Recent,format_subtotal)
worksheet_summary.write('D19', Muni_Market_Value_Summary_Recent,format_subtotal)
worksheet_summary.write('E19', Muni_Requirement_Summary_Recent,format_subtotal)
worksheet_summary.write('F19', Muni_Unreal_PNL_Summary_Recent,format_subtotal)
worksheet_summary.write('G19', Muni_Real_PNL_Summary_Recent,format_subtotal)
worksheet_summary.write('K19', Muni_Cost_Daily_Change,format_subtotal)
worksheet_summary.write('L19', Muni_Market_Value_Daily_Change,format_subtotal)
worksheet_summary.write('M19', Muni_Requirement_Daily_Change,format_subtotal)
worksheet_summary.write('N19', Muni_Unreal_PNL_Daily_Change,format_subtotal)
worksheet_summary.write('O19', Muni_Real_PNL_Daily_Change,format_subtotal)

worksheet_summary.write('A12', 'Account Name',format_grey_columnhead)
worksheet_summary.write('B12', 'Position Type',format_grey_columnhead)
worksheet_summary.write('C12', 'Cost',format_grey_columnhead)
worksheet_summary.write('D12', 'Market Value',format_grey_columnhead)
worksheet_summary.write('E12', 'Requirement',format_grey_columnhead)
worksheet_summary.write('F12', 'Unreal PNL',format_grey_columnhead)
worksheet_summary.write('G12', 'Real PNL',format_grey_columnhead)
worksheet_summary.write('I12', 'Account Name',format_grey_columnhead)
worksheet_summary.write('J12', 'Position Type',format_grey_columnhead)
worksheet_summary.write('K12', 'Cost',format_grey_columnhead)
worksheet_summary.write('L12', 'Market Value',format_grey_columnhead)
worksheet_summary.write('M12', 'Requirement',format_grey_columnhead)
worksheet_summary.write('N12', 'UnrealPNL',format_grey_columnhead)
worksheet_summary.write('O12', 'Real PNL',format_grey_columnhead)

if 'K72 Muni Inv Fl' in Muni_Summary_Recent_Short:
    worksheet_summary.write('B13', '*short*')
else:
    worksheet_summary.write('B13', '     -')
if 'K78 Taxable Mun' in Muni_Summary_Recent_Short:
    worksheet_summary.write('B14', '*short*')
else:
    worksheet_summary.write('B14', '     -')
if 'K79 Cali Tax Ex' in Muni_Summary_Recent_Short:
    worksheet_summary.write('B15', '*short*')
else:
    worksheet_summary.write('B15', '     -')
if 'K80 Muni Tax Ex' in Muni_Summary_Recent_Short:
    worksheet_summary.write('B16', '*short*')
else:
    worksheet_summary.write('B16', '     -')
if 'K81 Muni Tax' in Muni_Summary_Recent_Short:
    worksheet_summary.write('B17', '*short*')
else:
    worksheet_summary.write('B17', '     -')
if 'K82 Tax 0 Muni' in Muni_Summary_Recent_Short:
    worksheet_summary.write('B18', '*short*')
else:
    worksheet_summary.write('B18', '     -')


if 'K72 Muni Inv Fl' in Muni_Daily_Change_Short:
    worksheet_summary.write('J13', '*short*')
else:
    worksheet_summary.write('J13', '     -')
if 'K78 Taxable Mun' in Muni_Daily_Change_Short:
    worksheet_summary.write('J14', '*short*')
else:
    worksheet_summary.write('J14', '     -')
if 'K79 Cali Tax Ex' in Muni_Daily_Change_Short:
    worksheet_summary.write('J15', '*short*')
else:
    worksheet_summary.write('J15', '     -')
if 'K80 Muni Tax Ex' in Muni_Daily_Change_Short:
    worksheet_summary.write('J16', '*short*')
else:
    worksheet_summary.write('J16', '     -')
if 'K81 Muni Tax' in Muni_Daily_Change_Short:
    worksheet_summary.write('J17', '*short*')
else:
    worksheet_summary.write('J17', '     -')
if 'K82 Tax 0 Muni' in Muni_Daily_Change_Short:
    worksheet_summary.write('J18', '*short*')
else:
    worksheet_summary.write('J18', '     -')



worksheet_summary.write('A20', ' ')
worksheet_summary.write('B20', ' ')
worksheet_summary.write('C20', ' ')
worksheet_summary.write('D20', ' ')
worksheet_summary.write('E20', ' ')
worksheet_summary.write('F20', ' ')
worksheet_summary.write('G20', ' ')
worksheet_summary.write('I20', ' ')
worksheet_summary.write('J20', ' ')
worksheet_summary.write('K20', ' ')
worksheet_summary.write('L20', ' ')
worksheet_summary.write('M20', ' ')
worksheet_summary.write('N20', ' ')
worksheet_summary.write('O20', ' ')

worksheet_summary.write('A13', 'K72 MUNI',format_general)
worksheet_summary.write('A14', 'K78 MUNTAX',format_general)
worksheet_summary.write('A15', 'K79 MUNCC',format_general)
worksheet_summary.write('A16', 'K80 MUNBT',format_general)
worksheet_summary.write('A17', 'K81 MUNCCTAX',format_general)
worksheet_summary.write('A18', 'K82 MUNBTTAX',format_general)
worksheet_summary.write('A21', 'N88 CORPIG',format_general)
worksheet_summary.write('A25', 'N90 CORPNOTE',format_general)
worksheet_summary.write('A29', 'P01 CORPFRN',format_general)
worksheet_summary.write('A33', 'P02 CORPSP',format_general)
worksheet_summary.write('A37', 'K74 CORPHY',format_general)
worksheet_summary.write('A41', 'L81 Corp Other',format_general)
worksheet_summary.write('A45', 'P03 CORPDIST',format_general)
worksheet_summary.write('A49', 'N87 CORPXOVR',format_general)
worksheet_summary.write('A56', 'K76 CMO',format_general)
worksheet_summary.write('A57', 'M64 IO',format_general)

worksheet_summary.write('I13', 'K72 MUNI',format_general)
worksheet_summary.write('I14', 'K78 MUNTAX',format_general)
worksheet_summary.write('I15', 'K79 MUNCC',format_general)
worksheet_summary.write('I16', 'K80 MUNBT',format_general)
worksheet_summary.write('I17', 'K81 MUNCCTAX',format_general)
worksheet_summary.write('I18', 'K82 MUNBTTAX',format_general)
worksheet_summary.write('I21', 'N88 CORPIG',format_general)
worksheet_summary.write('I25', 'N90 CORPNOTE',format_general)
worksheet_summary.write('I29', 'P01 CORPFRN',format_general)
worksheet_summary.write('I33', 'P02 CORPSP',format_general)
worksheet_summary.write('I37', 'K74 CORPHY',format_general)
worksheet_summary.write('I41', 'L81 Corp Other',format_general)
worksheet_summary.write('I45', 'P03 CORPDIST',format_general)
worksheet_summary.write('I49', 'N87 CORPXOVR',format_general)
worksheet_summary.write('I56', 'K76 CMO',format_general)
worksheet_summary.write('I57', 'M64 IO',format_general)

worksheet_summary.write('A22', ' ')
worksheet_summary.write('I22', ' ')
worksheet_summary.write('B23', 'Total',format_mini_total)
worksheet_summary.write('J23', 'Total',format_mini_total)
worksheet_summary.write('C23', Corp_N88_Cost_Summary_Recent,format_mini_total)
worksheet_summary.write('D23', Corp_N88_Market_Value_Summary_Recent,format_mini_total)
worksheet_summary.write('E23', Corp_N88_Requirement_Summary_Recent,format_mini_total)
worksheet_summary.write('F23', Corp_N88_Unreal_PNL_Summary_Recent,format_mini_total)
worksheet_summary.write('G23', Corp_N88_Real_PNL_Summary_Recent,format_mini_total)
worksheet_summary.write('I23', ' ')
worksheet_summary.write('K23', Corp_N88_Cost_Daily_Change,format_mini_total)
worksheet_summary.write('L23', Corp_N88_Market_Value_Daily_Change,format_mini_total)
worksheet_summary.write('M23', Corp_N88_Requirement_Daily_Change ,format_mini_total)
worksheet_summary.write('N23', Corp_N88_Unreal_PNL_Daily_Change,format_mini_total)
worksheet_summary.write('O23', Corp_N88_Real_PNL_Daily_Change,format_mini_total)


worksheet_summary.write('A24', ' ')
worksheet_summary.write('B24', ' ')
worksheet_summary.write('C24', ' ')
worksheet_summary.write('D24', ' ')
worksheet_summary.write('E24', ' ')
worksheet_summary.write('F24', ' ')
worksheet_summary.write('G24', ' ')
worksheet_summary.write('I24', ' ')
worksheet_summary.write('J24', ' ')
worksheet_summary.write('K24', ' ')
worksheet_summary.write('L24', ' ')
worksheet_summary.write('M24', ' ')
worksheet_summary.write('N24', ' ')
worksheet_summary.write('O24', ' ')

worksheet_summary.write('I26', ' ')
worksheet_summary.write('A26', ' ')
worksheet_summary.write('B27', 'Total',format_mini_total)
worksheet_summary.write('J27', 'Total',format_mini_total)
worksheet_summary.write('C27', Corp_N90_Cost_Summary_Recent,format_mini_total)
worksheet_summary.write('D27', Corp_N90_Market_Value_Summary_Recent,format_mini_total)
worksheet_summary.write('E27', Corp_N90_Requirement_Summary_Recent,format_mini_total)
worksheet_summary.write('F27', Corp_N90_Unreal_PNL_Summary_Recent,format_mini_total)
worksheet_summary.write('G27', Corp_N90_Real_PNL_Summary_Recent,format_mini_total)
worksheet_summary.write('I27', ' ')
worksheet_summary.write('K27', Corp_N90_Cost_Daily_Change,format_mini_total)
worksheet_summary.write('L27', Corp_N90_Market_Value_Daily_Change,format_mini_total)
worksheet_summary.write('M27', Corp_N90_Requirement_Daily_Change,format_mini_total)
worksheet_summary.write('N27', Corp_N90_Unreal_PNL_Daily_Change,format_mini_total)
worksheet_summary.write('O27', Corp_N90_Real_PNL_Daily_Change,format_mini_total)

worksheet_summary.write('A28', ' ')
worksheet_summary.write('B28', ' ')
worksheet_summary.write('C28', ' ')
worksheet_summary.write('D28', ' ')
worksheet_summary.write('E28', ' ')
worksheet_summary.write('F28', ' ')
worksheet_summary.write('G28', ' ')
worksheet_summary.write('I28', ' ')
worksheet_summary.write('J28', ' ')
worksheet_summary.write('K28', ' ')
worksheet_summary.write('L28', ' ')
worksheet_summary.write('M28', ' ')
worksheet_summary.write('N28', ' ')
worksheet_summary.write('O28', ' ')

worksheet_summary.write('A30', ' ')
worksheet_summary.write('I30', ' ')
worksheet_summary.write('B31', 'Total',format_mini_total)
worksheet_summary.write('J31', 'Total',format_mini_total)
worksheet_summary.write('C31', Corp_P01_Cost_Summary_Recent ,format_mini_total)
worksheet_summary.write('D31', Corp_P01_Market_Value_Summary_Recent,format_mini_total)
worksheet_summary.write('E31', Corp_P01_Requirement_Summary_Recent,format_mini_total)
worksheet_summary.write('F31', Corp_P01_Unreal_PNL_Summary_Recent,format_mini_total)
worksheet_summary.write('G31', Corp_P01_Real_PNL_Summary_Recent,format_mini_total)
worksheet_summary.write('I31', ' ')
worksheet_summary.write('K31', Corp_P01_Cost_Daily_Change,format_mini_total)
worksheet_summary.write('L31', Corp_P01_Market_Value_Daily_Change,format_mini_total)
worksheet_summary.write('M31', Corp_P01_Requirement_Daily_Change,format_mini_total)
worksheet_summary.write('N31', Corp_P01_Unreal_PNL_Daily_Change,format_mini_total)
worksheet_summary.write('O31', Corp_P01_Real_PNL_Daily_Change,format_mini_total)

worksheet_summary.write('A32', ' ')
worksheet_summary.write('B32', ' ')
worksheet_summary.write('C32', ' ')
worksheet_summary.write('D32', ' ')
worksheet_summary.write('E32', ' ')
worksheet_summary.write('F32', ' ')
worksheet_summary.write('G32', ' ')
worksheet_summary.write('I32', ' ')
worksheet_summary.write('J32', ' ')
worksheet_summary.write('K32', ' ')
worksheet_summary.write('L32', ' ')
worksheet_summary.write('M32', ' ')
worksheet_summary.write('N32', ' ')
worksheet_summary.write('O32', ' ')


worksheet_summary.write('A34', ' ')
worksheet_summary.write('I34', ' ')
worksheet_summary.write('B35', 'Total',format_mini_total)
worksheet_summary.write('J35', 'Total',format_mini_total)
worksheet_summary.write('C35', Corp_P02_Cost_Summary_Recent ,format_mini_total)
worksheet_summary.write('D35', Corp_P02_Market_Value_Summary_Recent,format_mini_total)
worksheet_summary.write('E35', Corp_P02_Requirement_Summary_Recent,format_mini_total)
worksheet_summary.write('F35', Corp_P02_Unreal_PNL_Summary_Recent,format_mini_total)
worksheet_summary.write('G35', Corp_P02_Real_PNL_Summary_Recent,format_mini_total)
worksheet_summary.write('I35', ' ')
worksheet_summary.write('K35', Corp_P02_Cost_Daily_Change,format_mini_total)
worksheet_summary.write('L35', Corp_P02_Market_Value_Daily_Change,format_mini_total)
worksheet_summary.write('M35', Corp_P02_Requirement_Daily_Change,format_mini_total)
worksheet_summary.write('N35', Corp_P02_Unreal_PNL_Daily_Change,format_mini_total)
worksheet_summary.write('O35', Corp_P02_Real_PNL_Daily_Change,format_mini_total)


Corp_P02_Cost_Summary_Recent = P02_Summary_Recent['Cost'].sum()
Corp_P02_Market_Value_Summary_Recent = P02_Summary_Recent['Market Value'].sum()
Corp_P02_Requirement_Summary_Recent = P02_Summary_Recent['Requirement'].sum()
Corp_P02_Unreal_PNL_Summary_Recent = P02_Summary_Recent['Unreal PNL'].sum()
Corp_P02_Real_PNL_Summary_Recent = P02_Summary_Recent['Real PNL'].sum()

Corp_P02_Cost_Daily_Change = P02_Daily_Change['Cost'].sum()
Corp_P02_Market_Value_Daily_Change = P02_Daily_Change['Market Value'].sum()
Corp_P02_Requirement_Daily_Change = P02_Daily_Change['Requirement'].sum()
Corp_P02_Unreal_PNL_Daily_Change = P02_Daily_Change['Unreal PNL'].sum()
Corp_P02_Real_PNL_Daily_Change = P02_Daily_Change['Real PNL'].sum()


worksheet_summary.write('A36', ' ')
worksheet_summary.write('B36', ' ')
worksheet_summary.write('C36', ' ')
worksheet_summary.write('D36', ' ')
worksheet_summary.write('E36', ' ')
worksheet_summary.write('F36', ' ')
worksheet_summary.write('G36', ' ')
worksheet_summary.write('I36', ' ')
worksheet_summary.write('J36', ' ')
worksheet_summary.write('K36', ' ')
worksheet_summary.write('L36', ' ')
worksheet_summary.write('M36', ' ')
worksheet_summary.write('N36', ' ')
worksheet_summary.write('O36', ' ')

worksheet_summary.write('A38', ' ')
worksheet_summary.write('I38', ' ')
worksheet_summary.write('B39', 'Total',format_mini_total)
worksheet_summary.write('J39', 'Total',format_mini_total)
worksheet_summary.write('C39', Corp_K74_Cost_Summary_Recent ,format_mini_total)
worksheet_summary.write('D39', Corp_K74_Market_Value_Summary_Recent,format_mini_total)
worksheet_summary.write('E39', Corp_K74_Requirement_Summary_Recent,format_mini_total)
worksheet_summary.write('F39', Corp_K74_Unreal_PNL_Summary_Recent,format_mini_total)
worksheet_summary.write('G39', Corp_K74_Real_PNL_Summary_Recent,format_mini_total)
worksheet_summary.write('I39', ' ')
worksheet_summary.write('K39', Corp_K74_Cost_Daily_Change,format_mini_total)
worksheet_summary.write('L39', Corp_K74_Market_Value_Daily_Change,format_mini_total)
worksheet_summary.write('M39', Corp_K74_Requirement_Daily_Change,format_mini_total)
worksheet_summary.write('N39', Corp_K74_Unreal_PNL_Daily_Change,format_mini_total)
worksheet_summary.write('O39', Corp_K74_Real_PNL_Daily_Change,format_mini_total)

worksheet_summary.write('A42', ' ')
worksheet_summary.write('I42', ' ')
worksheet_summary.write('B43', 'Total',format_mini_total)
worksheet_summary.write('J43', 'Total',format_mini_total)
worksheet_summary.write('C43', Corp_L81_Cost_Summary_Recent ,format_mini_total)
worksheet_summary.write('D43', Corp_L81_Market_Value_Summary_Recent,format_mini_total)
worksheet_summary.write('E43', Corp_L81_Requirement_Summary_Recent,format_mini_total)
worksheet_summary.write('F43', Corp_L81_Unreal_PNL_Summary_Recent,format_mini_total)
worksheet_summary.write('G43', Corp_L81_Real_PNL_Summary_Recent,format_mini_total)
worksheet_summary.write('I43', ' ')
worksheet_summary.write('K43', Corp_L81_Cost_Daily_Change,format_mini_total)
worksheet_summary.write('L43', Corp_L81_Market_Value_Daily_Change,format_mini_total)
worksheet_summary.write('M43', Corp_L81_Requirement_Daily_Change,format_mini_total)
worksheet_summary.write('N43', Corp_L81_Unreal_PNL_Daily_Change,format_mini_total)
worksheet_summary.write('O43', Corp_L81_Real_PNL_Daily_Change,format_mini_total)

worksheet_summary.write('A44', ' ')
worksheet_summary.write('B44', ' ')
worksheet_summary.write('C44', ' ')
worksheet_summary.write('D44', ' ')
worksheet_summary.write('E44', ' ')
worksheet_summary.write('F44', ' ')
worksheet_summary.write('G44', ' ')
worksheet_summary.write('I44', ' ')
worksheet_summary.write('J44', ' ')
worksheet_summary.write('K44', ' ')
worksheet_summary.write('L44', ' ')
worksheet_summary.write('M44', ' ')
worksheet_summary.write('N44', ' ')
worksheet_summary.write('O44', ' ')


worksheet_summary.write('A48', ' ')
worksheet_summary.write('B48', ' ')
worksheet_summary.write('C48', ' ')
worksheet_summary.write('D48', ' ')
worksheet_summary.write('E48', ' ')
worksheet_summary.write('F48', ' ')
worksheet_summary.write('G48', ' ')
worksheet_summary.write('I48', ' ')
worksheet_summary.write('J48', ' ')
worksheet_summary.write('K48', ' ')
worksheet_summary.write('L48', ' ')
worksheet_summary.write('M48', ' ')
worksheet_summary.write('N48', ' ')
worksheet_summary.write('O48', ' ')

worksheet_summary.write('A50', ' ')
worksheet_summary.write('I50', ' ')
worksheet_summary.write('B51', 'Total',format_mini_total)
worksheet_summary.write('J51', 'Total',format_mini_total)
worksheet_summary.write('C51', Corp_N87_Cost_Summary_Recent ,format_mini_total)
worksheet_summary.write('D51', Corp_N87_Market_Value_Summary_Recent,format_mini_total)
worksheet_summary.write('E51', Corp_N87_Requirement_Summary_Recent,format_mini_total)
worksheet_summary.write('F51', Corp_N87_Unreal_PNL_Summary_Recent,format_mini_total)
worksheet_summary.write('G51', Corp_N87_Real_PNL_Summary_Recent,format_mini_total)
worksheet_summary.write('I51', ' ')
worksheet_summary.write('K51', Corp_N87_Cost_Daily_Change,format_mini_total)
worksheet_summary.write('L51', Corp_N87_Market_Value_Daily_Change,format_mini_total)
worksheet_summary.write('M51', Corp_N87_Requirement_Daily_Change,format_mini_total)
worksheet_summary.write('N51', Corp_N87_Unreal_PNL_Daily_Change,format_mini_total)
worksheet_summary.write('O51', Corp_N87_Real_PNL_Daily_Change,format_mini_total)


worksheet_summary.write('A40', ' ')
worksheet_summary.write('B40', ' ')
worksheet_summary.write('C40', ' ')
worksheet_summary.write('D40', ' ')
worksheet_summary.write('E40', ' ')
worksheet_summary.write('F40', ' ')
worksheet_summary.write('G40', ' ')
worksheet_summary.write('I40', ' ')
worksheet_summary.write('J40', ' ')
worksheet_summary.write('K40', ' ')
worksheet_summary.write('L40', ' ')
worksheet_summary.write('M40', ' ')
worksheet_summary.write('N40', ' ')
worksheet_summary.write('O40', ' ')

worksheet_summary.write('A46', ' ')
worksheet_summary.write('I46', ' ')
worksheet_summary.write('B47', 'Total',format_mini_total)
worksheet_summary.write('J47', 'Total',format_mini_total)
worksheet_summary.write('C47', Corp_P03_Cost_Summary_Recent ,format_mini_total)
worksheet_summary.write('D47', Corp_P03_Market_Value_Summary_Recent,format_mini_total)
worksheet_summary.write('E47', Corp_P03_Requirement_Summary_Recent,format_mini_total)
worksheet_summary.write('F47', Corp_P03_Unreal_PNL_Summary_Recent,format_mini_total)
worksheet_summary.write('G47', Corp_P03_Real_PNL_Summary_Recent,format_mini_total)
worksheet_summary.write('I47', ' ')
worksheet_summary.write('K47', Corp_P03_Cost_Daily_Change,format_mini_total)
worksheet_summary.write('L47', Corp_P03_Market_Value_Daily_Change,format_mini_total)
worksheet_summary.write('M47', Corp_P03_Requirement_Daily_Change,format_mini_total)
worksheet_summary.write('N47', Corp_P03_Unreal_PNL_Daily_Change,format_mini_total)
worksheet_summary.write('O47', Corp_P03_Real_PNL_Daily_Change,format_mini_total)


worksheet_summary.write('A52', 'Corp Total',format_subtotal)
worksheet_summary.write('B52', ' ',format_subtotal)
worksheet_summary.write('I52', 'Corp Total',format_subtotal)
worksheet_summary.write('J52', ' ',format_subtotal)
worksheet_summary.write('C52', Corp_Total_Cost_Summary,format_subtotal)
worksheet_summary.write('D52', Corp_Total_Market_Value_Summary,format_subtotal)
worksheet_summary.write('E52', Corp_Total_Requirement_Summary,format_subtotal)
worksheet_summary.write('F52', Corp_Total_Unreal_PNL_Summary,format_subtotal)
worksheet_summary.write('G52', Corp_Total_Real_PNL_Summary ,format_subtotal)
worksheet_summary.write('K52', Corp_Total_Cost_Daily,format_subtotal)
worksheet_summary.write('L52', Corp_Total_Market_Value_Daily ,format_subtotal)
worksheet_summary.write('M52', Corp_Total_Requirement_Daily,format_subtotal)
worksheet_summary.write('N52', Corp_Total_Unreal_PNL_Daily,format_subtotal)
worksheet_summary.write('O52', Corp_Total_Real_PNL_Daily,format_subtotal)

worksheet_summary.write('A53', ' ')
worksheet_summary.write('B53', ' ')
worksheet_summary.write('C53', ' ')
worksheet_summary.write('D53', ' ')
worksheet_summary.write('E53', ' ')
worksheet_summary.write('F53', ' ')
worksheet_summary.write('G53', ' ')
worksheet_summary.write('I53', ' ')
worksheet_summary.write('J53', ' ')
worksheet_summary.write('K53', ' ')
worksheet_summary.write('L53', ' ')
worksheet_summary.write('M53', ' ')
worksheet_summary.write('N53', ' ')
worksheet_summary.write('O53', ' ')


worksheet_summary.write('A54', 'CD Total',format_subtotal)
worksheet_summary.write('B54', ' ',format_subtotal)
worksheet_summary.write('I54', 'CD Total',format_subtotal)
worksheet_summary.write('J54', ' ',format_subtotal)
worksheet_summary.write('C54', CD_Cost_Summary_Recent ,format_subtotal)
worksheet_summary.write('D54', CD_Market_Value_Summary_Recent,format_subtotal)
worksheet_summary.write('E54', CD_Requirement_Summary_Recent,format_subtotal)
worksheet_summary.write('F54', CD_Unreal_PNL_Summary_Recent,format_subtotal)
worksheet_summary.write('G54', CD_Real_PNL_Summary_Recent,format_subtotal)
worksheet_summary.write('K54', CD_Cost_Daily_Change,format_subtotal)
worksheet_summary.write('L54', CD_Market_Value_Daily_Change,format_subtotal)
worksheet_summary.write('M54', CD_Requirement_Daily_Change ,format_subtotal)
worksheet_summary.write('N54', CD_Unreal_PNL_Daily_Change,format_subtotal)
worksheet_summary.write('O54', CD_Real_PNL_Daily_Change,format_subtotal)


worksheet_summary.write('A55', ' ')
worksheet_summary.write('B55', ' ')
worksheet_summary.write('C55', ' ')
worksheet_summary.write('D55', ' ')
worksheet_summary.write('E55', ' ')
worksheet_summary.write('F55', ' ')
worksheet_summary.write('G55', ' ')
worksheet_summary.write('I55', ' ')
worksheet_summary.write('J55', ' ')
worksheet_summary.write('K55', ' ')
worksheet_summary.write('L55', ' ')
worksheet_summary.write('M55', ' ')
worksheet_summary.write('N55', ' ')
worksheet_summary.write('O55', ' ')

worksheet_summary.write('B56', '     -')
worksheet_summary.write('B57', '     -')
worksheet_summary.write('J56', '     -')
worksheet_summary.write('J57', '     -')

worksheet_summary.write('A58', 'CMO Total',format_subtotal)
worksheet_summary.write('B58', ' ',format_subtotal)
worksheet_summary.write('I58', 'CMO Total',format_subtotal)
worksheet_summary.write('J58', ' ',format_subtotal)
worksheet_summary.write('C58', CMO_Cost_Summary_Recent,format_subtotal)
worksheet_summary.write('D58', CMO_Market_Value_Summary_Recent,format_subtotal)
worksheet_summary.write('E58', CMO_Requirement_Summary_Recent,format_subtotal)
worksheet_summary.write('F58', CMO_Unreal_PNL_Summary_Recent,format_subtotal)
worksheet_summary.write('G58', CMO_Real_PNL_Summary_Recent,format_subtotal)
worksheet_summary.write('K58', CMO_Cost_Daily_Change,format_subtotal)
worksheet_summary.write('L58', CMO_Market_Value_Daily_Change,format_subtotal)
worksheet_summary.write('M58', CMO_Requirement_Daily_Change,format_subtotal)
worksheet_summary.write('N58', CMO_Unreal_PNL_Daily_Change,format_subtotal)
worksheet_summary.write('O58', CMO_Real_PNL_Daily_Change,format_subtotal)


worksheet_summary.write('A60', 'Firm Total',format_subtotal)
worksheet_summary.write('B60', ' ',format_subtotal)
worksheet_summary.write('D60', Firm_Market_Value_Summary_Total,format_subtotal)
worksheet_summary.write('E60', Firm_Requirement_Summary_Total,format_subtotal)
worksheet_summary.write('F60', Firm_Unreal_PNL_Summary_Total,format_subtotal)
worksheet_summary.write('G60', Firm_Real_PNL_Summary_Total,format_subtotal)

worksheet_summary.write('I60', 'Firm Total',format_subtotal)
worksheet_summary.write('J60', ' ',format_subtotal)
worksheet_summary.write('L60', Firm_Market_Value_Daily_Total,format_subtotal)
worksheet_summary.write('M60', Firm_Requirement_Daily_Total,format_subtotal)
worksheet_summary.write('N60', Firm_Unreal_PNL_Daily_Total,format_subtotal)
worksheet_summary.write('O60', Firm_Real_PNL_Daily_Total,format_subtotal)


                                      
worksheet_summary.merge_range('A11:G11', 'Month Summary',merge_format)

worksheet_summary.merge_range('I11:O11', 'Daily Change',merge_format)
worksheet_summary.set_row(1,None,format_general_row) 
worksheet_summary.set_row(2,None,format_general_row) 
worksheet_summary.set_row(3,None,format_general_row) 
worksheet_summary.set_row(4,None,format_general_row) 
worksheet_summary.set_row(5,None,format_general_row) 
worksheet_summary.set_row(6,None,format_general_row) 
worksheet_summary.set_row(7,None,format_general_row) 
worksheet_summary.set_row(8,None,format_general_row) 

worksheet_summary.set_row(9,None,format_general_row)
worksheet_summary.set_row(10,None,format_general_row)
worksheet_summary.set_row(12,None,format_general_row)  
worksheet_summary.set_row(13,None,format_general_row)  
worksheet_summary.set_row(14,None,format_general_row)  
worksheet_summary.set_row(15,None,format_general_row) 
worksheet_summary.set_row(16,None,format_general_row) 
worksheet_summary.set_row(17,None,format_general_row) 
worksheet_summary.set_row(18,None,format_general_row) 
worksheet_summary.set_row(19,3,format_general_row) 
worksheet_summary.set_row(20,None,format_general_row) 
worksheet_summary.set_row(21,None,format_general_row) 
worksheet_summary.set_row(22,None,format_general_row) 
worksheet_summary.set_row(23,3,format_general_row) 
worksheet_summary.set_row(24,None,format_general_row) 
worksheet_summary.set_row(25,None,format_general_row) 
worksheet_summary.set_row(26,None,format_general_row) 
worksheet_summary.set_row(27,3,format_general_row) 
worksheet_summary.set_row(28,None,format_general_row) 
worksheet_summary.set_row(29,None,format_general_row) 
worksheet_summary.set_row(30,None,format_general_row) 
worksheet_summary.set_row(31,3,format_general_row) 
worksheet_summary.set_row(32,None,format_general_row) 
worksheet_summary.set_row(33,None,format_general_row) 
worksheet_summary.set_row(34,None,format_general_row) 
worksheet_summary.set_row(35,3,format_general_row) 
worksheet_summary.set_row(36,None,format_general_row) 
worksheet_summary.set_row(37,None,format_general_row) 
worksheet_summary.set_row(38,None,format_general_row) 
worksheet_summary.set_row(39,3,format_general_row) 
worksheet_summary.set_row(40,None,format_general_row) 
worksheet_summary.set_row(41,None,format_general_row) 
worksheet_summary.set_row(42,None,format_general_row) 
worksheet_summary.set_row(43,3,format_general_row) 
worksheet_summary.set_row(44,None,format_general_row) 
worksheet_summary.set_row(45,None,format_general_row) 
worksheet_summary.set_row(46,None,format_general_row) 
worksheet_summary.set_row(47,3,format_general_row) 
worksheet_summary.set_row(48,None,format_general_row) 
worksheet_summary.set_row(49,None,format_general_row) 
worksheet_summary.set_row(50,None,format_general_row) 
worksheet_summary.set_row(51,None,format_general_row) 
worksheet_summary.set_row(52,3) 
worksheet_summary.set_row(53,None,format_general_row) 
worksheet_summary.set_row(54,3) 
worksheet_summary.set_row(55,None,format_general_row)
worksheet_summary.set_row(56,None,format_general_row)
worksheet_summary.set_row(57,None,format_general_row)
worksheet_summary.set_row(58,3)
worksheet_summary.set_row(59,None,format_general_row)
# worksheet_summary.set_row(60,None,format_subtotal_row)

worksheet_summary.set_column('A:A',15,None)
worksheet_summary.set_column('B:B',13,None)
worksheet_summary.set_column('C:C',13,None,{'hidden':True})
worksheet_summary.set_column('D:D',15,None)
worksheet_summary.set_column('E:E',13,None)
worksheet_summary.set_column('F:F',13,None)
worksheet_summary.set_column('G:G',13,None)
worksheet_summary.set_column('H:H',2,None)
worksheet_summary.set_column('I:I',15,None)
worksheet_summary.set_column('J:J',13,None)
worksheet_summary.set_column('K:K',13,None,{'hidden':True})
worksheet_summary.set_column('L:L',13,None)
worksheet_summary.set_column('M:M',13,None)
worksheet_summary.set_column('N:N',13,None)
worksheet_summary.set_column('O:O',13,None)

worksheet_summary.write('A1', 'Page Links',format_top_summary)

worksheet_summary.write_url('A2',"internal:'Quantity Diff'!A1",format_url_links,string = '1. Quantity Diff')
worksheet_summary.write_url('A3',"internal:'PNL Diff'!A1",format_url_links,string = '2. PNL Diff')
worksheet_summary.write_url('A4',"internal:'Adj Unrealized PNL Change'!A1",format_url_links,string = '3. Adj Unrealized PNL Change')
worksheet_summary.write_url('A5',"internal:'Requirement Change'!A1",format_url_links,string = '4. Requirement Change')
worksheet_summary.write_url('A6',"internal:'HT Detail'!A1",format_url_links,string = '5. HT Detail')
worksheet_summary.write_url('A7',"internal:'TW Detail'!A1",format_url_links ,string = '6. TW Detail')

# Conditional Formating
worksheet_summary.conditional_format('M13:M60', {'type':'cell',
                                        'criteria': '<',
                                        'value':    0,
                                        'format':   format_general_row_green})
worksheet_summary.conditional_format('M13:M60', {'type':'cell',
                                        'criteria': '>',
                                        'value':    0,
                                        'format':   format_general_row_red})
worksheet_summary.conditional_format('N13:O60', {'type':'cell',
                                        'criteria': '<',
                                        'value':    0,
                                        'format':   format_general_row_red})
worksheet_summary.conditional_format('N13:O60', {'type':'cell',
                                        'criteria': '>',
                                        'value':    0,
                                        'format':   format_general_row_green})
"""
Create and format Detail sheet
"""



format1 = workbook.add_format({'num_format': '#,##0',
                               'font_size':'9'})
format2 = workbook.add_format({'num_format': '#,##0.00',
                               'font_size':'9'})

worksheet_summary.insert_image('M1', 'P:/1. Individual Folders/Chad/Python Scripts/PNL Report/Logo.png',{'x_scale':.7,'y_scale':.7})

# Format each colum to fit and display data correclty


Hilltop_x = Hilltop_Individual_Summary
Hilltop_y = Hilltop_x


def Summary_Individual_Sheets(Hilltop_Individual_Summary):
    column_summary = ('TW - HT Quantity Discrepancy',
                      'HT Change in Quantity',
                      'Adj Unreal PNL Change',
                      'HT-TW PNL Discrepancy',
                      'Requirement Change')
    for item in column_summary:
        Hilltop_Individual_Summary[item] = Hilltop_Individual_Summary[item]
    Hilltop_QTY_DSP = Hilltop_Individual_Summary[(Hilltop_Individual_Summary['TW - HT Quantity Discrepancy'] != 0)]
    Hilltop_HT_QTY_Change = Hilltop_Individual_Summary[(Hilltop_Individual_Summary['HT Change in Quantity'] != 0)]
    Hilltop_Adj_Unreal_PNL_Change = Hilltop_Individual_Summary[(Hilltop_Individual_Summary['Adj Unreal PNL Change'] != 0)]
    Hilltop_HT_TW_PNL_DSP = Hilltop_Individual_Summary[(Hilltop_Individual_Summary['HT-TW PNL Discrepancy'] != 0)]
    Hilltop_Requirement_Change = Hilltop_Individual_Summary[(Hilltop_Individual_Summary['Requirement Change'] != 0)]
    return Hilltop_QTY_DSP,Hilltop_HT_QTY_Change,Hilltop_Adj_Unreal_PNL_Change, Hilltop_HT_TW_PNL_DSP,Hilltop_Requirement_Change

Individual_Sheets = Summary_Individual_Sheets(Hilltop_Individual_Summary)


Hilltop_QTY_DSP  = Individual_Sheets[0]
Hilltop_QTY_DSP = pd.merge(Hilltop_QTY_DSP, Hilltop_x, on='Cusip', how='left')
Hilltop_QTY_DSP['TW - HT Quantity Discrepancy_x'] = Hilltop_QTY_DSP['TW - HT Quantity Discrepancy_x'].abs()
Hilltop_QTY_DSP.sort_values('TW - HT Quantity Discrepancy_x', axis=0, ascending=False, inplace=True)
Hilltop_QTY_DSP = Hilltop_QTY_DSP[['Security_x','Cusip','Account_x','TW - HT Quantity Discrepancy_y']]
Hilltop_QTY_DSP.rename(columns={'Security_x': 'Security','Account_x':'Account','TW - HT Quantity Discrepancy_y':"QTY DSP"}, inplace=True)
QTY_DSP_Cleared_Positions_Drop = QTY_DSP_Cleared_Positions.drop('Position Notes',axis = 1)
Hilltop_QTY_DSP = Hilltop_QTY_DSP.append(QTY_DSP_Cleared_Positions_Drop)
Hilltop_QTY_DSP.drop_duplicates(subset ='Cusip',keep = False, inplace = True)
Hilltop_QTY_DSP = Hilltop_QTY_DSP[['Security','Account','Cusip','QTY DSP']]
QTY_DSP_Cleared_Positions= QTY_DSP_Cleared_Positions[['Security','Account','Cusip','QTY DSP','Position Notes']]


Hilltop_HT_TW_PNL_DSP = Individual_Sheets[3]
Hilltop_HT_TW_PNL_DSP.sort_values('HT-TW PNL Discrepancy', axis=0, ascending=False, inplace=True)
Hilltop_HT_TW_PNL_DSP = pd.merge(Hilltop_HT_TW_PNL_DSP, Hilltop_x, on='Cusip', how='left')
Hilltop_HT_TW_PNL_DSP['HT-TW PNL Discrepancy_x'] = Hilltop_HT_TW_PNL_DSP['HT-TW PNL Discrepancy_x'].abs()
Hilltop_HT_TW_PNL_DSP.sort_values('HT-TW PNL Discrepancy_x', axis=0, ascending=False, inplace=True)
Hilltop_HT_TW_PNL_DSP = Hilltop_HT_TW_PNL_DSP[['Security_x','Cusip','Account_x','HT-TW PNL Discrepancy_y']]
Hilltop_HT_TW_PNL_DSP_Lower = Hilltop_HT_TW_PNL_DSP[(Hilltop_HT_TW_PNL_DSP['HT-TW PNL Discrepancy_y'] < -10)]
Hilltop_HT_TW_PNL_DSP_Upper = Hilltop_HT_TW_PNL_DSP[(Hilltop_HT_TW_PNL_DSP['HT-TW PNL Discrepancy_y'] > 10)]
Hilltop_Chunks = [Hilltop_HT_TW_PNL_DSP_Upper,Hilltop_HT_TW_PNL_DSP_Lower]
Hilltop_HT_TW_PNL_DSP = pd.concat(Hilltop_Chunks)
Hilltop_HT_TW_PNL_DSP.sort_values(by = 'HT-TW PNL Discrepancy_y',ascending = False)
x = subject_name(file_text)
y = x[3]
PNL_DSP_Date = y[:10]
Hilltop_HT_TW_PNL_DSP['Date'] = PNL_DSP_Date
Hilltop_HT_TW_PNL_DSP = Hilltop_HT_TW_PNL_DSP[['Date','Security_x','Account_x','Cusip','HT-TW PNL Discrepancy_y',   ]]


"""
Pull and Generate new PNL DIFF Position File
"""
# PNL_Report_Date = x[2]
# PNL_Report_File_Most_Recent = 'P:/2. Corps/PNL_Daily_Report/Reports/PNL_Report.xlsx'
# PNL_DSP_Yesterday = pd.read_excel(PNL_Report_File_Most_Recent,sheet_name = 'PNL Diff')  #DSP Items from previous report
# Additions_to_Running_PNL_DSP = PNL_DSP_Yesterday[['Date','Security','Account','Cusip','PNL DSP']]  #DSP Items from previous report sorted
# Additions_to_Running_PNL_DSP.dropna(inplace = True)
# Current_Running_PNL_DSP_Filepath = 'P:/2. Corps/PNL_Daily_Report/Cleared_Position_File/PNL_DSP_History.xlsx' 
# Current_Running_PNL_DSP = pd.read_excel(Current_Running_PNL_DSP_Filepath)# reads running PNL DSP file
# Current_Running_PNL_DSP = Current_Running_PNL_DSP[Current_Running_PNL_DSP['Previous PNL DSP'] > 5] # filters out 'closed' Positions
# Additions_to_Running_PNL_DSP['Previous PNL DSP'] = Additions_to_Running_PNL_DSP['PNL DSP']
# Additions_to_Running_PNL_DSP = Additions_to_Running_PNL_DSP[['Account','Cusip','Date','Previous PNL DSP','Security']]
# Current_Running_PNL_DSP_List = [Current_Running_PNL_DSP,Additions_to_Running_PNL_DSP]                                  # creates a list to concat the dataframes together
# Complete_Running_PNL_DSP = pd.concat(Current_Running_PNL_DSP_List)     # Concatinate the running PNL DSP and the Most recent DSP
# Complete_Running_PNL_DSP.drop_duplicates(subset = 'Cusip', keep = 'first', inplace = True)
# Complete_Running_PNL_DSP_with_Detail = pd.merge(Complete_Running_PNL_DSP,Hilltop_Recent, on = 'Cusip', how = 'left')
# Complete_Running_PNL_DSP_with_Detail['Net PNL DSP'] = Complete_Running_PNL_DSP_with_Detail['Previous PNL DSP'] + Complete_Running_PNL_DSP_with_Detail['HT-TW PNL Discrepancy']
# Complete_Running_PNL_DSP_with_Detail = Complete_Running_PNL_DSP_with_Detail[['Date','Security_x','Account_x','Cusip','Previous PNL DSP','HT-TW PNL Discrepancy','Net PNL DSP']]

# Complete_Running_PNL_DSP_with_Detail.rename(columns={'Date':'Date',
#                                         'Security_x':'Security',
#                                         'Account_x':'Account',
#                                         'Cusip':'Cusip',
#                                         'Previous PNL DSP':'Previous PNL DSP',
#                                          'HT-TW PNL Discrepancy':'Current PNL DSP'
#                                         },inplace = True)

Hilltop_Adj_Unreal_PNL_Change = Individual_Sheets[2]
Hilltop_Adj_Unreal_PNL_Change = pd.merge(Hilltop_Adj_Unreal_PNL_Change, Hilltop_x, on='Cusip', how='left')
Hilltop_Adj_Unreal_PNL_Change['Adj Unreal PNL Change_x'] = Hilltop_Adj_Unreal_PNL_Change['Adj Unreal PNL Change_x'].abs()
Hilltop_Adj_Unreal_PNL_Change.sort_values('Adj Unreal PNL Change_x', axis=0, ascending=False, inplace=True)
Hilltop_Adj_Unreal_PNL_Change = Hilltop_Adj_Unreal_PNL_Change[['Security_x','Cusip','Account_x','Adj Unreal PNL Change_y']]



Hilltop_Requirement_Change = Individual_Sheets[4]
Hilltop_Requirement_Change = pd.merge(Hilltop_Requirement_Change, Hilltop_x, on='Cusip', how='left')
Hilltop_Requirement_Change = Hilltop_Requirement_Change[['Cusip','Security_x','Account_x','Requirement Change_x','Requirement Change_y']]
Hilltop_Requirement_Change['Requirement Change_x'] = Hilltop_Requirement_Change['Requirement Change_x'].abs()
Hilltop_Requirement_Change.sort_values('Requirement Change_x', axis=0, ascending=False, inplace=True)
Hilltop_Requirement_Change = Hilltop_Requirement_Change[['Security_x','Cusip','Account_x','Requirement Change_y']]


"""
Write to Excel

"""

QTY_DSP_Cleared_Positions.to_excel(writer,sheet_name ='Quantity Diff',index=False,startrow=1,startcol=6)
Hilltop_QTY_DSP.to_excel(writer, sheet_name = 'Quantity Diff', index=False)
worksheet_Hilltop_QTY_DSP = writer.sheets['Quantity Diff']

Hilltop_HT_TW_PNL_DSP.to_excel(writer, sheet_name = 'PNL Diff', index=False)
Complete_Running_PNL_DSP_with_Detail.to_excel(writer, sheet_name = 'PNL Diff', index=False,startrow = 1,startcol=7)
worksheet_Hilltop_HT_TW_PNL_DSP = writer.sheets['PNL Diff']

Hilltop_Adj_Unreal_PNL_Change.to_excel(writer, sheet_name = 'Adj Unrealized PNL Change', index=False)
worksheet_Hilltop_Adj_Unreal_PNL_Change = writer.sheets['Adj Unrealized PNL Change']

Hilltop_Requirement_Change.to_excel(writer, sheet_name = 'Requirement Change', index=False)
worksheet_Hilltop_Requirement_Change = writer.sheets['Requirement Change']


Hilltop_Recent.to_excel(writer, sheet_name='HT Detail', index=False)
worksheet = writer.sheets['HT Detail']

worksheet_Hilltop_QTY_DSP.set_column('A:A', 35,format1)#, format7) #2
worksheet_Hilltop_QTY_DSP.set_column('B:B', 15,format1)#, format2) #2
worksheet_Hilltop_QTY_DSP.set_column('C:C', 12,format1)#, format7) #3
worksheet_Hilltop_QTY_DSP.set_column('D:D', 12,format1)#, format1)#4
worksheet_Hilltop_QTY_DSP.set_column('D:D', 12,format1)#, format1)#4
worksheet_Hilltop_QTY_DSP.set_column('E:E', 30,format1)#, format1)

worksheet_Hilltop_QTY_DSP.write('A1', 'Security',format_top_summary)#,format5)
worksheet_Hilltop_QTY_DSP.write('B1', 'Account',format_top_summary)#,format5)
worksheet_Hilltop_QTY_DSP.write('C1', 'Cusip',format_top_summary)#,format5)
worksheet_Hilltop_QTY_DSP.write('D1', 'QTY DSP',format_top_summary)#,format5)
worksheet_Hilltop_QTY_DSP.write('E1', 'Position Notes',format_top_summary)#,format5)

worksheet_Hilltop_QTY_DSP.write('G2', 'Security',format_top_summary)#,format5)
worksheet_Hilltop_QTY_DSP.write('H2', 'Account',format_top_summary)#,format5)
worksheet_Hilltop_QTY_DSP.write('I2', 'Cusip',format_top_summary)#,format5)
worksheet_Hilltop_QTY_DSP.write('J2', 'QTY DSP',format_top_summary)#,format5)
worksheet_Hilltop_QTY_DSP.write('K2', 'Position Notes',format_top_summary)#,format5)

worksheet_Hilltop_QTY_DSP.merge_range('G1:K1', 'Cleared QTY DSP',merge_format)

worksheet_Hilltop_QTY_DSP.set_column('G:G', 20,format1)#, format1) #2
worksheet_Hilltop_QTY_DSP.set_column('H:H', 12,format1)#, format1) #2
worksheet_Hilltop_QTY_DSP.set_column('I:I', 12,format1)#, format1) #3
worksheet_Hilltop_QTY_DSP.set_column('J:J', 15,format1)#, format1)#4
worksheet_Hilltop_QTY_DSP.set_column('K:K', 30,format1)#, format1)#4

# worksheet_Hilltop_QTY_DSP.freeze_panes(1, 1)
worksheet_Hilltop_QTY_DSP.autofilter('A1:V20000')
# worksheet_Hilltop_QTY_DSP.hide_gridlines(2)
# worksheet_Hilltop_QTY_DSP.protect('welcome123')
worksheet_Hilltop_QTY_DSP.set_zoom(90)
"""
"""
worksheet_Hilltop_Adj_Unreal_PNL_Change.set_column('A:A', 35,format1)#, format7) #2
worksheet_Hilltop_Adj_Unreal_PNL_Change.set_column('B:B', 12,format1)#, format2) #2
worksheet_Hilltop_Adj_Unreal_PNL_Change.set_column('C:C', 15,format1)#, format7) #3
worksheet_Hilltop_Adj_Unreal_PNL_Change.set_column('D:D', 12,format1)#, format1)#4

worksheet_Hilltop_Adj_Unreal_PNL_Change.write('A1', 'Security',format_top_summary)#,format5)
worksheet_Hilltop_Adj_Unreal_PNL_Change.write('B1', 'Cusip',format_top_summary)#,format5)
worksheet_Hilltop_Adj_Unreal_PNL_Change.write('C1', 'Account',format_top_summary)#,format5)
worksheet_Hilltop_Adj_Unreal_PNL_Change.write('D1', 'Adj Unreal PNL Change',format_top_summary)#,format5)
worksheet_Hilltop_Adj_Unreal_PNL_Change.set_zoom(90)
"""
"""
worksheet_Hilltop_HT_TW_PNL_DSP.set_column('A:A', 11,format1)#, format7) #2
worksheet_Hilltop_HT_TW_PNL_DSP.set_column('B:B', 25,format1)#, format2) #2
worksheet_Hilltop_HT_TW_PNL_DSP.set_column('C:C', 13,format1)#, format7) #3
worksheet_Hilltop_HT_TW_PNL_DSP.set_column('D:D', 11,format1)#, format1)#4
worksheet_Hilltop_HT_TW_PNL_DSP.set_column('E:E', 11,format1)#, format1)#4
worksheet_Hilltop_HT_TW_PNL_DSP.set_column('F:F', 15,format1)#, format1)#4

worksheet_Hilltop_HT_TW_PNL_DSP.write('A1', 'Date',format_top_summary)#,format5)
worksheet_Hilltop_HT_TW_PNL_DSP.write('B1', 'Security',format_top_summary)#,format5)
worksheet_Hilltop_HT_TW_PNL_DSP.write('C1', 'Account',format_top_summary)#,format5)
worksheet_Hilltop_HT_TW_PNL_DSP.write('D1', 'Cusip',format_top_summary)#,format5)
worksheet_Hilltop_HT_TW_PNL_DSP.write('E1', 'PNL DSP',format_top_summary)#,format5)
# worksheet_Hilltop_HT_TW_PNL_DSP.write('F1', 'Position Notes',format_top_summary)#,format5)


worksheet_Hilltop_HT_TW_PNL_DSP.write('H2', 'Date',format_top_summary)#,format5)
worksheet_Hilltop_HT_TW_PNL_DSP.write('I2', 'Security',format_top_summary)#,format5)
worksheet_Hilltop_HT_TW_PNL_DSP.write('J2', 'Account',format_top_summary)#,format5)
worksheet_Hilltop_HT_TW_PNL_DSP.write('K2', 'Cusip',format_top_summary)#,format5)
worksheet_Hilltop_HT_TW_PNL_DSP.write('L2', 'Previous PNL',format_top_summary)#,format5)
worksheet_Hilltop_HT_TW_PNL_DSP.write('M2', 'Current PNL',format_top_summary)#,format5)
worksheet_Hilltop_HT_TW_PNL_DSP.write('N2', 'Net PNL',format_top_summary)#,format5)

worksheet_Hilltop_HT_TW_PNL_DSP.merge_range('H1:N1', 'Unresolved PNL DSP',merge_format)

worksheet_Hilltop_HT_TW_PNL_DSP.set_column('G:G', 3,format1)#, format1) #2
worksheet_Hilltop_HT_TW_PNL_DSP.set_column('H:H', 12,format1)#, format1) #2
worksheet_Hilltop_HT_TW_PNL_DSP.set_column('I:I', 12,format1)#, format1) #3
worksheet_Hilltop_HT_TW_PNL_DSP.set_column('J:J', 15,format1)#, format1)#4
worksheet_Hilltop_HT_TW_PNL_DSP.set_column('K:K', 15,format1)#, format1)#4
worksheet_Hilltop_HT_TW_PNL_DSP.set_column('L:L', 15,format1)#, format1)#4
worksheet_Hilltop_HT_TW_PNL_DSP.set_column('M:M', 15,format1)#, format1)#4
worksheet_Hilltop_HT_TW_PNL_DSP.set_column('N:N', 15,format1)#, format1)#4
worksheet_Hilltop_HT_TW_PNL_DSP.set_zoom(90)

"""
"""
worksheet_Hilltop_Requirement_Change.set_column('A:A', 35,format1)#, format7) #2
worksheet_Hilltop_Requirement_Change.set_column('B:B', 15,format1)#, format2) #2
worksheet_Hilltop_Requirement_Change.set_column('C:C', 15,format1)#, format7) #3
worksheet_Hilltop_Requirement_Change.set_column('D:D', 15,format1)#, format1)#4

worksheet_Hilltop_Requirement_Change.write('A1', 'Security',format_top_summary)#,format5)
worksheet_Hilltop_Requirement_Change.write('B1', 'Cusip',format_top_summary)#,format5)
worksheet_Hilltop_Requirement_Change.write('C1', 'Account',format_top_summary)#,format5)
worksheet_Hilltop_Requirement_Change.write('D1', 'Requirement Change',format_top_summary)#,format5)
worksheet_Hilltop_Requirement_Change.set_zoom(90)


#
# worksheet_Hilltop_QTY_DSP.freeze_panes(1, 1)
worksheet_Hilltop_QTY_DSP.autofilter('A1:E20000')
# worksheet_Hilltop_QTY_DSP.hide_gridlines(2)

# worksheet_Hilltop_HT_TW_PNL_DSP.freeze_panes(1, 1)
worksheet_Hilltop_HT_TW_PNL_DSP.autofilter('A1:E20000')
# worksheet_Hilltop_HT_TW_PNL_DSP.hide_gridlines(2)


# worksheet_Hilltop_Adj_Unreal_PNL_Change.freeze_panes(1, 1)
worksheet_Hilltop_Adj_Unreal_PNL_Change.autofilter('A1:D20000')
# worksheet_Hilltop_Adj_Unreal_PNL_Change.hide_gridlines(2)

# worksheet_Hilltop_Requirement_Change.freeze_panes(1, 1)
worksheet_Hilltop_Requirement_Change.autofilter('A1:D20000')
# worksheet_Hilltop_Requirement_Change.hide_gridlines(2)

worksheet_summary.hide_gridlines(2)
# worksheet_TW_Detail = writer.sheets['TW Detail']


worksheet.set_column('A:A', 34,format1)#, format7) #2
worksheet.set_column('B:B', 13,format1)#, format2) #2
worksheet.set_column('C:C', 14,format1)#, format2) #3
worksheet.set_column('D:D', 11,format2)#, format1)#4
worksheet.set_column('E:E', 11,format1)#, format1)#5
worksheet.set_column('F:F', 11,format1)#, format1)#6
worksheet.set_column('G:G', 11,format1)#, format1)#7
worksheet.set_column('H:H', 11,format1)#, format1)#8
worksheet.set_column('I:I', 11,format1)#, format1)#9
worksheet.set_column('J:J', 11,format1)#, format1)#10
worksheet.set_column('K:K', 11,format1)#, format1)#11
worksheet.set_column('L:L', 11,format1)#, format1)#12
worksheet.set_column('M:M', 11,format1)#, format1)#13
worksheet.set_column('N:N', 11,format1)#, format1)#14
worksheet.set_column('O:O', 11,format1)#, format1)#15
worksheet.set_column('P:P', 15,format1)

worksheet.write('A1', 'Security',format_top_summary)#,format5)
worksheet.write('B1', 'Cusip',format_top_summary)#,format5)
worksheet.write('C1', 'Account',format_top_summary)#,format5)
worksheet.write('D1', 'Price',format_top_summary)#,format5)
worksheet.write('E1', 'TW QTY ',format_top_summary)#,format5)
worksheet.write('F1', 'HT QTY',format_top_summary)#,format5)
worksheet.write('G1', 'QTY Discrepancy',format_top_summary)#,format5)
worksheet.write('H1', 'HT QTY Change',format_top_summary)#,format5)
worksheet.write('I1', 'HT New Unreal PNL',format_top_summary)#,format5)
worksheet.write('J1', 'HT Old Unreal PNL',format_top_summary)#,format5)
worksheet.write('K1', 'Real PNL Change',format_top_summary)#,format5)
worksheet.write('L1', 'Adj Unreal PNL Change',format_top_summary)#,format5)
worksheet.write('M1', 'TW PNL',format_top_summary)#,format5)
worksheet.write('N1', 'HT-TW PNL Discrep.',format_top_summary)#,format5)
worksheet.write('O1', 'Req. Change',format_top_summary)#,format5)
worksheet.write('P1', 'Requirement',format_top_summary)#,format5)

worksheet.set_zoom(90)

worksheet.autofilter('A1:O20000')

TW_Detail.reset_index(inplace = True)
TW_Detail.to_excel(writer,sheet_name = 'TW Detail',index = False)
worksheet_TW_Detail = writer.sheets['TW Detail']
worksheet_TW_Detail.write('A1', 'Cusip',format_top_summary)#,format5)
worksheet_TW_Detail.write('B1', 'P&L',format_top_summary)#,format5)
worksheet_TW_Detail.write('C1', 'Security',format_top_summary)#,format5)
worksheet_TW_Detail.write('D1', 'Position',format_top_summary)#,format5)
worksheet_TW_Detail.write('E1', 'Symbol',format_top_summary)#,format5)
worksheet_TW_Detail.write('F1', 'Book',format_top_summary)#,format5)
worksheet_TW_Detail.write('G1', 'MTG Position',format_top_summary)#,format5)
worksheet_TW_Detail.set_column('A:A', 12,format1)#, format7) #2
worksheet_TW_Detail.set_column('B:B', 12,format1)#, format2) #2
worksheet_TW_Detail.set_column('C:C', 25,format1)#, format7) #3
worksheet_TW_Detail.set_column('D:D', 12,format1)#, format1)#4
worksheet_TW_Detail.set_column('E:E', 15,format1)#, format1)#4
worksheet_TW_Detail.set_column('F:F', 12,format1)#, format1)#4
worksheet_TW_Detail.set_column('G:G', 12,format1)#, format1)#4
worksheet_TW_Detail.autofilter('A1:G20000')
worksheet_TW_Detail.set_zoom(90)

workbook.close()



Complete_Running_History = Complete_Running_PNL_DSP_with_Detail
Complete_Running_History['Previous PNL DSP'] = Complete_Running_History['Net PNL DSP']
Complete_Running_History = Complete_Running_History[['Date','Security','Account','Cusip','Previous PNL DSP']]
writer = pd.ExcelWriter('P:/2. Corps/PNL_Daily_Report/Cleared_Position_File/PNL_DSP_History.xlsx', engine='xlsxwriter')
Complete_Running_History.to_excel(writer)
writer.save()


#For Quantity Diff and PNL Diff tabs, make notes in the "Position Notes" column to clear each line item.\n2. Save changes by following these steps:\n     1. Click "Save As"\n     2. Click "Browse\n     3. Copy and Paste the following into the top search bar "P:/2. Corps/PNL_Daily_Report/Reports"\n     4. Overwrite and save the File as "PNL_Report.xlsx" '
# if Firm_Real_PNL_Daily_Total == Daily_Total_Real_PNL_Check_Value:
#     if Firm_Unreal_PNL_Daily_Total == Daily_Total_Unreal_PNL_Check_Value:
#         if Firm_Requirement_Daily_Total == Daily_Total_Requirement_Check_Value:
#             if Firm_Market_Value_Daily_Total == Daily_Total_Market_Value_Check_Value:
#                 if Firm_Real_PNL_Summary_Total == Summary_Total_Real_PNL_Check_Value:
#                     if Firm_Unreal_PNL_Summary_Total == Summary_Total_Unreal_PNL_Check_Value:
#                         if Firm_Requirement_Summary_Total == Summary_Total_Requirement_Check_Value:
#                             if Firm_Market_Value_Summary_Total == Summary_Total_Market_Value_Check_Value:
#                                 text_body = 'Daily '+today+instructions
#                             else:
#                                 text_body = 'Daily '+today+'. Errors Found in Monthly Market Value' + instructions  
#                         else:
#                             text_body = 'Daily '+today+'. Errors Found in Monthly Requirement Value' + instructions
#                     else:
#                         text_body = 'Daily '+today+'. Errors Found in Monthly Unreal PNL Value' + instructions
#                 else:
#                     text_body = 'Daily '+today+'. Errors Found in Monthly Real PNL Value' + instructions
#             else:
#                 text_body = 'Daily '+today+'. Errors Found in Daily Market Value Value' + instructions
#         else:
#             text_body = 'Daily '+today+'. Errors Found in Daily Requirement Value' + instructions
#     else:
#         text_body = 'Daily '+today+'. Errors Found in Daily Unreal PNL Value' + instructions
# else:
#     text_body = 'Daily '+today+'. Errors Found in Daily Real PNL Value' + instructions

import win32com.client
from win32com.client import Dispatch, constants
const=win32com.client.constants

olMailItem = 0x0
obj = win32com.client.Dispatch("Outlook.Application")
newMail = obj.CreateItem(olMailItem)
newMail.Subject = "Inventory Report"
newMail.To = 'ccraig@sierrapacificsecurities.com;jblamire@sierrapacificsecurities.com;elankowsky@sierrapacificsecurities.com;jdean@sierrapacificsecurities.com;bburdick@sierrapacificsecurities.com'
newMail.Body = 'To clear items, click the link below and save using ctrl-S or File - Save'
newMail.HTMLBody ='<a href="file:///P:/2.%20Corps/PNL_Daily_Report/Reports/PNL_Report.xlsx">Inventory Report</a>'
newMail.Attachments.Add('C:/Users/ccraig/Desktop/New folder/'+ str(current) + ' Inventory Report.xlsx')
newMail.Send()
