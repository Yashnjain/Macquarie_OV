import xlwings as xw
import shutil    
from datetime import datetime,timedelta,date
import numpy as np
import pandas as pd
import os
import calendar
import logging
import bu_alerts
import time
from tabula import read_pdf

def MACQUARIE_OV_MONTHEND(date,fname):
        url = 'file:///C:/Users/Pakhi.laad/Macquarie_ov/downloads/3101F.pdf'
        logging.info('Inside MACQUARIE_OV_MONTHEND')
        loc = f'S:\Position Report\MBL Statement Recon\Source\{date}\{fname}'    
        # df  = read_pdf(url, pages = 'all', guess = False, stream = True,
        #                      pandas_options={'header':None}, area = ["18.754,19.176,727.144,595.986"], columns=["80,116,144,450,482,520"])
        df  = read_pdf(url, pages = 'all', guess = False, stream = True,
                             pandas_options={'header':None}, area = ["18.754,19.176,727.144,595.986"], columns=["80,116,140,430,482,520"])
        df  = read_pdf(url, pages = 'all', guess = False, stream = True,
                       pandas_options={'header':None}, area = ["18.754,19.176,727.144,595.986"], columns=["80,116,140,415,482,520"])
        main_df = pd.concat(df, ignore_index=True)
        start = main_df.index[main_df[3] == 'OPEN POSITIONS'].tolist()
        end = main_df.index[main_df[3] == '** This is the end of the open position summary **'].tolist()
        i = 1
        j = 1
        list = []
        list.append((start[0],end[0]))
        while i < len(start) and j < len(end):
            if(start[i]<end[j-1]):
                i = i+1
            else:
                list.append((start[i],end[j]))
                i = i+1
                j = j+1

        print(list)
        for i in range(len(list)):
            a = main_df[list[i][0]:list[i][1]]
            des = a.index[a[3]=="Description Price"]
            sp_index = a.index[a[3].apply(lambda x: str(x).startswith("S.P"))]
            for i in sp_index:
                bought = a.loc[i][1]
                sold = a.loc[i][2]
                price = a.loc[i][4]
                valuation = a.loc[i][6]









        dic = {'52311430' : 'sheet1','52311431': [34,'WC-431'],'52311432': [23,'GC-432'],'52311433': [18,'Power-433'],'52311434': [35,'Bulk-434'],'52311435': [14,'Power-435'],\
                '52311436': [30,'Power-436'],'52311437': [14,'Power-437'],'52311438': [14,'Power-438'],'52311439': [15,'Spread-439'],'52311440': [42,'Spread-440'],'52311441': [16,'NG-441'],\
                '52311442': [27,'Center-442'],'52311443': [20,'Center-443'],'52311444': [24,'Center-444'], '52311445': [57,'Power-445'],'52311446': [33,'Power-446'],'52311448': [33,'Power-448']}
        logging.info('Dataframe made from TC file')
        today = datetime.now()
        year = time.strftime("%Y")
        month = time.strftime("%m")
        days_in_month = str(calendar.monthrange(int(year), int(month))[1])
        folder_name = inputloc + '\\' + today.strftime("%Y%m") + '\\Test'
        filename = folder_name + '\\'+'MBL-'+ today.strftime("%Y%m")+days_in_month+'.xlsx'
        if not os.path.exists(folder_name):
            logging.info(' New month : Folder not found making new Folder')
            os.makedirs(folder_name)
        if not os.path.exists(filename):
            logging.info(' New month : File not found making new File')
            yesterday = datetime.now()-timedelta(days=1)
            year = yesterday.strftime("%Y")
            month = yesterday.strftime("%m")
            days_in_month = str(calendar.monthrange(int(year), int(month))[1])
            pre_folder_name = inputloc + '\\' + yesterday.strftime("%Y%m")
            pre_filename = pre_folder_name +'\\'+'MBL-'+ today.strftime("%Y%m")+days_in_month+'.xlsx'
            shutil.copy(pre_filename, folder_name)
            os.rename(pre_filename,filename)
        wb = xw.Book(filename)
        logging.info('File loaded in workbook')

def main():
    try:
        logging.info('process started')
        date_day_before= datetime.now() - timedelta(1)
        date_file = date_day_before.strftime("%m%d%Y")
        fname = date_day_before.strftime("%d%m") + 'F'
        log_json='[{"JOB_ID": "'+str(job_id)+'","CURRENT_DATETIME": "'+str(datetime.now())+'"}]'
        logging.info('Calling Macquarie_ov')
        MACQUARIE_OV_MONTHEND(date_file,fname)
        logging.info(' Engine Disposed ---- end')
        bu_alerts.send_mail(
                    receiver_email = 'pakhi.laad@biourja.com',
                    mail_subject ='JOB SUCCESS - MACQUARIE OTHER_VERTICAL MONTHEND',
                    mail_body=f"Process completed successfully for all the Projects",
                    attachment_location = log_file_location
                )
    except Exception as e:
        logger.exception(f'Error occurred in . {e}')
        bu_alerts.send_mail(
                            receiver_email= 'pakhi.laad@biourja.com',
                            mail_subject=f"JOB FAILED ::  MACQUARIE OTHER_VERTICAL MONTHEND",
                            mail_body=f"{e}",
                            attachment_location = log_file_location)
          


if __name__ == '__main__':
    today_date = date.today()
    today = today_date.strftime("%m%d%Y")
    job_name = "Macquarie_OV_monthend"
    log_file_location =os.getcwd()+'\\logs\\' + str(job_name)+str(today_date)+'.txt'
        
    logging.basicConfig(
    level=logging.INFO, 
    format='%(asctime)s [%(levelname)s] - %(message)s',
    filename=log_file_location)
    # inputloc =  os.getcwd()+'\\downloads\\'
    inputloc =  r'\\BIO-INDIA-FS\India Sync$\India\Macquarie\2023'
    job_id=np.random.randint(1000000,9999999)
    logger = logging.getLogger(__name__)
    logger.setLevel(logging.INFO) 
    

    main()