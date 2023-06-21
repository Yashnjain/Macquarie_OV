import xlwings as xw
import shutil    
import datetime
from datetime import datetime,timedelta,date
import numpy as np
import pandas as pd
import os
import calendar
import logging
import bu_alerts
import time

     
def MACQUARIE_OV(date,fname):
    try:
        logging.info('Inside MACQUARIE_OV')
        loc = f'S:\Position Report\MBL Statement Recon\Source\{date}\{fname}'
        df  = pd.read_csv(os.path.join(loc,'TC'+fname+'.csv'))
        dic = {'52311430' : 'sheet1','52311431': [34,'WC-431'],'52311432': [23,'Power-432'],'52311433': [18,'Power-433'],'52311434': [35,'Bulk-434'],'52311435': [14,'Power-435'],\
                '52311436': [30,'Power-436'],'52311437': [14,'Power-437'],'52311438': [14,'Power-438'],'52311439': [15,'Spread-439'],'52311440': [42,'Spread-440'],'52311441': [16,'NG-441'],\
                '52311442': [27,'Center-442'],'52311443': [20,'Center-443'],'52311444': [24,'Center-444'], '52311445': [57,'Power-445'],'52311446': [33,'Power-446'],'52311448': [33,'Power-448']}
        logging.info('Dataframe made from TC file')
        # dic_list = list(dic.items())[11:]
        daybefore= datetime.now() - timedelta(days =1)
        year = daybefore.year
        month = daybefore.month
        days_in_month = str(calendar.monthrange(int(year), int(month))[1])
        folder_name = inputloc + '\\' + daybefore.strftime("%Y%m") + '\\Test'
        filename = folder_name + '\\'+'MBL-'+ daybefore.strftime("%Y%m")+ days_in_month+'.xlsx'
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
            pre_filename = pre_folder_name +'\\Test\\'+'MBL-'+ yesterday.strftime("%Y%m")+days_in_month+'.xlsx'
            copied_file_name = folder_name +'\\'+ 'MBL-'+ yesterday.strftime("%Y%m")+days_in_month+'.xlsx'
            shutil.copy(pre_filename, folder_name)
            os.rename(copied_file_name,filename)
            wb = xw.Book(filename)
            for key,value in dic.items():
                if key == '52311430':
                    sheet = wb.sheets['Margin-430']
                    last_row = sheet.range('A' + str(sheet.cells.last_cell.row)).end('up').row
                    sec_last_row = sheet.range('A' + str(sheet.cells.last_cell.row)).end('up').end('up').row
                    lines_to_copy = sheet.range(f'A{sec_last_row+1}:AB{last_row}').value
                    sheet.range(f'A{last_row+1}').value = lines_to_copy
                    last_row = sheet.range('A' + str(sheet.cells.last_cell.row)).end('up').row
                    current_date = datetime.now()
                    next_month = current_date.month + 1 if current_date.month < 12 else 1
                    next_year = current_date.year + 1 if current_date.month == 12 else current_date.year
                    first_day_next_month = datetime(next_year, next_month, 1)
                    next_month_date = first_day_next_month.strftime("%m-%d-%Y")
                    sheet.range(f"A{last_row}").value = next_month_date
                else:
                    sheet = wb.sheets[value[1]]
                    sheet.activate()
                    try:
                        last_row = sheet.range('A' + str(sheet.cells.last_cell.row)).end('up').row
                        sec_last_row = sheet.range('A' + str(sheet.cells.last_cell.row)).end('up').end('up').end('up').end('up').row
                        lines_to_copy = sheet.range(f'A{sec_last_row+1}:AB{last_row}').value
                    except AttributeError:
                        last_row = sheet.range('A' + str(sheet.cells.last_cell.row)).end('up').row
                        sec_last_row = sheet.range('A' + str(sheet.cells.last_cell.row)).end('up').end('up').end('up').end('up').end('up').end('up').row
                        lines_to_copy = sheet.range(f'A{sec_last_row+1}:AB{last_row}').value
                    sheet.range(f'A{last_row+1}').value = lines_to_copy
                    last_row = sheet.range('A' + str(sheet.cells.last_cell.row)).end('up').row
                    current_date = datetime.now()
                    next_month = current_date.month + 1 if current_date.month < 12 else 1
                    next_year = current_date.year + 1 if current_date.month == 12 else current_date.year
                    first_day_next_month = datetime(next_year, next_month, 1)
                    next_month_date = first_day_next_month.strftime("%m-%d-%Y")
                    sheet.range(f"A{last_row}").value = next_month_date
                
                continue

            wb.save()
            logging.info('Save changes to workbook')
            wb.close()
        wb = xw.Book(filename)
        logging.info('File loaded in workbook')
        time.sleep(5)
        for key,value in dic.items():
        # for key,value in dic_list:
            val = df[df["Client code"] ==  int(key)].reset_index()
            if(val.empty):
                    continue
            else:
                val['Input Date'] = val['Input Date'].apply(lambda x: datetime.strptime(x,"%d/%m/%y"))
                Input_Date = val['Input Date'] 
                trade_date = val['Trade Date']
                amount = val['Settlement Amount']
                bought_quantity = val['Bought Quantity']
                sold_quantity = val['Sold  Quantity']
                bought_price = val['Bought Price']
                sold_price = val['Sold Price']
                exec_comm = val['Executing Commission Amount']
                fees = val['Exchange Fee Amount']
                clr_comm = val['Clearing Commission Amount']
                NFA_Fees = val['EFP Amount']
                desc = val['Description']
                if key == '52311430':
                    sheet = wb.sheets["Margin-430"]
                    last_row = sheet.range('A' + str(sheet.cells.last_cell.row)).end('up').row
                    sheet.range(f"A{last_row+1}").api.EntireRow.Insert()
                    time.sleep(1)     
                            #inserting Date in A column last row
                    for i in range(len(val)):
                        #two line space between final month entrty and current entry
                        sheet.range(f'A{last_row-2}').api.EntireRow.Insert()
                        sheet.range(f"A{last_row-2}").value = Input_Date[i].strftime("%m-%d-%Y")
                        sheet.range(f"Y{last_row-2}").value = amount[i]
                        amount[i] = amount[i].replace(" ","")
                        if(float(amount[i])>0):
                            sheet.range(f"B{last_row-2}").value = 'WIRE TRFR RECVD'
                        else :
                            sheet.range(f"B{last_row-2}").value = 'WIRE TRFR SENT'
                        sheet.range(f"AB{last_row-2}").formula = f'=+AB{last_row-3}+I{last_row-2}+N{last_row-2}+W{last_row-2}+Y{last_row-2}'
                        last_row += 1
                        logging.info('data went for first sheet')
                    continue
                else:
                        sheet = wb.sheets[value[1]]
                        sheet.activate()
                        logging.info(f'sheet activated for {value[1]}')
                        for i in range(len(val)):
                            exec_comm[i] = exec_comm[i].replace(" ","")
                            fees[i] = fees[i].replace(" ","")
                            clr_comm[i] = clr_comm[i].replace(" ","")
                            NFA_Fees[i] = NFA_Fees[i].replace(" ","")
                            trade_date[i] = trade_date[i].replace(" ","")
                            bought_quantity[i] = bought_quantity[i].replace(" ","")
                            sold_quantity[i] = sold_quantity[i].replace(" ","")
                            exec_comm_val = float(exec_comm[i]) if exec_comm[i] else 0.0
                            fees_val = float(fees[i]) if fees[i] else 0.0
                            clr_comm_val = float(clr_comm[i]) if clr_comm[i] else 0.0
                            NFA_Fees_val = float(NFA_Fees[i]) if NFA_Fees[i] else 0.0
                            total_fee = exec_comm_val + fees_val + clr_comm_val + NFA_Fees_val
                            try:
                                #one line space between final month entrty and current entry
                                last_row = sheet.range('A' + str(sheet.cells.last_cell.row)).end('up').end('up').end('up').end('up').row
                                pre_date = sheet.range('A' + str(last_row)).value
                                pre_month = pre_date.strftime("%m")
                                curr_month = daybefore.strftime("%m")
                            except AttributeError:
                                last_row = sheet.range('A' + str(sheet.cells.last_cell.row)).end('up').end('up').end('up').end('up').end('up').end('up').row
                                pre_date = sheet.range('A' + str(last_row)).value
                                pre_month = pre_date.strftime("%m")
                                curr_month = daybefore.strftime("%m")
                            if (pre_month == curr_month):
                                bought_quantity[i] = bought_quantity[i].replace(" ","")
                                sheet.range(f"A{last_row+1}").api.EntireRow.Insert()
                                sheet.range(f"A{last_row+1}").value = Input_Date[i].strftime("%m-%d-%Y")
                                amount[i] = amount[i].replace(" ","")

                                if(trade_date[i] == '' and bought_quantity[i] == '' and sold_quantity[i] == ''):
                                    sheet.range(f"B{last_row+1}").value = 'COMMISSION ADJUSTMENTS'
                                    sheet.range(f"I{last_row+1}").value = amount[i]
                                
                                elif(bought_quantity[i] == ''):
                                            sheet.range(f"K{last_row+1}").value = desc[i]
                                            sheet.range(f"L{last_row+1}").value = '-' + sold_quantity[i]
                                            sheet.range(f"M{last_row+1}").value = sold_price[i]
                                            sheet.range(f"N{last_row+1}").value = total_fee
                                            sheet.range(f"W{last_row+1}").value = amount[i]
                                            sheet.range(f"AB{last_row+1}").formula = f'=AB{last_row-1}+I{last_row}+N{last_row}+W{last_row}+Y{last_row}' 

                                else:
                                            sheet.range(f"F{last_row+1}").value = desc[i]
                                            sheet.range(f"G{last_row+1}").value = bought_quantity[i]
                                            sheet.range(f"H{last_row+1}").value = bought_price[i]
                                            sheet.range(f"I{last_row+1}").value = total_fee
                                            sheet.range(f"W{last_row+1}").value = amount[i]
                                            sheet.range(f"AB{last_row+1}").formula = f'=AB{last_row-1}+I{last_row}+N{last_row}+W{last_row}+Y{last_row}'
                                
                            continue
        wb.save()
        logging.info('Save changes to workbook')
        wb.close()
        logging.info('Closed workbook')
                             
    except FileNotFoundError:
        return 
    except Exception as ex:
        logging.exception(f'Exception caught in download because: {ex}')
        raise Exception("No such file or directory")
def main():
    try:
        logging.info('process started')
        date_day_before= datetime.now() - timedelta(6)
        date_file = date_day_before.strftime("%m%d%Y")
        fname = date_day_before.strftime("%d%m") + 'F'
        log_json='[{"JOB_ID": "'+str(job_id)+'","CURRENT_DATETIME": "'+str(datetime.now())+'"}]'
        logging.info('Calling Macquarie_ov')
        MACQUARIE_OV(date_file,fname)
        logging.info(' Engine Disposed ---- end')
        bu_alerts.send_mail(
                    receiver_email = 'pakhi.laad@biourja.com',
                    mail_subject ='JOB SUCCESS - MACQUARIE OTHER_VERTICAL',
                    mail_body=f"Process completed successfully for all the Projects",
                    attachment_location = log_file_location
                )
                
    except Exception as e:
        logger.exception(f'Error occurred in . {e}')
        bu_alerts.send_mail(
                            receiver_email= 'pakhi.laad@biourja.com',
                            mail_subject=f"JOB FAILED ::  MACQUARIE OTHER_VERTICAL",
                            mail_body=f"{e}",
                            attachment_location = log_file_location)



if __name__ == '__main__':
    today_date = date.today()
    today = today_date.strftime("%m_%d_%Y")
    job_name = "Macquarie_OV"
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