from openpyxl import Workbook
import calendar
import sys
import setup1
import ebitda_cash_analysis_all as ecaa
import gl_trxn
import dw_trxn
import schedule_b
import consolidated_p_and_l as cpl
import profit_loss_ht_myop as plhm
import profit_loss_rac as plr
import profit_loss_ht as plh
import profit_loss_myop as plm
import location_inventory
import balance_sheet_consolidated
import balance_sheet_ht_myop
import balance_sheet_ht
import balance_sheet_myop
import balance_sheet_rac
import schedule_c
import profit_loss_trend_ht_my_rac
import profit_loss_trend_ht_my
import profit_loss_trend_ht
import profit_loss_trend_myop
import profit_loss_trend_rac

#Excel Book
wb = Workbook()

#Setup data
cont_p = setup1.setup_f1()
date_start = cont_p[0]
path_str = cont_p[1]
date_end = cont_p[2]
dict_db = setup1.db_dictionary()
month_actual = [x for x in range(1, date_end.month + 1)]
month_budget = [x for x in range(len(month_actual) + 1, 13)]
date_ytd_start = setup1.date_ytd_start(date_start=date_start)
date_trail_12_start = setup1.date_trail_12_start(date_start=date_start)
fin_package_input_directory = "Z:/Accounting/Accounting/Financial Package/"

#Cover Page
wb = setup1.get_cover_page(wb, fin_package_input_directory, date_end)

#Location Inventory
cont_location = location_inventory.location_inventory_update(wb=wb)

#check for error

if len(cont_location) > 1:
    sys.exit()
   
wb = cont_location[0]


#Populate GL Transactions in data warehouse
#Income Statement
net_income_recon = gl_trxn.update1(date_end=date_end, dict_db=dict_db)
    
net_income_variance = sum([x[1] for x in net_income_recon])
print('Net Income Variance =',int(net_income_variance))
     
if abs(net_income_variance) > 0.1:
    print('Net Income Variance Error')
    print(*net_income_recon, sep='\n')
    sys.exit()

#Balance Sheet
balance_sheet_recon = gl_trxn.update2(date_end=date_end, dict_db=dict_db)
      
balance_sheet_variance = sum([x[1] for x in balance_sheet_recon])
print('Balance Sheet Variance =',int(balance_sheet_variance))
       
if abs(balance_sheet_variance) > 0.1:
    print('Balance Sheet Variance Error')
    print(*balance_sheet_recon, sep='\n')
    sys.exit()


#----EBITDA Cash Analysis----
wb = ecaa.excelupdate(Workbook=wb, dict_db=dict_db, date_end=date_end, 
                      month_actual=month_actual, month_budget=month_budget)

#Schedule B
#Populate Office Products Product Class in data warehouse
dw_trxn.update1(date_end=date_end, dict_db=dict_db)
     
wb = schedule_b.excelupdate(wb=wb, 
                            dict_db=dict_db, 
                            date_start=date_start, 
                            date_end=date_end, 
                            date_ytd_start = date_ytd_start,
                            date_trail_12_start=date_trail_12_start)

#Consolidated P&L
wb = cpl.build_consolidated_p_and_l(wb, date_start, date_ytd_start, 
                                    dict_db, date_end)

#P&L HT and MYOP
wb = plhm.build_consolidated_p_and_l(wb, date_start, date_ytd_start, 
                                    dict_db, date_end)

#P&L HT
wb = plh.build_consolidated_p_and_l(wb, date_start, date_ytd_start, 
                                    dict_db, date_end)

#P&L MYOP
wb = plm.build_consolidated_p_and_l(wb, date_start, date_ytd_start, 
                                    dict_db, date_end)

#P&L RAC
wb = plr.build_consolidated_p_and_l(wb, date_start, date_ytd_start, 
                                    dict_db, date_end)

#P&L Consolidated Trend
wb = profit_loss_trend_ht_my_rac.build_report(dict_db, date_end, wb)
wb = profit_loss_trend_ht_my.build_report(dict_db, date_end, wb)
wb = profit_loss_trend_ht.build_report(dict_db, date_end, wb)
wb = profit_loss_trend_myop.build_report(dict_db, date_end, wb)
wb = profit_loss_trend_rac.build_report(dict_db, date_end, wb)

#Balance Sheet Consolidated
wb = balance_sheet_consolidated.build_consolidated_bs(wb, dict_db, date_end)

#Balance Sheet by Company
wb = balance_sheet_ht_myop.build_consolidated_bs(wb, dict_db, date_end)
wb = balance_sheet_ht.build_consolidated_bs(wb, dict_db, date_end)
wb = balance_sheet_myop.build_consolidated_bs(wb, dict_db, date_end)
wb = balance_sheet_rac.build_consolidated_bs(wb, dict_db, date_end)

#Schedule C Rebate Components
wb = schedule_c.build_report(dict_db, date_end, wb)

#Cleanup report
ws_del = wb.get_sheet_by_name('Sheet')
wb.remove_sheet(ws_del)
wb.save(path_str + 'Financial Package - {dstr} {d1:%y}.xlsx'.format(
        d1=date_start, dstr=calendar.month_name[date_start.month].upper()))

print('Financial Package Successfully Completed')





 





