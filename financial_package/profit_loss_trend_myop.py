from datetime import datetime
import setup1
from class_repo import SQLExchange
from class_repo import ExcelWidget
from datetime import date
from dateutil.relativedelta import relativedelta
from excel_column_number import dict_col_converter as dcc

def gross_margin_data(dict_db, date_end):
    sql_obj = SQLExchange(driver=dict_db.get('driver'), 
                          server=dict_db.get('server_2'), 
                          db=dict_db.get('db_playground'))
    
    sql = """
    declare @dateend date = '{d1}'
    declare @datestart date = dateadd("d", -datepart("d", @dateend) + 1, @dateend)
    set @datestart = dateadd("m", -11, @datestart)
    
    select
        glb.period,
        glr.level_1,    
        sum(glb.amount * glr.report_multiplier) as Amount
    from Playground.[myop\jason.walker].gl_balance glb
        inner join Playground.[myop\jason.walker].gl_account_reporting glr
            on glb.company = glr.company
            and glb.gl_account = glr.gl_account
    where
        glr.level_1 in ('Sales', 'Cost')
        and glb.period between @datestart and @dateend
        and glb.company = 'MYOP'        
    group by
        glb.period,
        glr.level_1
    """.format(d1=date_end)    
    
    sql_result = sql_obj.sql_retrieve(sql)
    return sql_result

def profit_loss_data(dict_db, date_end):
    sql_obj = SQLExchange(driver=dict_db.get('driver'), 
                      server=dict_db.get('server_2'), 
                      db=dict_db.get('db_playground'))
        
    sql = """
    declare @dateend date = '{d1}'
    declare @datestart date = dateadd("d", -datepart("d", @dateend) + 1, @dateend)
    set @datestart = dateadd("m", -11, @datestart)
    
    select
        glb.period,
        glr.level_1,
        glr.level_2,    
        sum(glb.amount * glr.report_multiplier) as Amount
    from Playground.[myop\jason.walker].gl_balance glb
        inner join Playground.[myop\jason.walker].gl_account_reporting glr
            on glb.company = glr.company
            and glb.gl_account = glr.gl_account
    where
        glr.financial_statement = 'P&L'
        and glr.level_1 not in ('Sales', 'Cost')
        and glb.period between @datestart and @dateend
        and glr.level_1 is not null
        and glr.level_2 is not null
        and glb.company = 'MYOP'
    group by
        glb.period,
        glr.level_1,
        glr.level_2
    """.format(d1=date_end)        
    
    sql_result = sql_obj.sql_retrieve(sql)
    return sql_result


def build_report(dict_db, date_end, wb):
    sales_and_cost = gross_margin_data(dict_db, date_end)
    pl_trend_data = profit_loss_data(dict_db, date_end)
    wb_cur = wb    
    ws_cur = wb_cur.create_sheet('P&L_Trend_MYOP')
    ws_cur.cell(row=1, column=1,value='P&L Trend MYOP')
    ws_cur.cell(row=2, column=1,value='Income Statement')
    ws_cur.cell(row=3, 
                column=1,
                value='For the 12 Months Ending {d1}'.format(
                    d1 = datetime.strftime(date_end, '%B %d, %Y') ))
        
    for r in range(1, 4):        
        c1 = ws_cur.cell(row=r, column=1)
        wb_cur = ExcelWidget(wb_cur).arial_8_bold_center(c1)

        
    ws_cur.merge_cells(start_row=1, end_row=1, start_column=1, 
                       end_column = 14)
    ws_cur.merge_cells(start_row=2, end_row=2, start_column=1, 
                       end_column = 14)
    ws_cur.merge_cells(start_row=3, end_row=3, start_column=1, 
                       end_column = 14)        
    
    date_next = date_end + relativedelta(days=-date_end.day + 1)
    trend_period = []
    
    for i in range(0, 12):
        trend_period.append(date_next)
        date_next = date_next + relativedelta(months=-1) 
    
    trend_period.sort(key=None, reverse=False)
    row_next = 5    
    c = 2
    
    #Header
    for d in range(0, len(trend_period)):        
        c1 = ws_cur.cell(row=row_next, column=c,
                          value=datetime.strftime(trend_period[d], '%b %Y'))
        wb_cur = ExcelWidget(wb_cur).header_blue_arial_8(c1)        
        c += 1    
    
    c1 = ws_cur.cell(row=row_next, column=14, value='Total')
    wb_cur = ExcelWidget(wb_cur).header_blue_arial_8(c1)
    c_last = ws_cur.max_column
    
    #Gross Margin
    row_next += 1    
    c1 = ws_cur.cell(row=row_next, column=1, value='Total Revenue')
    wb_cur = ExcelWidget(wb_cur).arial_8_bold(c1)
    
    c = 2
    
    for d in range(0, len(trend_period)):
        c1 = ws_cur.cell(row=row_next, column=c, 
                         value = sum([x[2] for x in sales_and_cost 
                                  if x[0] == trend_period[d] 
                                  if x[1] == 'Sales']))
        wb_cur = ExcelWidget(wb_cur).arial_8_bold_number0(c1)
        c += 1
    
    
    c1 = ws_cur.cell(row=row_next, column=c)
    formula1 = ExcelWidget(wb_cur).row_total(2, c_last - 1, row_next)
    c1.value = formula1
    wb_cur = ExcelWidget(wb_cur).arial_8_bold_number0(c1)
    
    row_next += 1
    c1 = ws_cur.cell(row=row_next, column=1, value='Cost of Sales')
    wb_cur = ExcelWidget(wb_cur).arial_8_bold(c1)
    
    c = 2
    
    for d in range(0, len(trend_period)):
        c1 = ws_cur.cell(row=row_next, column=c, 
                         value = sum([x[2] for x in sales_and_cost 
                                  if x[0] == trend_period[d] 
                                  if x[1] == 'Cost']))
        wb_cur = ExcelWidget(wb_cur).arial_8_bold_number0(c1)
        c += 1
    
    c1 = ws_cur.cell(row=row_next, column=c)
    formula1 = ExcelWidget(wb_cur).row_total(2, c_last - 1, row_next)
    c1.value = formula1
    wb_cur = ExcelWidget(wb_cur).arial_8_bold_number0(c1)

    row_next += 1
    c1 = ws_cur.cell(row=row_next, column=1, value='Gross Margin')
    wb_cur = ExcelWidget(wb_cur).arial_8_bold(c1)    
    row_gm = row_next    
    
    for c in range(2, 15):
        c1 = ws_cur.cell(row=row_next, column=c, 
                         value = '={c1}{r1}-{c1}{r2}'.format(
                             c1=dcc.get(c), 
                             r1=row_next - 2, 
                             r2=row_next - 1))
        wb_cur = ExcelWidget(wb_cur).arial_8_bold_number0(c1)
        wb_cur = ExcelWidget(wb_cur).border_tb(c1)

    #Adj to Margin
    row_next += 2    
    c1 = ws_cur.cell(row=row_next, column=1, value='Adjustments to Margin:')
    wb_cur = ExcelWidget(wb_cur).arial_8_bold(c1)
    
    row_next += 1
    adjustment_to_margin = ['Rebates', 'Cash Discounts', 'Customer Discounts', 
                            'Cost of Goods Adjustments', 
                            'Other Adjustments to Margin']
    
    for x in adjustment_to_margin:
        c1 = ws_cur.cell(row=row_next, column=1, value=x)
        wb_cur = ExcelWidget(wb_cur).arial_8(c1)
        c = 2        
        for d in trend_period:
            amount = sum([p[3] for p in pl_trend_data
                      if p[0] == d
                      if p[1] == 'Adjustment to Margin'
                      if p[2] == x])
            c1 = ws_cur.cell(row=row_next, column=c, value=amount)
            wb_cur = ExcelWidget(wb_cur).arial_8_number0(c1)
            c += 1
        formula1 = ExcelWidget(wb_cur).row_total(2, c_last - 1, row_next)
        c1 = ws_cur.cell(row=row_next, column=c, value = formula1)
        wb_cur = ExcelWidget(wb_cur).arial_8_number0(c1)
        row_next +=1
        
    c1 = ws_cur.cell(row=row_next, column=1, 
                     value='Total Adjustments to Margin')
    wb_cur = ExcelWidget(wb_cur).arial_8_bold(c1)
    
    for c in range(2, c_last + 1):
        c1 = ws_cur.cell(row=row_next, column=c)
        formula1 = ExcelWidget(wb_cur).col_total(row_next - 
                                                 len(adjustment_to_margin), 
                                                 row_next - 1, c)
        c1.value=formula1
        wb_cur = ExcelWidget(wb_cur).arial_8_bold_number0(c1)
        wb_cur = ExcelWidget(wb_cur).border_tb(c1)
    
    #Adjusted Gross Margin
    row_next += 1
    row_agm = row_next
    c1 = ws_cur.cell(row=row_next, column=1, value='Adjusted Gross Margin')
    wb_cur = ExcelWidget(wb_cur).arial_8_bold(c1)
    
    for c in range(2, c_last + 1):
        c1 = ws_cur.cell(row=row_next, column=c)
        cs1 = ws_cur.cell(row=row_gm, column=c)
        cs2 = ws_cur.cell(row=row_next - 1, column=c)
        formula1 = ExcelWidget(wb_cur).sum_2(cs1, cs2)
        c1.value=formula1
        wb_cur = ExcelWidget(wb_cur).arial_8_bold_number0(c1)
        wb_cur = ExcelWidget(wb_cur).border_tb(c1)        
    
    #Commission
    row_next += 2
    c1 = ws_cur.cell(row=row_next, column=1, value='Commissions')
    wb_cur = ExcelWidget(wb_cur).arial_8_bold(c1)
    commission = ['Commissions']
    
    for x in commission:
        c1 = ws_cur.cell(row=row_next, column=1, value=x)
        wb_cur = ExcelWidget(wb_cur).arial_8(c1)
        c = 2        
        for d in trend_period:
            amount = sum([p[3] for p in pl_trend_data
                      if p[0] == d
                      if p[1] == 'Commissions'
                      if p[2] == x])
            c1 = ws_cur.cell(row=row_next, column=c, value=amount)
            wb_cur = ExcelWidget(wb_cur).arial_8_number0(c1)
            c += 1
        formula1 = ExcelWidget(wb_cur).row_total(2, c_last - 1, row_next)
        c1 = ws_cur.cell(row=row_next, column=c, value = formula1)
        wb_cur = ExcelWidget(wb_cur).arial_8_number0(c1)
        row_next +=1
        
    c1 = ws_cur.cell(row=row_next, column=1, 
                     value='Total Commissions')
    wb_cur = ExcelWidget(wb_cur).arial_8_bold(c1)
    row_commission = row_next
    
    for c in range(2, c_last + 1):
        c1 = ws_cur.cell(row=row_next, column=c)
        formula1 = ExcelWidget(wb_cur).col_total(row_next - 1, 
                                                 row_next - 1, c)
        c1.value=formula1
        wb_cur = ExcelWidget(wb_cur).arial_8_bold_number0(c1)
        wb_cur = ExcelWidget(wb_cur).border_tb(c1)
    
    row_next += 1
    c1 = ws_cur.cell(row=row_next, column=1, value='% To Margin')
    wb_cur = ExcelWidget(wb_cur).arial_8_italic(c1)
    
    for c in range(2, c_last + 1):
        c1 = ws_cur.cell(row=row_next, column=c)
        cdivisor = ws_cur.cell(row=row_next - 1, column=c)
        cdividend = ws_cur.cell(row=row_gm, column=c)
        formula1 = ExcelWidget(wb_cur).divide_2(cdivisor, cdividend)
        c1.value = formula1
        wb_cur = ExcelWidget(wb_cur).arial_8_italic_percent1(c1)
    
    
    #Other Personnel Expense
    row_next += 2    
    c1 = ws_cur.cell(row=row_next, column=1, value='Other Personnel Expenses:')
    wb_cur = ExcelWidget(wb_cur).arial_8_bold(c1)
    
    row_next += 1
    personnel_exp = ['Salaries',
                     'Bonuses',
                     'Benefits',
                     'Payroll Taxes']
    
    for x in personnel_exp:        
        c1 = ws_cur.cell(row=row_next, column=1, value=x)
        wb_cur = ExcelWidget(wb_cur).arial_8(c1)
        c = 2        
        for d in trend_period:
            amount = sum([p[3] for p in pl_trend_data
                      if p[0] == d
                      if p[1] == 'Other Personnel Expense'
                      if p[2] == x])
            c1 = ws_cur.cell(row=row_next, column=c, value=amount)
            wb_cur = ExcelWidget(wb_cur).arial_8_number0(c1)
            c += 1
        formula1 = ExcelWidget(wb_cur).row_total(2, c_last - 1, row_next)
        c1 = ws_cur.cell(row=row_next, column=c, value = formula1)
        wb_cur = ExcelWidget(wb_cur).arial_8_number0(c1)
        row_next +=1
        
    c1 = ws_cur.cell(row=row_next, column=1, 
                     value='Total Other Personnel Expenses')
    wb_cur = ExcelWidget(wb_cur).arial_8_bold(c1)
    row_oth_pers_exp = row_next
    
    for c in range(2, c_last + 1):
        c1 = ws_cur.cell(row=row_next, column=c)
        formula1 = ExcelWidget(wb_cur).col_total(row_next - len(personnel_exp), 
                                                 row_next - 1, c)
        c1.value=formula1
        wb_cur = ExcelWidget(wb_cur).arial_8_bold_number0(c1)
        wb_cur = ExcelWidget(wb_cur).border_tb(c1)
        
    #Other Operating Expense
    row_next += 2    
    c1 = ws_cur.cell(row=row_next, column=1, value='Other Operating Expenses:')
    wb_cur = ExcelWidget(wb_cur).arial_8_bold(c1)
    
    row_next += 1
    opex = ['Travel',
                     'Meals & Entertainment',
                     'Facility Rent',
                     'Utilities',
                     'Facility Maintenance',
                     'Fleet Fuel',
                     'Fleet Repair & Maintenance',
                     'Vehicle Rent Expense',
                     'Outsourced Delivery',
                     'Outbound Freight',
                     'Outbound Freight Rebates',
                     'Equipment Rental',
                     'MRO Supplies',
                     'Office Freight & Postage',
                     'Outside Services',
                     'IT Services & Maintenance',
                     'Telecom Expense',
                     'Business Insurance',
                     'Training, Dues & Subscriptions',
                     'Advertising & Marketing',
                     'Diversity Partner Fees',
                     'e-Commerce Fees',
                     'Professional Fees',
                     'Bank Charges',
                     'Bad Debt Expense',
                     'Other Taxes',
                     'Other Expenses']
    
    for x in opex:        
        c1 = ws_cur.cell(row=row_next, column=1, value=x)
        wb_cur = ExcelWidget(wb_cur).arial_8(c1)
        c = 2        
        for d in trend_period:
            amount = sum([p[3] for p in pl_trend_data
                      if p[0] == d
                      if p[1] == 'Other Operating Expense'
                      if p[2] == x])
            c1 = ws_cur.cell(row=row_next, column=c, value=amount)
            wb_cur = ExcelWidget(wb_cur).arial_8_number0(c1)
            c += 1
        formula1 = ExcelWidget(wb_cur).row_total(2, c_last - 1, row_next)
        c1 = ws_cur.cell(row=row_next, column=c, value = formula1)
        wb_cur = ExcelWidget(wb_cur).arial_8_number0(c1)
        row_next +=1
        
    c1 = ws_cur.cell(row=row_next, column=1, 
                     value='Total Other Operating Expenses')
    wb_cur = ExcelWidget(wb_cur).arial_8_bold(c1)
    
    for c in range(2, c_last + 1):
        c1 = ws_cur.cell(row=row_next, column=c)
        formula1 = ExcelWidget(wb_cur).col_total(row_next - len(opex), 
                                                 row_next - 1, c)
        c1.value=formula1
        wb_cur = ExcelWidget(wb_cur).arial_8_bold_number0(c1)
        wb_cur = ExcelWidget(wb_cur).border_tb(c1)        
        
            
    #Total Opex
    row_next += 1
    c1 = ws_cur.cell(row=row_next, column=1, value='Total Operating Expenses')
    wb_cur = ExcelWidget(wb_cur).arial_8_bold(c1)
    row_total_opex = row_next
    
    for c in range(2, c_last + 1):
        c1 = ws_cur.cell(row=row_next, column=c)
        cs1 = ws_cur.cell(row=row_commission, column=c)
        cs2 = ws_cur.cell(row=row_oth_pers_exp, column=c)
        cs3 = ws_cur.cell(row=row_next - 1, column=c)
        formula1 = ExcelWidget(wb_cur).sum_3(cs1, cs2, cs3)
        c1.value=formula1
        wb_cur = ExcelWidget(wb_cur).arial_8_bold_number0(c1)
        wb_cur = ExcelWidget(wb_cur).border_tb(c1)    


    #Other Income/(Expense)
    row_next += 2    
    c1 = ws_cur.cell(row=row_next, column=1, value='Other Income / (Expense):')
    wb_cur = ExcelWidget(wb_cur).arial_8_bold(c1)
    
    row_next += 1
    oth_inc_exp = ['Corporate Allocation',
                     'Deferred Partner Revenue',
                     'Excluded Expenses',
                     'Misc. Income',
                     'Gain / (Loss) Fixed Asset Disposal']
    
    for x in oth_inc_exp:        
        c1 = ws_cur.cell(row=row_next, column=1, value=x)
        wb_cur = ExcelWidget(wb_cur).arial_8(c1)
        c = 2        
        for d in trend_period:
            amount = sum([p[3] for p in pl_trend_data
                      if p[0] == d
                      if p[1] == 'Other Income / (Expense)'
                      if p[2] == x])
            c1 = ws_cur.cell(row=row_next, column=c, value=amount)
            wb_cur = ExcelWidget(wb_cur).arial_8_number0(c1)
            c += 1
        formula1 = ExcelWidget(wb_cur).row_total(2, c_last - 1, row_next)
        c1 = ws_cur.cell(row=row_next, column=c, value = formula1)
        wb_cur = ExcelWidget(wb_cur).arial_8_number0(c1)
        row_next +=1
        
    c1 = ws_cur.cell(row=row_next, column=1, 
                     value='Total Other Income / (Expense)')
    wb_cur = ExcelWidget(wb_cur).arial_8_bold(c1)
    
    for c in range(2, c_last + 1):
        c1 = ws_cur.cell(row=row_next, column=c)
        formula1 = ExcelWidget(wb_cur).col_total(row_next - len(oth_inc_exp), 
                                                 row_next - 1, c)
        c1.value=formula1
        wb_cur = ExcelWidget(wb_cur).arial_8_bold_number0(c1)
        wb_cur = ExcelWidget(wb_cur).border_tb(c1)

    #EBITDA
    row_next += 2
    c1 = ws_cur.cell(row=row_next, column=1, value='EBITDA')
    wb_cur = ExcelWidget(wb_cur).arial_8_bold(c1)
    row_ebitda = row_next
    
    for c in range(2, c_last + 1):
        c1 = ws_cur.cell(row=row_next, column=c)
        c_agm = ws_cur.cell(row=row_agm, column=c)
        c_opex = ws_cur.cell(row=row_total_opex, column=c)
        c_oth_income = ws_cur.cell(row=row_next - 2, column=c)
        formula1 = ExcelWidget(wb_cur).ebitda_calc(c_agm,
                                                   c_opex,
                                                   c_oth_income)
        c1.value = formula1
        wb_cur = ExcelWidget(wb_cur).arial_8_yellow_bold_number0(c1)
        wb_cur = ExcelWidget(wb_cur).border_tb(c1)
        
    #IDA
    row_next += 2
    ida = ['Interest Expense',
           'Depreciation',
           'Amortization'
          ]
    
    for x in ida:        
        c1 = ws_cur.cell(row=row_next, column=1, value=x)
        wb_cur = ExcelWidget(wb_cur).arial_8(c1)
        c = 2        
        for d in trend_period:
            amount = sum([p[3] for p in pl_trend_data
                      if p[0] == d
                      if p[1] == 'ITDA'
                      if p[2] == x])
            c1 = ws_cur.cell(row=row_next, column=c, value=amount)
            wb_cur = ExcelWidget(wb_cur).arial_8_number0(c1)
            c += 1
        formula1 = ExcelWidget(wb_cur).row_total(2, c_last - 1, row_next)
        c1 = ws_cur.cell(row=row_next, column=c, value = formula1)
        wb_cur = ExcelWidget(wb_cur).arial_8_number0(c1)
        row_next +=1
        
    c1 = ws_cur.cell(row=row_next, column=1, 
                     value='Net Income Before Taxes')
    wb_cur = ExcelWidget(wb_cur).arial_8_bold(c1)
    
    for c in range(2, c_last + 1):
        c1 = ws_cur.cell(row=row_next, column=c)
        ebitda = ws_cur.cell(row=row_ebitda, column=c)        
        ida_start = ws_cur.cell(row=row_next - len(ida),
                                column=c)
        ida_end = ws_cur.cell(row=row_next - 1, column=c)
        formula1 = ExcelWidget(wb_cur).net_income_before_tax(ebitda, 
                                                             ida_start, 
                                                             ida_end)
        c1.value=formula1
        wb_cur = ExcelWidget(wb_cur).arial_8_bold_number0(c1)
        wb_cur = ExcelWidget(wb_cur).border_tb(c1)    
            

    row_next += 2
    tax = ['State Income Taxes']

    for x in tax:        
        c1 = ws_cur.cell(row=row_next, column=1, value=x)
        wb_cur = ExcelWidget(wb_cur).arial_8(c1)
        c = 2        
        for d in trend_period:
            amount = sum([p[3] for p in pl_trend_data
                      if p[0] == d
                      if p[1] == 'ITDA'
                      if p[2] == x])
            c1 = ws_cur.cell(row=row_next, column=c, value=amount)
            wb_cur = ExcelWidget(wb_cur).arial_8_number0(c1)
            c += 1
        formula1 = ExcelWidget(wb_cur).row_total(2, c_last - 1, row_next)
        c1 = ws_cur.cell(row=row_next, column=c, value = formula1)
        wb_cur = ExcelWidget(wb_cur).arial_8_number0(c1)
    row_next +=1
    
    
    #Net Income
    c1 = ws_cur.cell(row=row_next, column=1, value='Net Income')
    wb_cur = ExcelWidget(wb_cur).arial_8_bold(c1)
    
    for c in range(2, c_last + 1):
        c1 = ws_cur.cell(row=row_next, column=c)
        c_nibt = ws_cur.cell(row=row_next - 3, column=c)
        c_tax = ws_cur.cell(row=row_next - 1, column=c) 
        formula1 = ExcelWidget(wb_cur).subtract_2(c_nibt, c_tax)
        c1.value=formula1
        wb_cur = ExcelWidget(wb_cur).arial_8_bold_number0(c1)
        wb_cur = ExcelWidget(wb_cur).border_tb2(c1)
    
    #Excluded Expenses
    row_next += 2
    c1 = ws_cur.cell(row=row_next, column=1, value='EXCLUDED EXPENSES')
    wb_cur = ExcelWidget(wb_cur).arial_8_bold(c1)
        
    for r in range(row_next, row_next - 30, -1):
        if ws_cur.cell(row=r, column=1).value == 'Excluded Expenses':
            row_excluded_expense = r
            break
        
    for c in range(2, c_last + 1):
        c1 = ws_cur.cell(row=row_next, column=c)
        c_excluded_expense = ws_cur.cell(row=row_excluded_expense, column=c)
        formula1 = (ExcelWidget(wb_cur).
                    excluded_expense_inverse(c_excluded_expense))
        c1.value = formula1
        wb_cur = ExcelWidget(wb_cur).arial_8_yellow_bold_number0(c1)
        wb_cur = ExcelWidget(wb_cur).border_tb(c1)
        
    #Adjusted EBITDA
    row_next += 2
    c1 = ws_cur.cell(row=row_next, column=1, value='ADJUSTED EBITDA')
    wb_cur = ExcelWidget(wb_cur).arial_8_bold(c1)        
        
    for c in range(2, c_last + 1):
        c1 = ws_cur.cell(row=row_next, column=c)
        c_ebitda = ws_cur.cell(row=row_ebitda, column=c)
        c_excl_exp = ws_cur.cell(row=row_next - 2, column=c)
        formula1 = ExcelWidget(wb_cur).sum_2(c_ebitda, c_excl_exp)
        c1.value = formula1
        wb_cur = ExcelWidget(wb_cur).arial_8_yellow_bold_number0(c1)
        wb_cur = ExcelWidget(wb_cur).border_tb(c1)
        
                
    #Row height
    for r in range(1, ws_cur.max_row + 1):
        ws_cur.row_dimensions[r].height = 11.5

    #Column Width
    ws_cur.column_dimensions['A'].width = 27
    
    for c in range(2, 14):
        ws_cur.column_dimensions[dcc.get(c)].width = 10
        
    ws_cur.column_dimensions['N'].width = 10.5
            
    #Page Setup
    wb_cur = ExcelWidget(wb_cur).page_setup_portrait_1X1(ws_cur, 5)    
    
    #Freeze Panes
    c1 = ws_cur.cell(row=6, column=1)
    ws_cur.freeze_panes = c1
    
    return wb_cur

if __name__ == '__main__':
    build_report()
    gross_margin_data()

