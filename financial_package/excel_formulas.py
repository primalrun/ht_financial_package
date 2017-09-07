from excel_column_number import dict_col_converter as dcc

def sum_row_1(row, col1, col2):
    formula1 = "=sum({col_letter1}{row1}:{col_letter2}{row1})".format(
        col_letter1=dcc.get(col1), row1=row, col_letter2=dcc.get(col2))
    return formula1

def sum_col_1(row1, row2, col1):
    formula1 = "=sum({col1}{row1}:{col1}{row2})".format(
        row1=row1, row2=row2, col1=dcc.get(col1))
    return formula1

def pct_of_total1(col1, r_total, r_row):
    formula1 = "=if({col1}{r_total}=0,0,{col1}{r_row}/{col1}{r_total})".format(
        col1=dcc.get(col1-1), r_total=r_total, r_row=r_row)
    return formula1

if __name__ == "__main__":
    sum_row_1()
    sum_col_1()
    pct_of_total1()
    