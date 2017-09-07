
from class_repo import SQLExchange
import setup1

dict_db = setup1.db_dictionary()
sql_obj = SQLExchange(driver=dict_db.get('driver'), 
                       server=dict_db.get('server_2'), 
                       db=dict_db.get('db_playground')) 
                       
connection_string = sql_obj.connection_string()
sql = """ 
    drop table playground.[myop\jason.walker].grid1
    """
    
sql_obj.sql_no_retrieve(sql)

print('Complete')
