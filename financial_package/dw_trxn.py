import sql_update



def update1(date_end, dict_db):
    sql = """
            USE [Playground]
            DECLARE @RC int
            DECLARE @dateend date
            
            EXECUTE @RC = [myop\jason.walker].[proc_office_products_prod_class] 
               '{d1}'                
        """.format(d1=date_end)
        
    sql_update.update1(a_driver=dict_db.get('driver'), 
                       a_server=dict_db.get('server_2'), 
                       a_db=dict_db.get('db_playground'), 
                       a_sql=sql)

    
if __name__ == "__main__":
    update1()

        