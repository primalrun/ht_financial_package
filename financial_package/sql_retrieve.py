import pypyodbc


def getdata1(a_driver, a_server, a_db, a_sql):

    connection_string = '''Driver={d1};
                            Server={s1};
                            Database={db1};
                            Trusted_Connection=Yes;
                        '''.format(d1=a_driver, 
                                   s1=a_server, 
                                   db1=a_db)
                        
    connection = pypyodbc.connect(connection_string)
    cur = connection.cursor()
    
    result1 = cur.execute(a_sql).fetchall()
    cur.close()
    connection.close    
    return result1


    
    

if __name__ == "__main__":
    getdata1()
    
    