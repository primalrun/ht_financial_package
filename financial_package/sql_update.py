import pypyodbc


def update1(a_driver, a_server, a_db, a_sql):

    connection_string = '''Driver={d1};
                            Server={s1};
                            Database={db1};
                            Trusted_Connection=Yes;
                        '''.format(d1=a_driver, 
                                   s1=a_server, 
                                   db1=a_db)
                        
    connection = pypyodbc.connect(connection_string)
    cur = connection.cursor()    
    cur.execute(a_sql)
    cur.commit()
    cur.close()
    connection.close

def update2(a_driver, a_server, a_db, a_table_schema, a_table, a_list):

    connection_string = '''Driver={d1};
                            Server={s1};
                            Database={db1};
                            Trusted_Connection=Yes;
                        '''.format(d1=a_driver, 
                                   s1=a_server, 
                                   db1=a_db)
                        
    connection = pypyodbc.connect(connection_string)
    cur = connection.cursor()    
    cur.executemany("""insert into {db1}.{ts}.{t1} 
                    values (?, ?, ?, ?)""".format(
                        db1=a_db, ts=a_table_schema, 
                        t1=a_table), a_list)
    cur.commit()
    cur.close()
    connection.close
    
    
if __name__ == "__main__":
    update1()
    update2()