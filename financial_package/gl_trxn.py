import sql_update
import sql_retrieve


def update1(date_end, dict_db):
    #Build GL Reporting Table
    sql = """
            USE [Playground]
            DECLARE @RC int            
            
            EXECUTE @RC = [myop\jason.walker].[proc_gl_account_reporting]
        """  
    sql_update.update1(a_driver=dict_db.get('driver'), 
                       a_server=dict_db.get('server_2'), 
                       a_db=dict_db.get('db_playground'), 
                       a_sql=sql)
    
    
    #Populate GL Balances
    sql = """
            USE [Playground]
            DECLARE @RC int
            DECLARE @dateend date
            
            EXECUTE @RC = [myop\jason.walker].[proc_gl_balance] 
               '{d1}'                
        """.format(d1=date_end)
        
    sql_update.update1(a_driver=dict_db.get('driver'), 
                       a_server=dict_db.get('server_2'), 
                       a_db=dict_db.get('db_playground'), 
                       a_sql=sql)

    #Validate Net Income
    sql = """
        declare @dateend date = '{d1}'
        declare @datestart date = dateadd("d",-day(@dateend)+1,@dateend)
        
        select
            q1.company,
            sum(q1.amount)
        from (
        
        select
            glb.company,
            sum(glb.amount) as amount --positive
        from Playground.[myop\jason.walker].gl_balance glb
            inner join Playground.[myop\jason.walker].gl_account_reporting glr
                on glb.gl_account = glr.gl_account
                and glb.company = glr.company
        where
            glb.period = @datestart
            and glr.financial_statement = 'P&L'
            and glr.level_1 is not null
            and glr.level_2 is not null
        group by
            glb.company
        
        union all
        
        select
            'HT' as company,
            -sum(gl.Amount) --negative
        from TNDCSQL03.NAVRep.dbo.[Hi Touch$G_L Entry] gl with(nolock)
        where
            gl.[G_L Account No_] between '30000' and '69999'
            and gl.[Posting Date] between @datestart and @dateend
        
        union all
        
        select
            'MYOP' as company,
            -sum(gl.Amount) --negative
        from TNDCSQL03.NAVRep.dbo.[MYOP$G_L Entry] gl with(nolock)
        where
            gl.[G_L Account No_] between '40000' and '79999'
            and gl.[Posting Date] between @datestart and @dateend
        
        union all
        
        select
            'RAC' as company,
            -sum(gl.Amount) --negative
        from TNDCSQL03.NAVRep.dbo.[Rentacrate$G_L Entry] gl with(nolock)
        where
            gl.[G_L Account No_] between '40000' and '79999'
            and gl.[Posting Date] between @datestart and @dateend
        
        ) q1
        group by
            q1.company
    """.format(d1=date_end)
    
    net_income_recon = sql_retrieve.getdata1(a_driver=dict_db.get('driver'), 
                                    a_server=dict_db.get('server_2'), 
                                    a_db=dict_db.get('db_playground'), 
                                    a_sql=sql)    
    
    return net_income_recon

def update2(date_end, dict_db):
    #Populate GL Balances
    sql = """
            USE [Playground]
            DECLARE @RC int
            DECLARE @dateend date
            
            EXECUTE @RC = [myop\jason.walker].[proc_gl_balance_bs] 
               '{d1}'                
        """.format(d1=date_end)
        
    sql_update.update1(a_driver=dict_db.get('driver'), 
                       a_server=dict_db.get('server_2'), 
                       a_db=dict_db.get('db_playground'), 
                       a_sql=sql)    
    
    #Validate Balance Sheet
    sql = """
    declare @dateend date = '{d1}'
    
    select
        q1.company,
        sum(q1.amount) as Amount
    from (
    
    select
        'GLBalance' as source,
        glb.company,
        glb.gl_account,
        -sum(glb.amount) as amount --negative
    from Playground.[myop\jason.walker].gl_balance_bs glb
        inner join Playground.[myop\jason.walker].gl_account_reporting glr
            on glb.gl_account = glr.gl_account
            and glb.company = glr.company
    where
        glb.period = @dateend
        and glr.financial_statement = 'BS'
        and glr.level_1 is not null
        and glr.level_2 is not null
    group by
        glb.company,
        glb.gl_account
    
    union all
    
    select
        'GLEntry' as source,
        'HT' as company,
        gl.[G_L Account No_],
        sum(gl.Amount) --positive
    from TNDCSQL03.NAVRep.dbo.[Hi Touch$G_L Entry] gl with(nolock)
    where
        gl.[G_L Account No_] between '10000' and '29999'
        and gl.[Posting Date] <= @dateend
    group by
        gl.[G_L Account No_]
    
    union all
    
    select
        'GLEntry' as source,
        'MYOP' as company,
        gl.[G_L Account No_],
        sum(gl.Amount) --positive
    from TNDCSQL03.NAVRep.dbo.[MYOP$G_L Entry] gl with(nolock)
    where
        gl.[G_L Account No_] between '10000' and '39999'
        and gl.[Posting Date] <= @dateend
    group by
        gl.[G_L Account No_]
    
    union all
    
    select
        'GLEntry' as source,
        'RAC' as company,
        gl.[G_L Account No_] collate Latin1_General_100_CS_AS,
        sum(gl.Amount) --positive
    from TNDCSQL03.NAVRep.dbo.[Rentacrate$G_L Entry] gl with(nolock)
    where
        gl.[G_L Account No_] between '10000' and '39999'
        and gl.[Posting Date] <= @dateend
    group by
        gl.[G_L Account No_]
    
    ) q1
    group by
        q1.company
    """.format(d1=date_end)
    
    
    balance_sheet_recon = sql_retrieve.getdata1(a_driver=dict_db.get('driver'), 
                                    a_server=dict_db.get('server_2'), 
                                    a_db=dict_db.get('db_playground'), 
                                    a_sql=sql)    
    
    
    return balance_sheet_recon
    
if __name__ == "__main__":
    update1()
    update2()
        