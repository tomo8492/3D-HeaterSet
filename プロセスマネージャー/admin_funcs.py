import module as M

def sql_admin():
    import asyncio
    result = asyncio.run(M.update_lead_time('DB06105'))

if __name__ == '__main__':
    sql_admin()