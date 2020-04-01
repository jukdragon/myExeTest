import pandas as pd
import pymssql

conn = pymssql.connect('tengfeng898.xicp.net:36489','sa', '89260037', 'TengFeng')
sql='''
    SELECT SA.仓库名称,SA.编号,SA.品名,SA.颜色,SA.销售条数,SA.销售数量,JIN.进货条数,VK_CunCp.条数 as 库存条数
    FROM
    (SELECT VI_Sale.仓库名称,VI_Sale.编号,VI_Sale.品名,VI_Sale.颜色, Sum(VI_Sale.条数) As 销售条数,Sum(VI_Sale.数量) As 销售数量
    FROM VI_Sale 
    where DATEDIFF(day,GETDATE(),VI_Sale.日期)>-1 
    group by VI_Sale.仓库名称,VI_Sale.编号,VI_Sale.品名,VI_Sale.颜色)SA
    left join
    (SELECT VI_INS.仓库名称,VI_INS.编号,VI_INS.品名,VI_INS.颜色, Sum(VI_INS.条数) As 进货条数
    FROM VI_INS
    where DATEDIFF(day,GETDATE(),VI_INS.日期)>-1 
    group by VI_INS.仓库名称,VI_INS.编号,VI_INS.品名,VI_INS.颜色)JIN
    on SA.仓库名称  = JIN.仓库名称 and SA.品名 = JIN.品名 and SA.颜色 = JIN.颜色
    left join VK_CunCp on SA.仓库名称  = VK_CunCp.仓库名称 and SA.品名 = VK_CunCp.品名 and SA.颜色 = VK_CunCp.颜色
    order by SA.编号,SA.销售数量 desc
    '''
df = pd.read_sql(sql, conn)
# df['wind']=(df.symbol+'.'+df.exchange.apply(lambda x :x[-2:]))



print(df)
df.to_excel('本日销售分析.xlsx', index=0)
print('ok')
conn.close()
