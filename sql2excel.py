import pandas as pd
import pymssql

def translate(value:str):
    # if (value.index('&') == -1):
    if (value.find('&') == -1):
        # 没有匹配到&
        return ""
    else:
        s = value.split('&')
        gz = s[0]
        # if (gz.index('~') == -1):
        if (gz.find('~') == -1):
            # 没有匹配到~
            return gz
        else:
            # 有~,要切分
            gzs = gz.split('~')
            gzl = gzs[0]
            return gzl

def tran_pinming(code, name:str):
    # 去除多余空格
    name = name.strip()
    if (code == '222'):
        return name
    else:
        name = '$'+name
        return name


conn = pymssql.connect('tengfeng898.xicp.net:36489','sa', '89260037', 'TengFeng')
sql='''
    SELECT 编号, 品名, 颜色, 条数, Y.HpsBz as 定位 
    FROM VI_Ins as V inner join B_Yszl as Y 
    on V.颜色 = Y.HpsName and V.产品ID = Y.HpID
    where V.仓库名称 ='临时仓' and datediff(dd,V.日期,GETDATE())=0
    order by 编号,颜色
    '''
df = pd.read_sql(sql, conn)
# df['wind']=(df.symbol+'.'+df.exchange.apply(lambda x :x[-2:]))

# 加标识符
# df["then"] = pd.np.where(df.编号 == '222', '', '$')
df['品名'] = df.apply(lambda row: tran_pinming(row['编号'], row['品名']), axis=1)

# 改定位
# df['广州'] = df.apply(lambda row: translate(row['定位']), axis=1)
df['定位'] = df.apply(lambda row: translate(row['定位']), axis=1)

print(df)
df.to_excel('进货信息.xlsx', index=0)
print('ok')
conn.close()
