

import pyodbc
#15151385.uttcare.com,29075
#120.24.68.150,1453
outname=pyodbc.connect('DRIVER={SQL Server};SERVER=15151385.uttcare.com,29075;DATABASE=outweb;UID=sa;PWD=p0o9i8u7')
cursor2=outname.cursor()

name=cursor2.execute("select name from sysobjects where xtype='U'").fetchall()
k=[]
for i in name:
    k.append(i[0])

k.sort()
print k

for i in range(len(k)):
    sql="drop table [{0}]".format(k[i].encode("gbk"))
    cursor2.execute(sql)
    outname.commit()