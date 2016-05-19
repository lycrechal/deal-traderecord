

import pyodbc
#15151385.uttcare.com,29075
#120.24.68.150,1453
outname=pyodbc.connect('DRIVER={SQL Server};SERVER=120.24.68.150,1453,29075;DATABASE=outname;UID=dbUser;PWD=db+123-456')
cursor2=outname.cursor()

name=cursor2.execute("select name from sysobjects where xtype='U'").fetchall()
k=[]
for i in name:
    k.append(i[0])

k.sort()
print k

for i in range(139,len(k)):
    sql="drop table [{0}]".format(k[i].encode("gbk"))
    cursor2.execute(sql)
    outname.commit()