import pyodbc

outname=pyodbc.connect('DRIVER={SQL Server};SERVER=120.24.68.150,1453;DATABASE=outweb;UID=dbUser;PWD=db+123-456')
cursor2=outname.cursor()

name=cursor2.execute("select name from sysobjects where xtype='U'").fetchall()
k=[]
for i in name:
    k.append(i[0])

k.sort()
print k
for i in range(len(k)):
    sql="drop table [{0}]".format(k[i])
    cursor2.execute(sql)
    outname.commit()