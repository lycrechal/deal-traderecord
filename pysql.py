#!/usr/bin/env python
# -*- coding: utf-8 -*-

import xlrd
import xlwt
from xlrd import open_workbook
from xlutils.copy import copy
from xlwt import easyxf
import os
import time
import datetime
import pyodbc
multiple=int(raw_input("Y(^o^)Yenter the times "))
poundage=int(raw_input("Y(^o^)Yenter the poundage "))
tr = pyodbc.connect('DRIVER={SQL Server};SERVER=120.234.10.34,34333;DATABASE=TradeRecord;UID=wyq;PWD=321#@!cba')
cursor1 = tr.cursor()
od = pyodbc.connect('DRIVER={SQL Server};SERVER=120.234.10.34,34333;DATABASE=outdate;UID=wyq;PWD=321#@!cba')
cursor2 = od.cursor()
out = pyodbc.connect('DRIVER={SQL Server};SERVER=120.234.10.34,34333;DATABASE=out;UID=wyq;PWD=321#@!cba')
cursor3 = out.cursor()
op = pyodbc.connect('DRIVER={SQL Server};SERVER=120.234.10.34,34333;DATABASE=outperson;UID=wyq;PWD=321#@!cba')
cursor4 = op.cursor()
mk = pyodbc.connect('DRIVER={SQL Server};SERVER=120.234.10.34,34333;DATABASE=marketdata;UID=wyq;PWD=321#@!cba')
cursor5 = mk.cursor()
sql="drop table alldate,allname"
cursor1.execute(sql)
tr.commit()
class data():
    def rdata(self):


        row=cursor1.execute("select count(*) from mock").fetchall()
        date=cursor1.execute("select date from mock").fetchall()

        sql="""
            create table alldate(
            direction varchar(50),
            offsetflag varchar(50),
            price decimal(20, 4),
            number int,
            timedate varchar(50))
            """
        cursor1.execute(sql)
        tr.commit()


        for i in range(row[0][0]):
            mylist.append(date[i][0])



        
        cursor1.execute("insert into alldate(direction,offsetflag,price,number,timedate)select direction,offsetflag,price,number,date+' '+time from mock")
        tr.commit()


    def deal(self):

        sql="""
            create table {0}(
            direction varchar(50),
            offsetflag varchar(50),
            price decimal(20, 4),
            number int,
            timedate varchar(50))
            """.format('d_'+strfilename)
        cursor2.execute(sql)
        od.commit()
        sql="""
            insert into {0} (
            direction,offsetflag,price,number,timedate)
            select direction,offsetflag,price,number,date+' '+time
            from [TradeRecord].[dbo].[mock] where date='{1}'
            """.format('d_'+strfilename,strfilename)
        cursor2.execute(sql)
        od.commit()


if __name__=='__main__':


    mylist=[]
    t=data()
    t.rdata()
    newlist=sorted(list(set(mylist)))
    print newlist





    allsheet=[]

    calount=cursor2.execute("select count(*) from sysobjects where xtype='U'").fetchone()
    print calount[0]
    sheetname=cursor2.execute("select name from sysobjects where xtype='U'").fetchall()
    for i in range(calount[0]):
        allsheet.append(sheetname[i][0].split('_')[1])
    for i in range(len(newlist)):
        if newlist[i] not in allsheet:
            strfilename=newlist[i]
            t.deal()


class TradeRecord(object):
    def __init__(self,time,offsetFlag,direction,price,number):
        self.time=str(time)
        self.offsetFlag = offsetFlag
        self.direction = direction
        self.price = price
        self.number = number
class OutputTradeRecord(object):
    def __init__(self,change,openinterest,avgprice,avgprofit,pingprofit,amount,duokong,poundage,finalfinal):
        self.change=change
        self.openinterest=openinterest
        self.avgprice=avgprice
        self.avgprofit=avgprofit
        self.pingprofit=pingprofit
        self.amount=amount
        self.duokong=duokong
        self.poundage=poundage
        self.finalfinal=finalfinal
class outnet(object):
    def __init__(self,direction,offsetFlag,price,number,timedate):
        self.direction=direction
        self.offsetFlag=offsetFlag
        self.price=price
        self.number=number
        self.timedate=timedate
class OutputtradeRecord(object):
    def __init__(self,change,amount,avprice,position,closeprofit,movement):
        self.change=change
        self.amount=amount
        self.avprice=avprice
        self.position=position
        self.closeprofit=closeprofit
        self.movement=movement


class Methods(object):


    def readExcel(self,strfilename):
        self.OrginDate=[]
        self.daylist=[]
        self.rows=cursor2.execute("select count(*) from {0}".format('d_'+strfilename)).fetchone()
        self.data1=cursor2.execute("select direction, offsetflag, price, number, timedate from {0}".format('d_'+strfilename)).fetchall()

        print self.rows[0]
        self.rows=self.rows[0]
        for i in range(self.rows):
            _record = TradeRecord(self.data1[i][4],self.data1[i][1],self.data1[i][0],self.data1[i][2],self.data1[i][3])
            self.OrginDate.append(_record)
            self.daylist.append(self.data1[i][4].split(' ')[0])
            self.daylist=list(set(self.daylist))


    def readExcel2(self,strfilename):



        self.rows2=cursor5.execute("select count(*) from min").fetchone()
        self.rows2=self.rows2[0]
        self.data=cursor5.execute("select timedate, price from min").fetchall()
        for i in range(self.rows2):
            if self.data[i][0].split(' ')[0] in self.daylist:
                if self.data[i][0].split(' ')[0]==strfilename:

                    if int(self.data[i][0].split(' ')[1].split(':')[0])<17:
                        if int(self.data[i][0].split(' ')[1].split(':')[0])==9:
                            if int(self.data[i][0].split(' ')[1].split(':')[1])>14:
                                _record = TradeRecord(self.data[i][0].split('\x00')[0],'','',self.data[i][1],'')
                                self.OrginDate.append(_record)
                        else:
                            _record = TradeRecord(self.data[i][0].split('\x00')[0],'','',self.data[i][1],'')
                            self.OrginDate.append(_record)

        self.OrginDate.sort(key=lambda x:x.time.split(':'))

    def readExcel3(self,strfilename):


        self.rows3=cursor1.execute("select count(*) from actual").fetchone()
        self.rows3=self.rows3[0]
        self.data=cursor1.execute("select price, time, date from actual").fetchall()
        for i in range(self.rows3):
            if str(int(self.data[i][2]))==strfilename:
                self.OrginDate.append(TradeRecord(str(int(self.data[i][2]))+' '+str(self.data[i][1]),'','',self.data[i][0],''))
        self.OrginDate.sort(key=lambda x:x.time.split(':'))
        self.finalrows=len(self.OrginDate)
        print self.finalrows


    def deal(self):
        a={}
        for i in range(self.finalrows):
            if self.OrginDate[i].direction == u'买'.encode("gbk"):
                a[i]=1
            else:
                a[i]=-1
        b={}
        for i in range(self.finalrows):
            if self.OrginDate[i].offsetFlag== u'开仓'.encode("gbk"):
                b[i]=1
            else:
                b[i]=-1

        b1={}
        for i in range(self.finalrows):
            if self.OrginDate[i].direction==u''.encode("gbk"):
                b1[i]=1
            else:
                if self.OrginDate[i].offsetFlag == u'开仓'.encode("gbk"):
                    b1[i]=1
                else:
                    b1[i]=-1
        self.n={}
        if self.OrginDate[0].direction==u''.encode("gbk"):

            self.n[0]=1
        else:
            self.n[0]=a[0]*b[0]
        for i in range(1,self.finalrows):
            if self.OrginDate[i].direction==u''.encode("gbk"):
                self.n[i]=self.n[i-1]
            else:
                self.n[i]=a[i]*b[i]

        abc={}
        for i in range(self.finalrows):
            if self.OrginDate[i].direction==u''.encode("gbk"):
                abc[i]=0
            else:
                abc[i]=self.OrginDate[i].number

        self.c={}
        for i in range(self.finalrows):
            if self.OrginDate[i].direction == u''.encode("gbk"):
                self.c[i]=0
            else:
                self.c[i]=b[i]*self.OrginDate[i].number

        self.pouno={}
        dageo=0
        self.pouno[0]=abs(self.c[0])*poundage
        for i in range(1,self.finalrows):
            dageo=abs(self.c[i])*poundage
            self.pouno[i]=self.pouno[i-1]+dageo


        z=0
        m=0
        j=0
        k=0
        h1={}
        h2={}
        d1={}
        d2={}
        self.d={}
        for i in range(self.finalrows):
            if self.n[i]==1:
                z=z+self.c[i]
                h1[i]=j
                d1[j]=z
                self.d[i]=z
                j=j+1
            else:
                m=self.c[i]+m
                h2[i]=k
                d2[k]=m
                self.d[i]=m
                k=k+1

        v1={}
        v2={}
        self.e={}
        ba1={}
        ba2={}
        r=0
        l=0
        for i in range(self.finalrows):
            if self.n[i]==1:
                if h1[i]==0:
                    ba1[i]=r
                    self.e[i]=self.OrginDate[i].price
                    v1[r]=self.OrginDate[i].price
                    r=r+1
                elif self.d[i]==0:
                    ba1[i]=r
                    self.e[i]=0
                    v1[r]=0
                    r=r+1
                elif d1[h1[i]-1]==0:
                    ba1[i]=r
                    self.e[i]=self.OrginDate[i].price
                    v1[r]=self.OrginDate[i].price
                    r=r+1
                elif b[i]==1:
                    ba1[i]=r
                    self.e[i]=(v1[ba1[i]-1]*d1[h1[i]-1]+self.OrginDate[i].price*self.OrginDate[i].number)/self.d[i]
                    v1[r]=(v1[ba1[i]-1]*d1[h1[i]-1]+self.OrginDate[i].price*self.OrginDate[i].number)/self.d[i]
                    r=r+1
                else:
                    ba1[i]=r
                    self.e[i]=v1[ba1[i]-1]
                    v1[r]=v1[ba1[i]-1]
                    r=r+1
            else:
                if h2[i]==0:
                    ba2[i]=l
                    self.e[i]=self.OrginDate[i].price
                    v2[l]=self.OrginDate[i].price
                    l=l+1
                elif self.d[i]==0:
                    ba2[i]=l
                    self.e[i]=0
                    v2[l]=0
                    l=l+1
                elif d2[h2[i]-1]==0:
                    ba2[i]=l
                    self.e[i]=self.OrginDate[i].price
                    v2[l]=self.OrginDate[i].price
                    l=l+1
                elif b[i]==1:
                    ba2[i]=l
                    self.e[i]=(v2[ba2[i]-1]*d2[h2[i]-1]+self.OrginDate[i].price*self.OrginDate[i].number)/self.d[i]
                    v2[l]=(v2[ba2[i]-1]*d2[h2[i]-1]+self.OrginDate[i].price*self.OrginDate[i].number)/self.d[i]
                    l=l+1
                else:
                    ba2[i]=l
                    self.e[i]=v2[ba2[i]-1]
                    v2[l]=v2[ba2[i]-1]
                    l=l+1

        z=0
        m=0
        j=0
        k=0
        h1={}
        h2={}
        d1={}
        d2={}
        self.d={}
        for i in range(self.finalrows):
            if self.n[i]==1:
                z=z+self.c[i]
                h1[i]=j
                d1[j]=z
                self.d[i]=z
                j=j+1
            else:
                m=self.c[i]+m
                h2[i]=k
                d2[k]=m
                self.d[i]=m
                k=k+1

        self.j={}
        for i in range(self.finalrows):
            if self.n[i]==1:
                if ba1[i]==0:
                    self.j[i]=0
                elif self.d[i]==0 or d1[h1[i]-1]==0:
                    self.j[i]=0
                elif b[i]==1:
                    self.j[i]=(self.OrginDate[i].price-v1[ba1[i]-1])*d1[h1[i]-1]*multiple
                else:
                    self.j[i]=(self.OrginDate[i].price-v1[ba1[i]-1])*self.d[i]*multiple

            else:
                if ba2[i]==0:
                    self.j[i]=0
                elif self.d[i]==0 or d2[h2[i]-1]==0:
                    self.j[i]=0
                elif b[i]==1:
                    self.j[i]=(self.OrginDate[i].price-v2[ba2[i]-1])*d2[h2[i]-1]*(-1)*multiple
                else:
                    self.j[i]=(self.OrginDate[i].price-v2[ba2[i]-1])*self.d[i]*(-1)*multiple

        self.r={}
        o={}
        s={}
        u1={}
        u2={}
        x1=0
        x2=0
        for i in range(self.finalrows):
            if self.n[i]==1:
                if ba1[i]==0:
                    u1[i]=x1
                    self.r[i]=0
                    o[x1]=0
                    x1=x1+1
                elif b1[i]==1:
                    u1[i]=x1
                    self.r[i]=o[u1[i]-1]
                    o[x1]=o[u1[i]-1]
                    x1=x1+1
                else:
                    u1[i]=x1
                    self.r[i]=o[u1[i]-1]+(self.OrginDate[i].price-v1[ba1[i]-1])*self.OrginDate[i].number*multiple
                    o[x1]=o[u1[i]-1]+(self.OrginDate[i].price-v1[ba1[i]-1])*self.OrginDate[i].number*multiple
                    x1=x1+1
            else:
                if ba2[i]==0:
                    u2[i]=x2
                    self.r[i]=0
                    s[x2]=0
                    x2=x2+1
                elif b1[i]==1:
                    u2[i]=x2
                    self.r[i]=s[u2[i]-1]
                    s[x2]=s[u2[i]-1]
                    x2=x2+1
                else:
                    u2[i]=x2
                    self.r[i]=s[u2[i]-1]+(self.OrginDate[i].price-v2[ba2[i]-1])*self.OrginDate[i].number*(-1)*multiple
                    s[x2]=s[u2[i]-1]+(self.OrginDate[i].price-v2[ba2[i]-1])*self.OrginDate[i].number*(-1)*multiple
                    x2=x2+1
        if x1>0:
            self.duo=o[x1-1]
        else:
            self.duo=0
        if x2>0:
            self.kong=s[x2-1]
        else:
            self.kong=0
        sum=0
        self.fin={}
        for i in range(self.finalrows):
            if i==0:
                self.fin[i]=0
            elif self.n[i]==self.n[i-1]:
                self.fin[i]=self.r[i]+self.j[i]+sum
            else:
                sum=self.r[i-1]+self.j[i-1]
                self.fin[i]=self.r[i]+self.j[i]+sum

        self.supfino={}
        for i in range(self.finalrows):
            self.supfino[i]=self.fin[i]-self.pouno[i]

        net=0
        self.net={}
        for i in range(self.finalrows):
            net=net+self.c[i]*self.n[i]
            self.net[i]=net

        self.num1={}
        for i in range(self.finalrows):
            if i==0:
                if self.net[i]>0:
                    self.num1[i]=self.OrginDate[i].number
                else:
                    self.num1[i]=0
            elif self.net[i]>0:
                if self.net[i]>0 and self.net[i-1]<0:
                    self.num1[i]=self.net[i-1]+self.OrginDate[i].number
                else:
                    self.num1[i]=self.OrginDate[i].number
            else:
                if self.net[i]<0 and self.net[i-1]<=0:
                    self.num1[i]=0
                elif self.net[i]<0 and self.net[i-1]>0:
                    self.num1[i]=self.net[i-1]

                elif self.net[i]==0 and self.net[i-1]<0:
                    self.num1[i]=0
                else:
                    self.num1[i]=self.OrginDate[i].number


        self.num2={}
        for i in range(self.finalrows):
            if i==0:
                if self.net[i]<0:
                    self.num2[i]=self.OrginDate[i].number
                else:
                    self.num2[i]=0
            elif self.net[i]<0:
                if self.net[i]<0 and self.net[i-1]>0:
                    self.num2[i]=-self.net[i-1]+self.OrginDate[i].number
                else:
                    self.num2[i]=self.OrginDate[i].number
            else:
                if self.net[i]>0 and self.net[i-1]>=0:
                    self.num2[i]=0
                elif self.net[i]>0 and self.net[i-1]<0:
                    self.num2[i]=-self.net[i-1]
                elif self.net[i]==0 and self.net[i-1]>0:
                    self.num2[i]=0
                else:
                    self.num2[i]=self.OrginDate[i].number
                    
    def spduo(self):

        self.OrginDate1=[]
        for i in range(self.finalrows):
            self.OrginDate1.append(outnet(self.OrginDate[i].direction,self.OrginDate[i].offsetFlag,self.OrginDate[i].price,self.num1[i],self.OrginDate[i].time))
        self.OrginDate1.sort(key=lambda x:x.timedate.split(':'))
    def func1(self):
        ad={}
        for i in range(self.finalrows):
            if self.OrginDate1[i].direction == u''.encode("gbk"):
                ad[i]=1
            else:
                if self.OrginDate1[i].direction == u'买'.encode("gbk"):
                    ad[i]=1
                else:
                    ad[i]=-1

        bd={}
        for i in range(self.finalrows):
            if self.OrginDate1[i].offsetFlag== u'开仓'.encode("gbk"):
                bd[i]=1
            else:
                bd[i]=-1
        duonumber={}
        for i in range(self.finalrows):
            if self.OrginDate1[i].direction==u''.encode("gbk"):
                duonumber[i]=0
            else:
                duonumber[i]=self.OrginDate1[i].number
        self.cd={}
        for i in range(self.finalrows):
            if self.OrginDate1[i].direction == u''.encode("gbk"):
                self.cd[i]=0
            else:
                self.cd[i]=ad[i]*self.OrginDate1[i].number

        nd={}

        for i in range(self.finalrows):
            nd[i]=bd[i]*ad[i]

        zad=0
        self.dad={}
        for i in range(self.finalrows):
            zad=zad+self.cd[i]
            self.dad[i]=zad


        self.ed={}
        for i in range(self.finalrows):
            if i==0:
                self.ed[i]=self.OrginDate1[i].price
            elif self.dad[i]==0:
                self.ed[i]=0
            elif self.dad[i-1]==0:
                self.ed[i]=self.OrginDate1[i].price
            elif ad[i]==1:
                self.ed[i]=(self.ed[i-1]*self.dad[i-1]+self.OrginDate1[i].price*duonumber[i])/self.dad[i]
            else:
                self.ed[i]=self.ed[i-1]

        self.jd={}
        for i in range(self.finalrows):
            if i==0:
                self.jd[i]=0
            elif self.dad[i]==0 or self.dad[i-1]==0:
                self.jd[i]=0
            elif ad[i]==1:
                self.jd[i]=(self.OrginDate1[i].price-self.ed[i-1])*self.dad[i-1]*multiple
            else:
                self.jd[i]=(self.OrginDate1[i].price-self.ed[i-1])*self.dad[i-1]*multiple


        self.rd={}
        for i in range(self.finalrows):
            if i==0:
                self.rd[i]=0
            elif ad[i]==1:
                self.rd[i]=self.rd[i-1]
            else:
                self.rd[i]=self.rd[i-1]+(self.OrginDate1[i].price-self.ed[i-1])*self.OrginDate1[i].number*multiple


        self.find={}
        for i in range(self.finalrows):
            if i==0:
                self.find[i]=0
            else:
                self.find[i]=self.rd[i]+self.jd[i]

    def spduokong(self):
        self.final2=[]
        for i in range(self.finalrows):
            self.final2.append(OutputtradeRecord(self.cd[i],self.dad[i],self.ed[i],self.jd[i],self.rd[i],self.find[i]))


    def spkong(self):

        self.OrginDate2=[]
        for i in range(self.finalrows):
            self.OrginDate2.append(outnet(self.OrginDate[i].direction,self.OrginDate[i].offsetFlag,self.OrginDate[i].price,self.num2[i],self.OrginDate[i].time,))
        self.OrginDate2.sort(key=lambda x:x.timedate.split(':'))

    def func2(self):
        ak={}
        for i in range(self.finalrows):
            if self.OrginDate2[i].direction == u''.encode("gbk"):
                ak[i]=-1
            else:
                if self.OrginDate2[i].direction == u'买'.encode("gbk"):
                    ak[i]=1
                else:
                    ak[i]=-1
        bk={}
        for i in range(self.finalrows):
            if self.OrginDate2[i].offsetFlag== u'开仓'.encode("gbk"):
                bk[i]=1
            else:
                bk[i]=-1

        self.ck={}
        for i in range(self.finalrows):
            if self.OrginDate2[i].direction == u''.encode("gbk"):
                self.ck[i]=0
            else:
                self.ck[i]=ak[i]*self.OrginDate2[i].number

        nk={}
        for i in range(self.finalrows):
            nk[i]=bk[i]*ak[i]

        zak=0
        self.dak={}
        for i in range(self.finalrows):
            zak=zak+self.ck[i]
            self.dak[i]=-zak
        kongnumber={}
        for i in range(self.finalrows):
            if self.OrginDate2[i].direction==u''.encode("gbk"):
                kongnumber[i]=0
            else:
                kongnumber[i]=self.OrginDate2[i].number

        self.ek={}
        for i in range(0,self.finalrows):
            if i==0:
                self.ek[i]=self.OrginDate2[i].price
            elif self.dak[i]==0:
                self.ek[i]=0
            elif self.dak[i-1]==0:
                self.ek[i]=self.OrginDate2[i].price
            elif ak[i]==-1:
                self.ek[i]=(self.ek[i-1]*self.dak[i-1]+self.OrginDate2[i].price*kongnumber[i])/self.dak[i]
            else:
                self.ek[i]=self.ek[i-1]

        self.jk={}
        for i in range(0,self.finalrows):
            if i==0:
                self.jk[i]=0
            elif self.dak[i]==0 or self.dak[i-1]==0:
                self.jk[i]=0
            elif ak[i]==-1:
                self.jk[i]=(self.OrginDate2[i].price-self.ek[i-1])*self.dak[i-1]*multiple*(-1)
            else:
                self.jk[i]=(self.OrginDate2[i].price-self.ek[i-1])*self.dak[i-1]*multiple*(-1)

        self.rk={}
        for i in range(0,self.finalrows):
            if i==0:
                self.rk[i]=0
            elif ak[i]==-1:
                self.rk[i]=self.rk[i-1]
            else:
                self.rk[i]=self.rk[i-1]+(self.OrginDate2[i].price-self.ek[i-1])*self.OrginDate2[i].number*multiple*(-1)


        self.fink={}
        for i in range(0,self.finalrows):
            if i==0:
                self.fink[i]=0
            else:
                self.fink[i]=self.rk[i]+self.jk[i]

    def spduokong2(self):
        self.final3=[]
        for i in range(self.finalrows):
            self.final3.append(OutputtradeRecord(self.ck[i],self.dak[i],self.ek[i],self.jk[i],self.rk[i],self.fink[i]))



    def finalresult(self,strfilename):
        final=[]
        ping[strfilename]=self.fin[self.finalrows-1]
        for i in range(self.finalrows):
            final.append(OutputTradeRecord(self.c[i],self.d[i],self.e[i],self.j[i],self.r[i],self.fin[i],self.net[i],self.pouno[i],self.supfino[i]))
            sql="""
                insert into {0} (
                timedate,price,net,result,poundage,pureresult,duonet,duomovement,kongnet,kongmovement)
                values('{1}',{2},{3},{4},{5},{6},{7},{8},{9},{10})
                """.format('out_'+strfilename,self.OrginDate[i].time,self.OrginDate[i].price,final[i].duokong,final[i].amount,final[i].poundage,final[i].finalfinal,self.final2[i].amount,self.final2[i].movement,self.final3[i].amount,self.final3[i].movement)
            cursor3.execute(sql)
            out.commit()

    def everyone(self,k,strfilename):
        c=0
        duocount=0
        kongcount=0
        for i in range(self.rows):
            c=c+self.data1[i][3]
        for i in range(self.finalrows):
            if self.n[i]==1 and self.OrginDate[i].number!='':
                duocount=duocount+self.OrginDate[i].number
            elif self.n[i]==-1 and self.OrginDate[i].number!='':
                kongcount=kongcount+self.OrginDate[i].number
        sql="""
            insert into [statement(date)] (
            date,pingprofit,volume,pingprofitduo,pingprofitkong,volumeduo,volumekong)
            values('{0}',{1},{2},{3},{4},{5},{6})
            """.format(strfilename,ping[strfilename],c,self.duo,self.kong,duocount,kongcount)
        cursor3.execute(sql)
        out.commit()  

if __name__=='__main__':

    allsheet=[]
    existsheet=[]
    newsheet=[]
    calount=cursor2.execute("select count(*) from sysobjects where xtype='U'").fetchone()
    print calount[0]
    sheetname=cursor2.execute("select name from sysobjects where xtype='U'").fetchall()
    existname=cursor3.execute("select name from sysobjects where xtype='U'").fetchall()
    row=cursor3.execute("select count(*) from sysobjects where xtype='U'").fetchone()
    row=row[0]
    for i in range(row):
        if "out" in existname[i][0]:
            existsheet.append(existname[i][0].split('_')[1])
    for i in range(calount[0]):
        allsheet.append(sheetname[i][0].split('_')[1])
    print allsheet
    for i in range(calount[0]):
        if allsheet[i] not in existsheet:
            newsheet.append(allsheet[i])
    t=Methods()


    ping={}


    for k in range(len(newsheet)):
        strfilename=newsheet[k]
        t.readExcel(strfilename)
        t.readExcel2(strfilename)
        t.readExcel3(strfilename)
        sql="""
            create table {0}(
            timedate varchar(50),
            price decimal(20, 4),
            result decimal(20, 4),
            net decimal(20, 4),
            poundage decimal(20, 4),
            pureresult decimal(20, 4),
            duonet decimal(20, 4),
            duomovement decimal(20, 4),
            kongnet decimal(20, 4),
            kongmovement decimal(20, 4)
            )
            """.format('out_'+strfilename)
        cursor3.execute(sql)
        out.commit()
        t.deal()
        t.spduo()
        t.func1()
        t.spduokong()
        t.spkong()
        t.func2()
        t.spduokong2()
        t.finalresult(strfilename)
        t.everyone(k,strfilename)

class OperExcel():
    def rExcel(self):

        row=cursor1.execute("select count(*) from mock").fetchall()
        sql="""
            create table allname(
            direction varchar(50),
            offsetflag varchar(50),
            price decimal(20, 4),
            number int,
            timedate varchar(50),
            name varchar(50))
            """
        cursor1.execute(sql)
        tr.commit()
        sql="insert into allname(direction,offsetflag,price,number,timedate,name)select direction,offsetflag,price,number,date+' '+time,name from mock"
        cursor1.execute(sql)
        tr.commit()
        name=cursor1.execute("select name from mock").fetchall()
        self.rows=cursor1.execute("select count(*) from mock").fetchall()
        for i in range(self.rows[0][0]):
            mylist.append(name[i][0])


    def deal(self):

        sql="""
            create table {0}(
            direction varchar(50),
            offsetflag varchar(50),
            price decimal(20, 4),
            number int,
            timedate varchar(50))
            """.format(strfilename)
        cursor4.execute(sql)
        op.commit()
        sql="""
            insert into {0} (
            direction,offsetflag,price,number,timedate)
            select direction,offsetflag,price,number,timedate
            from [TradeRecord].[dbo].[allname] where name='{1}'
            """.format(strfilename,strfilename)
        cursor4.execute(sql)
        op.commit()
    def deal2(self):
        sql="delete from {0}".format(strfilename)
        cursor4.execute(sql)
        op.commit()
        sql="""
            insert into {0} (
            direction,offsetflag,price,number,timedate)
            select direction,offsetflag,price,number,timedate
            from [TradeRecord].[dbo].[allname] where name='{1}'
            """.format(strfilename,strfilename)
        cursor4.execute(sql)
        op.commit()


if __name__=='__main__':
    newlist=[]
    mylist=[]
    t=OperExcel()
    t.rExcel()

    allsheet=[]
    newlist=list(set(mylist))
    calount=cursor4.execute("select count(*) from sysobjects where xtype='U'").fetchone()
    print calount[0]
    sheetname=cursor4.execute("select name from sysobjects where xtype='U'").fetchall()
    for i in range(calount[0]):
        allsheet.append(sheetname[i][0])
    print type(allsheet[0])
    print type(newlist[0].decode("gbk"))
    for i in range(len(newlist)):
        if newlist[i].decode("gbk") not in allsheet:
            print i
            strfilename=newlist[i]
            t.deal()
        else:
            print i
            strfilename=newlist[i]

            t.deal2()





OrginDate = []
final=[]
duo=[]
kong=[]

class TradeRecord(object):
    def __init__(self,offsetFlag,direction,price,number,time):
        self.time=str(time)
        self.offsetFlag = offsetFlag
        self.direction = direction
        self.price = price
        self.number = number
class OutputTradeRecord(object):
    def __init__(self,change,openinterest,avgprice,avgprofit,pingprofit,amount,duokong):
        self.change=change
        self.openinterest=openinterest
        self.avgprice=avgprice
        self.avgprofit=avgprofit
        self.pingprofit=pingprofit
        self.duokong=duokong
        self.amount=amount


class Methods(object):


    def readExcel(self):


        sql="delete from [duo];delete from [kong]"
        cursor4.execute(sql)
        op.commit()
        self.rows=cursor1.execute("select count(*) from alldate").fetchall()
        self.rows=self.rows[0][0]
        self.data=cursor1.execute("select direction, offsetflag, price, number, timedate from alldate").fetchall()
        for i in range(self.rows):
            _record = TradeRecord(self.data[i][1],self.data[i][0],self.data[i][2],self.data[i][3],self.data[i][4])
            OrginDate.append(_record)
        OrginDate.sort(key=lambda x:x.time.split(':'))
    def deal(self):
        a={}
        for i in range(self.rows):
            if OrginDate[i].direction== u'买'.encode("gbk"):
                a[i]=1
            else:
                a[i]=-1
        b={}
        for i in range(self.rows):
            if OrginDate[i].offsetFlag == u'开仓'.encode("gbk"):
                b[i]=1
            else:
                b[i]=-1
        self.n={}
        for i in range(self.rows):
            self.n[i]=a[i]*b[i]
    def addduokong(self):


        for i in range(self.rows):
            if self.n[i]==1:
                sql="""
                 insert into {0} (
                 direction,offsetflag,price,number,timedate)
                 values('{1}','{2}',{3},{4},'{5}')
                 """.format('duo',OrginDate[i].direction,OrginDate[i].offsetFlag,OrginDate[i].price,OrginDate[i].number,OrginDate[i].time)
                cursor4.execute(sql)
                op.commit()

            else:
                sql="""
                 insert into {0} (
                 direction,offsetflag,price,number,timedate)
                 values('{1}','{2}',{3},{4},'{5}')
                 """.format('kong',OrginDate[i].direction,OrginDate[i].offsetFlag,OrginDate[i].price,OrginDate[i].number,OrginDate[i].time)
                cursor4.execute(sql)
                op.commit()

if __name__=='__main__':
    t=Methods()
    t.readExcel()
    t.deal()

    t.addduokong()

OrginDate = []    
class OutputTradeRecord(object):
    def __init__(self,number,change,openinterest,avgprice,avgprofit,pingprofit,amount,duokong,net,poundage,finalfinal):
        self.change=change
        self.openinterest=openinterest
        self.avgprice=avgprice
        self.avgprofit=avgprofit
        self.pingprofit=pingprofit
        self.duokong=duokong
        self.amount=amount
        self.net=net
        self.number=number
        self.poundage=poundage
        self.finalfinal=finalfinal
        
class OutputtradeRecord(object):
    def __init__(self,change,amount,avprice,position,closeprofit,movement,poundage,finalfinal):
        self.change=change
        self.amount=amount
        self.avgprice=avprice
        self.position=position
        self.closeprofit=closeprofit
        self.movement=movement
        self.poundage=poundage
        self.finalfinal=finalfinal
class outnet(object):
    def __init__(self,direction,offsetFlag,price,number,time):
        self.direction=direction
        self.offsetFlag=offsetFlag
        self.price=price
        self.number=number
        self.time=time        
class deal(object):
    def readExcel(self):

        self.daylist=[]
        sql="delete from {0};delete from {1};delete from {2}".format('[summary statement]','netduo','netkong')
        cursor3.execute(sql)
        out.commit()
        self.rows=cursor1.execute("select count(*) from alldate").fetchall()
        self.rows=self.rows[0][0]
        self.data=cursor1.execute("select direction, offsetflag, price, number, timedate from alldate").fetchall()
        for i in range(self.rows):
            _record = TradeRecord(self.data[i][1],self.data[i][0],self.data[i][2],self.data[i][3],self.data[i][4])
            OrginDate.append(_record)
            self.daylist.append(self.data[i][4].split(' ')[0])
        self.daylist=list(set(self.daylist))
        OrginDate.sort(key=lambda x:x.time.split(':'))
    def readExcel2(self):
        self.ordata=[]
        self.OrginDate=[]
        self.OrginDate2=[]

        self.rows2=cursor5.execute("select count(*) from min").fetchone()
        self.rows2=self.rows2[0]
        self.data=cursor5.execute("select timedate, price from min").fetchall()
        for i in range(self.rows2):
            if self.data[i][0].split(' ')[0] in self.daylist:
                if int(self.data[i][0].split(' ')[1].split(':')[0])<17:
                    _record = TradeRecord('','',self.data[i][1],'',self.data[i][0])
                    self.OrginDate.append(_record)
                    self.OrginDate2.append(_record)
        self.OrginDate.sort(key=lambda x:x.time.split(':'))
        self.OrginDate2.sort(key=lambda x:x.time.split(':'))
        self.rows2=len(self.OrginDate)
    def deal2(self):

        self.a={}
        for i in range(self.rows):
            if OrginDate[i].direction== u'买'.encode("gbk"):
                self.a[i]=1
            else:
                self.a[i]=-1
        b={}
        for i in range(self.rows):
            if OrginDate[i].offsetFlag == u'开仓'.encode("gbk"):
                b[i]=1
            else:
                b[i]=-1

        self.c={}
        for i in range(self.rows):
            self.c[i]=b[i]*OrginDate[i].number

        self.pounda={}
        dageda=0
        self.pounda[0]=abs(self.c[0])*poundage
        for i in range(1,self.rows):
            dageda=abs(self.c[i])*poundage
            self.pounda[i]=self.pounda[i-1]+dageda
            pre=self.pounda[i]
        self.expre=pre

        self.n={}
        for i in range(self.rows):
            self.n[i]=self.a[i]*b[i]

        z=0
        m=0
        j=0
        k=0
        h1={}
        h2={}
        d1={}
        d2={}
        self.d={}
        for i in range(self.rows):
            if self.n[i]==1:
                z=z+self.c[i]
                h1[i]=j
                d1[j]=z
                self.d[i]=z
                j=j+1
            else:
                m=self.c[i]+m
                h2[i]=k
                d2[k]=m
                self.d[i]=m
                k=k+1
        v1={}
        v2={}
        self.e={}
        ba1={}
        ba2={}
        r=0
        l=0
        for i in range(self.rows):
            if self.n[i]==1:
                if h1[i]==0:
                    ba1[i]=r
                    self.e[i]=OrginDate[i].price
                    v1[r]=OrginDate[i].price
                    r=r+1
                elif self.d[i]==0:
                    ba1[i]=r

                    self.e[i]=0
                    v1[r]=0
                    r=r+1
                elif d1[h1[i]-1]==0:
                    ba1[i]=r
                    self.e[i]=OrginDate[i].price
                    v1[r]=OrginDate[i].price
                    r=r+1
                elif b[i]==1:
                    ba1[i]=r
                    self.e[i]=(v1[ba1[i]-1]*d1[h1[i]-1]+OrginDate[i].price*OrginDate[i].number)/self.d[i]
                    v1[r]=(v1[ba1[i]-1]*d1[h1[i]-1]+OrginDate[i].price*OrginDate[i].number)/self.d[i]
                    r=r+1
                else:
                    ba1[i]=r
                    self.e[i]=v1[ba1[i]-1]
                    v1[r]=v1[ba1[i]-1]
                    r=r+1
            else:
                if h2[i]==0:
                    ba2[i]=l
                    self.e[i]=OrginDate[i].price
                    v2[l]=OrginDate[i].price
                    l=l+1
                elif self.d[i]==0:
                    ba2[i]=l
                    self.e[i]=0
                    v2[l]=0
                    l=l+1
                elif d2[h2[i]-1]==0:
                    ba2[i]=l
                    self.e[i]=OrginDate[i].price
                    v2[l]=OrginDate[i].price
                    l=l+1
                elif b[i]==1:
                    ba2[i]=l
                    self.e[i]=(v2[ba2[i]-1]*d2[h2[i]-1]+OrginDate[i].price*OrginDate[i].number)/self.d[i]
                    v2[l]=(v2[ba2[i]-1]*d2[h2[i]-1]+OrginDate[i].price*OrginDate[i].number)/self.d[i]
                    l=l+1
                else:
                    ba2[i]=l
                    self.e[i]=v2[ba2[i]-1]
                    v2[l]=v2[ba2[i]-1]
                    l=l+1
        self.j={}
        for i in range(self.rows):
            if self.n[i]==1:
                if ba1[i]==0:
                    self.j[i]=0
                elif self.d[i]==0 or d1[h1[i]-1]==0:
                    self.j[i]=0
                elif b[i]==1:
                    self.j[i]=(OrginDate[i].price-v1[ba1[i]-1])*d1[h1[i]-1]*multiple
                else:
                    self.j[i]=(OrginDate[i].price-v1[ba1[i]-1])*self.d[i]*multiple

            else:
                if ba2[i]==0:
                    self.j[i]=0
                elif self.d[i]==0 or d2[h2[i]-1]==0:
                    self.j[i]=0
                elif b[i]==1:
                    self.j[i]=(OrginDate[i].price-v2[ba2[i]-1])*d2[h2[i]-1]*multiple*(-1)
                else:
                    self.j[i]=(OrginDate[i].price-v2[ba2[i]-1])*self.d[i]*multiple*(-1)
        self.r={}
        o={}
        s={}
        u1={}
        u2={}
        x1=0
        x2=0
        for i in range(self.rows):
            if self.n[i]==1:
                if ba1[i]==0:
                    u1[i]=x1
                    self.r[i]=0
                    o[x1]=0
                    x1=x1+1
                elif b[i]==1:
                    u1[i]=x1
                    self.r[i]=o[u1[i]-1]
                    o[x1]=o[u1[i]-1]
                    x1=x1+1
                else:
                    u1[i]=x1
                    self.r[i]=o[u1[i]-1]+(OrginDate[i].price-v1[ba1[i]-1])*OrginDate[i].number*multiple
                    o[x1]=o[u1[i]-1]+(OrginDate[i].price-v1[ba1[i]-1])*OrginDate[i].number*multiple
                    x1=x1+1
            else:
                if ba2[i]==0:
                    u2[i]=x2
                    self.r[i]=0
                    s[x2]=0
                    x2=x2+1
                elif b[i]==1:
                    u2[i]=x2
                    self.r[i]=s[u2[i]-1]
                    s[x2]=s[u2[i]-1]
                    x2=x2+1
                else:
                    u2[i]=x2
                    self.r[i]=s[u2[i]-1]+(OrginDate[i].price-v2[ba2[i]-1])*OrginDate[i].number*(-1)*multiple
                    s[x2]=s[u2[i]-1]+(OrginDate[i].price-v2[ba2[i]-1])*OrginDate[i].number*(-1)*multiple
                    x2=x2+1
        sum=0
        self.fin={}
        for i in range(self.rows):
            if i==0:
                self.fin[i]=0
            elif self.a[i]*b[i]==self.a[i-1]*b[i-1]:
                self.fin[i]=self.r[i]+self.j[i]+sum
            else:
                sum=self.r[i-1]+self.j[i-1]
                self.fin[i]=self.r[i]+self.j[i]+sum
        self.supfinda={}
        for i in range(self.rows):
            self.supfinda[i]=self.fin[i]-self.pounda[i]

        net=0
        self.net={}
        for i in range(self.rows):
            net=net+self.c[i]*self.n[i]
            self.net[i]=net

        self.num1={}
        for i in range(self.rows):
            if i==0:
                if self.net[i]>0:
                    self.num1[i]=OrginDate[i].number
                else:
                    self.num1[i]=0
            elif self.net[i]>0:
                if self.net[i]>0 and self.net[i-1]<0:
                    self.num1[i]=self.net[i-1]+OrginDate[i].number
                else:
                    self.num1[i]=OrginDate[i].number
            else:
                if self.net[i]<0 and self.net[i-1]<=0:
                    self.num1[i]=0
                elif self.net[i]<0 and self.net[i-1]>0:
                    self.num1[i]=self.net[i-1]

                elif self.net[i]==0 and self.net[i-1]<0:
                    self.num1[i]=0
                else:
                    self.num1[i]=OrginDate[i].number


        self.num2={}
        for i in range(self.rows):
            if i==0:
                if self.net[i]<0:
                    self.num2[i]=OrginDate[i].number
                else:
                    self.num2[i]=0
            elif self.net[i]<0:
                if self.net[i]<0 and self.net[i-1]>0:
                    self.num2[i]=-self.net[i-1]+OrginDate[i].number
                else:
                    self.num2[i]=OrginDate[i].number
            else:
                if self.net[i]>0 and self.net[i-1]>=0:
                    self.num2[i]=0
                elif self.net[i]>0 and self.net[i-1]<0:
                    self.num2[i]=-self.net[i-1]
                elif self.net[i]==0 and self.net[i-1]>0:
                    self.num2[i]=0
                else:
                    self.num2[i]=OrginDate[i].number  
    def finalresult1(self):
        final=[]
        for i in range(self.rows):
            final.append(OutputTradeRecord(self.num1[i],self.c[i],self.d[i],self.e[i],self.j[i],self.r[i],self.fin[i],self.n[i],self.net[i],self.pounda[i],self.supfinda[i]))
            sql="""
                insert into {0} (
                timedate,price,net,result,poundage,pureresult)
                values('{1}',{2},{3},{4},{5},{6})
                """.format('[summary statement]',OrginDate[i].time,OrginDate[i].price,final[i].net,final[i].amount,final[i].poundage,final[i].finalfinal)
            cursor3.execute(sql)
            out.commit()
    def addchicangliang(self):
        sql="delete from {0}".format('[openposition]')
        cursor3.execute(sql)
        out.commit()
        down={}
        write={}
        timetime={}
        for j in range(10):
            yy=-100+j*20
            print yy
            p1={}
            p2={}
            volume=0
            for i in range(self.rows):
                if self.net[i]>=yy and self.net[i]<yy+20:
                    p1[i]=self.fin[i]
                    p2[i]=OrginDate[i].time
                    volume=volume+OrginDate[i].number
            down[j]=volume
            csv=0
            timecalculate=0
            if len(p1)==0:
                write[j]=csv
                timetime[j]=timecalculate
                print 0
            else:
                key=p1.keys()
                zero=p1.get(key[0])
                time1=p2.get(key[0])
                k=0
                for i in range(1,len(key)):

                    if key[i]-key[i-1]!=1:

                        if k>=1:
                            cclj=p1.get(key[i-1])-zero

                            zero=p1.get(key[i])
                            csv=csv+cclj

                            k=0
                        else:
                            cclj=p1.get(key[i-1])
                            zero=p1.get(key[i])
                            csv=csv+cclj

                            k=0
                    else:
                        k=k+1
                cclj=p1.get(key[i])-zero
                csv=csv+cclj
                print csv
                write[j]=csv
                for i in range(1,len(key)):
                    if key[i]-key[i-1]!=1:
                        time2=p2.get(key[i-1])

                        a=time.strptime(time1, "%Y%m%d %H:%M:%S")
                        b=time.strptime(time2, "%Y%m%d %H:%M:%S")
                        starttime=datetime.datetime(a[0],a[1],a[2],a[3],a[4],a[5])
                        endtime=datetime.datetime(b[0],b[1],b[2],b[3],b[4],b[5])
                        cxsj=(endtime-starttime).seconds
                        time1=p2.get(key[i])
                        timecalculate=timecalculate+cxsj
                time2=p2.get(key[i])
                a=time.strptime(time1, "%Y%m%d %H:%M:%S")
                b=time.strptime(time2, "%Y%m%d %H:%M:%S")
                starttime=datetime.datetime(a[0],a[1],a[2],a[3],a[4],a[5])
                endtime=datetime.datetime(b[0],b[1],b[2],b[3],b[4],b[5])
                cxsj=(endtime-starttime).seconds
                timecalculate=timecalculate+cxsj
                timetime[j]=timecalculate

        sql="""
                insert into {0} (
                "-100to-80","-80to-60","-60to-40","-40to-20","-20to0","0to20","20to40","40to60","60to80","80to100")
                values({1},{2},{3},{4},{5},{6},{7},{8},{9},{10})
                """.format('[openposition]',write[0],write[1],write[2],write[3],write[4],write[5],write[6],write[7],write[8],write[9])
        cursor3.execute(sql)
        out.commit()
        sql="""
                insert into {0} (
                "-100to-80","-80to-60","-60to-40","-40to-20","-20to0","0to20","20to40","40to60","60to80","80to100")
                values({1},{2},{3},{4},{5},{6},{7},{8},{9},{10})
                """.format('[openposition]',down[0],down[1],down[2],down[3],down[4],down[5],down[6],down[7],down[8],down[9])
        cursor3.execute(sql)
        out.commit()
        sql="""
                insert into {0} (
                "-100to-80","-80to-60","-60to-40","-40to-20","-20to0","0to20","20to40","40to60","60to80","80to100")
                values({1},{2},{3},{4},{5},{6},{7},{8},{9},{10})
                """.format('[openposition]',timetime[0],timetime[1],timetime[2],timetime[3],timetime[4],timetime[5],timetime[6],timetime[7],timetime[8],timetime[9])
        cursor3.execute(sql)
        out.commit()
    def spduo(self):
        for i in range(self.rows):
            self.OrginDate.append(outnet(OrginDate[i].direction,OrginDate[i].offsetFlag,OrginDate[i].price,self.num1[i],OrginDate[i].time))
        self.OrginDate.sort(key=lambda x:x.time.split(':'))
    def func1(self):
        ad={}
        for i in range(self.rows+self.rows2):
            if self.OrginDate[i].direction == u''.encode("gbk"):
                ad[i]=1
            else:
                if self.OrginDate[i].direction == u'买'.encode("gbk"):
                    ad[i]=1
                else:
                    ad[i]=-1

        bd={}
        for i in range(self.rows+self.rows2):
            if self.OrginDate[i].offsetFlag== u'开仓'.encode("gbk"):
                bd[i]=1
            else:
                bd[i]=-1
        abc={}
        for i in range(self.rows+self.rows2):
            if self.OrginDate[i].direction==u''.encode("gbk"):
                abc[i]=0
            else:
                abc[i]=self.OrginDate[i].number

        self.cd={}
        for i in range(self.rows+self.rows2):
            if self.OrginDate[i].direction == u''.encode("gbk"):
                self.cd[i]=0
            else:
                self.cd[i]=ad[i]*self.OrginDate[i].number
        self.poundageduo={}
        dage=0
        self.poundageduo[0]=abs(self.cd[0])*poundage
        for i in range(1,self.rows+self.rows2):
            dage=abs(self.cd[i])*poundage
            self.poundageduo[i]=self.poundageduo[i-1]+dage

        nd={}

        for i in range(self.rows+self.rows2):
            nd[i]=bd[i]*ad[i]

        zad=0
        self.dad={}
        for i in range(self.rows+self.rows2):
            zad=zad+self.cd[i]
            self.dad[i]=zad


        self.ed={}
        for i in range(0,self.rows+self.rows2):
            if i==0:
                self.ed[i]=self.OrginDate[i].price
            elif self.dad[i]==0:
                self.ed[i]=0
            elif self.dad[i-1]==0:
                self.ed[i]=self.OrginDate[i].price
            elif ad[i]==1:
                self.ed[i] = (self.ed[i-1]*self.dad[i-1]+self.OrginDate[i].price*abc[i])/self.dad[i]
            else:
                self.ed[i]=self.ed[i-1]

        self.jd={}
        for i in range(0,self.rows+self.rows2):
            if i==0:
                self.jd[i]=0
            elif self.dad[i]==0 or self.dad[i-1]==0:
                self.jd[i]=0
            elif ad[i]==1:
                self.jd[i]=(self.OrginDate[i].price-self.ed[i-1])*self.dad[i-1]*multiple
            else:
                self.jd[i]=(self.OrginDate[i].price-self.ed[i-1])*self.dad[i-1]*multiple


        self.rd={}
        for i in range(0,self.rows+self.rows2):
            if i==0:
                self.rd[i]=0
            elif ad[i]==1:
                self.rd[i]=self.rd[i-1]
            else:
                self.rd[i]=self.rd[i-1]+(self.OrginDate[i].price-self.ed[i-1])*self.OrginDate[i].number*multiple*(-1)


        self.find={}
        for i in range(0,self.rows+self.rows2):
            if i==0:
                self.find[i]=0
            else:
                self.find[i]=self.rd[i]+self.jd[i]

        self.supfin={}
        for i in range(self.rows+self.rows2):
            self.supfin[i]=self.find[i]-self.poundageduo[i]       
    def spduokong(self):
        final2=[]
        for i in range(self.rows+self.rows2):
            final2.append(OutputtradeRecord(self.cd[i],self.dad[i],self.ed[i],self.jd[i],self.rd[i],self.find[i],self.poundageduo[i],self.supfin[i]))  
            sql="""
                insert into {0} (
                timedate,price,net,result,poundage,pureresult)
                values('{1}',{2},{3},{4},{5},{6})
                """.format('netduo',self.OrginDate[i].time,self.OrginDate[i].price,final2[i].amount,final2[i].movement,final2[i].poundage,final2[i].finalfinal)
            cursor3.execute(sql)
            out.commit()    
    def spkong(self):
        for i in range(self.rows):
            self.OrginDate2.append(outnet(OrginDate[i].direction,OrginDate[i].offsetFlag,OrginDate[i].price,self.num2[i],OrginDate[i].time,))
        self.OrginDate2.sort(key=lambda x:x.time.split(':'))

    def func2(self):
        ak={}
        for i in range(self.rows+self.rows2):
            if self.OrginDate2[i].direction == u''.encode("gbk"):
                ak[i]=-1
            else:
                if self.OrginDate2[i].direction == u'买'.encode("gbk"):
                    ak[i]=1
                else:
                    ak[i]=-1
        bk={}
        for i in range(self.rows+self.rows2):
            if self.OrginDate2[i].offsetFlag== u'开仓'.encode("gbk"):
                bk[i]=1
            else:
                bk[i]=-1

        self.ck={}
        for i in range(self.rows+self.rows2):
            if self.OrginDate2[i].direction == u''.encode("gbk"):
                self.ck[i]=0
            else:
                self.ck[i]=ak[i]*self.OrginDate2[i].number
        self.poundagekong={}
        dage=0
        self.poundagekong[0]=abs(self.ck[0])*poundage
        for i in range(1,self.rows+self.rows2):
            dage=abs(self.ck[i])*poundage
            self.poundagekong[i]=self.poundagekong[i-1]+dage
        nk={}
        for i in range(self.rows+self.rows2):
            nk[i]=bk[i]*ak[i]

        zak=0
        self.dak={}
        for i in range(self.rows+self.rows2):
            zak=zak+self.ck[i]
            self.dak[i]=-zak
        abd={}
        for i in range(self.rows+self.rows2):
            if self.OrginDate2[i].direction==u''.encode("gbk"):
                abd[i]=0
            else:
                abd[i]=self.OrginDate2[i].number

        self.ek={}
        for i in range(0,self.rows+self.rows2):
            if i==0:
                self.ek[i]=self.OrginDate2[i].price
            elif self.dak[i]==0:
                self.ek[i]=0
            elif self.dak[i-1]==0:
                self.ek[i]=self.OrginDate2[i].price
            elif ak[i]==-1:
                self.ek[i]=(self.ek[i-1]*self.dak[i-1]+self.OrginDate2[i].price*abd[i])/self.dak[i]
            else:
                self.ek[i]=self.ek[i-1]

        self.jk={}
        for i in range(0,self.rows+self.rows2):
            if i==0:
                self.jk[i]=0
            elif self.dak[i]==0 or self.dak[i-1]==0:
                self.jk[i]=0
            elif ak[i]==-1:
                self.jk[i]=(self.OrginDate2[i].price-self.ek[i-1])*self.dak[i-1]*multiple
            else:
                self.jk[i]=(self.OrginDate2[i].price-self.ek[i-1])*self.dak[i-1]*multiple

        self.rk={}
        for i in range(0,self.rows+self.rows2):
            if i==0:
                self.rk[i]=0
            elif ak[i]==-1:
                self.rk[i]=self.rk[i-1]
            else:
                self.rk[i]=self.rk[i-1]+(self.OrginDate2[i].price-self.ek[i-1])*self.OrginDate2[i].number*multiple


        self.fink={}
        for i in range(0,self.rows+self.rows2):
            if i==0:
                self.fink[i]=0
            else:
                self.fink[i]=self.rk[i]+self.jk[i]

        self.supfin={}
        for i in range(self.rows+self.rows2):
            self.supfin[i]=self.fink[i]-self.poundagekong[i]

    def spduokong2(self):
        final3=[]
        for i in range(self.rows+self.rows2):
            final3.append(OutputtradeRecord(self.ck[i],self.dak[i],self.ek[i],self.jk[i],self.rk[i],self.fink[i],self.poundagekong[i],self.supfin[i])) 
            sql="""
                insert into {0} (
                timedate,price,net,result,poundage,pureresult)
                values('{1}',{2},{3},{4},{5},{6})
                """.format('netkong',self.OrginDate2[i].time,self.OrginDate2[i].price,final3[i].amount,final3[i].movement,final3[i].poundage,final3[i].finalfinal)
            cursor3.execute(sql)
            out.commit()              
if __name__=='__main__':
    t=deal()


    t.readExcel()
    t.readExcel2()
    t.deal2()
    t.finalresult1()
    #t.addchicangliang()
    t.spduo()
    t.func1()
    t.spduokong()
    t.spkong()
    t.func2()
    t.spduokong2()
    
class TradeRecord(object):
    def __init__(self,time,offsetFlag,direction,price,number):
        self.time=str(time)
        self.offsetFlag = offsetFlag
        self.direction = direction
        self.price = price
        self.number = number
class OutputTradeRecord(object):
    def __init__(self,change,openinterest,avgprice,avgprofit,pingprofit,amount,duokong,poundage,finalfinal):
        self.change=change
        self.openinterest=openinterest
        self.avgprice=avgprice
        self.avgprofit=avgprofit
        self.pingprofit=pingprofit
        self.amount=amount
        self.duokong=duokong
        self.poundage=poundage
        self.finalfinal=finalfinal
class OutputtradeRecord(object):
    def __init__(self,change,amount,avprice,position,closeprofit,movement):
        self.change=change
        self.amount=amount
        self.avprice=avprice
        self.position=position
        self.closeprofit=closeprofit
        self.movement=movement
class outnet(object):
    def __init__(self,direction,offsetFlag,price,number,timedate):
        self.direction=direction
        self.offsetFlag=offsetFlag
        self.price=price
        self.number=number
        self.timedate=timedate        
class Methods(object):


    def readExcel(self,strfilename):
        self.OrginDate=[]
        self.daylist=[]
        self.rows=cursor4.execute("select count(*) from {0}".format(strfilename)).fetchone()
        self.data1=cursor4.execute("select direction, offsetflag, price, number, timedate from {0}".format(strfilename)).fetchall()

        print self.rows[0]
        self.rows=self.rows[0]
        for i in range(self.rows):
            _record = TradeRecord(self.data1[i][4],self.data1[i][1],self.data1[i][0],self.data1[i][2],self.data1[i][3])
            self.OrginDate.append(_record)
            self.daylist.append(self.data1[i][4].split(' ')[0])
            self.daylist=list(set(self.daylist))


    def readExcel2(self):



        self.rows2=cursor5.execute("select count(*) from min").fetchone()
        self.rows2=self.rows2[0]
        self.data=cursor5.execute("select timedate, price from min").fetchall()
        for i in range(self.rows2):
            if self.data[i][0].split(' ')[0] in self.daylist:
                if int(self.data[i][0].split(' ')[1].split(':')[0])<17:
                    _record = TradeRecord(self.data[i][0],'','',self.data[i][1],'')
                    self.OrginDate.append(_record)
        self.OrginDate.sort(key=lambda x:x.time.split(':'))
        self.rows2=len(self.OrginDate)-self.rows
        
    def deal(self):
        a={}
        for i in range(self.rows+self.rows2):
            if self.OrginDate[i].direction == u'买'.encode("gbk"):
                a[i]=1
            else:
                a[i]=-1
        b={}
        for i in range(self.rows+self.rows2):
            if self.OrginDate[i].offsetFlag== u'开仓'.encode("gbk"):
                b[i]=1
            else:
               b[i]=-1

        b1={}
        for i in range(self.rows+self.rows2):
            if self.OrginDate[i].direction==u''.encode("gbk"):
                b1[i]=1
            else:
                if self.OrginDate[i].offsetFlag == u'开仓'.encode("gbk"):
                    b1[i]=1
                else:
                    b1[i]=-1
        self.n={}
        if self.OrginDate[0].direction==u''.encode("gbk"):
            self.n[0]=1
        else:
            self.n[0]=a[0]*b[0]
        for i in range(1,self.rows+self.rows2):
            if self.OrginDate[i].direction==u''.encode("gbk"):
                self.n[i]=self.n[i-1]
            else:
                self.n[i]=a[i]*b[i]

        abc={}
        for i in range(self.rows+self.rows2):
            if self.OrginDate[i].direction==u''.encode("gbk"):
                abc[i]=0
            else:
                abc[i]=self.OrginDate[i].number

        self.c={}
        for i in range(self.rows+self.rows2):
            if self.OrginDate[i].direction == u''.encode("gbk"):
                self.c[i]=0
            else:
                self.c[i]=b[i]*self.OrginDate[i].number
        self.pouno={}
        dageo=0
        self.pouno[0]=abs(self.c[0])*poundage
        for i in range(1,self.rows+self.rows2):
            dageo=abs(self.c[i])*poundage
            self.pouno[i]=self.pouno[i-1]+dageo

        z=0
        m=0
        j=0
        k=0
        h1={}
        h2={}
        d1={}
        d2={}
        self.d={}
        for i in range(self.rows+self.rows2):
            if self.n[i]==1:
                z=z+self.c[i]
                h1[i]=j
                d1[j]=z
                self.d[i]=z
                j=j+1
            else:
                m=self.c[i]+m
                h2[i]=k
                d2[k]=m
                self.d[i]=m
                k=k+1

        v1={}
        v2={}
        self.e={}
        ba1={}
        ba2={}
        r=0
        l=0
        for i in range(self.rows+self.rows2):
            if self.n[i]==1:
                if h1[i]==0:
                    ba1[i]=r
                    self.e[i]=self.OrginDate[i].price
                    v1[r]=self.OrginDate[i].price
                    r=r+1
                elif self.d[i]==0:
                    ba1[i]=r
                    self.e[i]=0
                    v1[r]=0
                    r=r+1
                elif d1[h1[i]-1]==0:
                    ba1[i]=r
                    self.e[i]=self.OrginDate[i].price
                    v1[r]=self.OrginDate[i].price
                    r=r+1
                elif b[i]==1:
                    ba1[i]=r
                    self.e[i]=(v1[ba1[i]-1]*d1[h1[i]-1]+self.OrginDate[i].price*self.OrginDate[i].number)/self.d[i]
                    v1[r]=(v1[ba1[i]-1]*d1[h1[i]-1]+self.OrginDate[i].price*self.OrginDate[i].number)/self.d[i]
                    r=r+1
                else:
                    ba1[i]=r
                    self.e[i]=v1[ba1[i]-1]
                    v1[r]=v1[ba1[i]-1]
                    r=r+1
            else:
                if h2[i]==0:
                    ba2[i]=l
                    self.e[i]=self.OrginDate[i].price
                    v2[l]=self.OrginDate[i].price
                    l=l+1
                elif self.d[i]==0:
                    ba2[i]=l
                    self.e[i]=0
                    v2[l]=0
                    l=l+1
                elif d2[h2[i]-1]==0:
                    ba2[i]=l
                    self.e[i]=self.OrginDate[i].price
                    v2[l]=self.OrginDate[i].price
                    l=l+1
                elif b[i]==1:
                    ba2[i]=l
                    self.e[i]=(v2[ba2[i]-1]*d2[h2[i]-1]+self.OrginDate[i].price*self.OrginDate[i].number)/self.d[i]
                    v2[l]=(v2[ba2[i]-1]*d2[h2[i]-1]+self.OrginDate[i].price*self.OrginDate[i].number)/self.d[i]
                    l=l+1
                else:
                    ba2[i]=l
                    self.e[i]=v2[ba2[i]-1]
                    v2[l]=v2[ba2[i]-1]
                    l=l+1

        z=0
        m=0
        j=0
        k=0
        h1={}
        h2={}
        d1={}
        d2={}
        self.d={}
        for i in range(self.rows+self.rows2):
            if self.n[i]==1:
                z=z+self.c[i]
                h1[i]=j
                d1[j]=z
                self.d[i]=z
                j=j+1
            else:
                m=self.c[i]+m
                h2[i]=k
                d2[k]=m
                self.d[i]=m
                k=k+1

        self.j={}
        for i in range(self.rows+self.rows2):
            if self.n[i]==1:
                if ba1[i]==0:
                    self.j[i]=0
                elif self.d[i]==0 or d1[h1[i]-1]==0:
                    self.j[i]=0
                elif b[i]==1:
                    self.j[i]=(self.OrginDate[i].price-v1[ba1[i]-1])*d1[h1[i]-1]*multiple
                else:
                    self.j[i]=(self.OrginDate[i].price-v1[ba1[i]-1])*self.d[i]*multiple

            else:
                if ba2[i]==0:
                    self.j[i]=0
                elif self.d[i]==0 or d2[h2[i]-1]==0:
                    self.j[i]=0
                elif b[i]==1:
                    self.j[i]=(self.OrginDate[i].price-v2[ba2[i]-1])*d2[h2[i]-1]*(-1)*multiple
                else:
                    self.j[i]=(self.OrginDate[i].price-v2[ba2[i]-1])*self.d[i]*(-1)*multiple

        self.r={}
        o={}
        s={}
        u1={}
        u2={}
        x1=0
        x2=0
        for i in range(self.rows+self.rows2):
            if self.n[i]==1:
                if ba1[i]==0:
                    u1[i]=x1
                    self.r[i]=0
                    o[x1]=0
                    x1=x1+1
                elif b1[i]==1:
                    u1[i]=x1
                    self.r[i]=o[u1[i]-1]
                    o[x1]=o[u1[i]-1]
                    x1=x1+1
                else:
                    u1[i]=x1
                    self.r[i]=o[u1[i]-1]+(self.OrginDate[i].price-v1[ba1[i]-1])*self.OrginDate[i].number*multiple
                    o[x1]=o[u1[i]-1]+(self.OrginDate[i].price-v1[ba1[i]-1])*self.OrginDate[i].number*multiple
                    x1=x1+1
            else:
                if ba2[i]==0:
                    u2[i]=x2
                    self.r[i]=0
                    s[x2]=0
                    x2=x2+1
                elif b1[i]==1:
                    u2[i]=x2
                    self.r[i]=s[u2[i]-1]
                    s[x2]=s[u2[i]-1]
                    x2=x2+1
                else:
                    u2[i]=x2
                    self.r[i]=s[u2[i]-1]+(self.OrginDate[i].price-v2[ba2[i]-1])*self.OrginDate[i].number*(-1)*multiple
                    s[x2]=s[u2[i]-1]+(self.OrginDate[i].price-v2[ba2[i]-1])*self.OrginDate[i].number*(-1)*multiple
                    x2=x2+1
        if x1>0:
            self.duo=o[x1-1]
        else:
            self.duo=0
        if x2>0:
            self.kong=s[x2-1]
        else:
            self.kong=0
        sum=0
        self.fin={}
        for i in range(self.rows+self.rows2):
            if i==0:
                self.fin[i]=0
            elif self.n[i]==self.n[i-1]:
                self.fin[i]=self.r[i]+self.j[i]+sum
            else:
                sum=self.r[i-1]+self.j[i-1]
                self.fin[i]=self.r[i]+self.j[i]+sum
        self.supfino={}
        for i in range(self.rows+self.rows2):
            self.supfino[i]=self.fin[i]-self.pouno[i]

        net=0
        self.net={}
        for i in range(self.rows+self.rows2):
            net=net+self.c[i]*self.n[i]
            self.net[i]=net

        self.num1={}
        for i in range(self.rows+self.rows2):
            if i==0:
                if self.net[i]>0:
                    self.num1[i]=self.OrginDate[i].number
                else:
                    self.num1[i]=0
            elif self.net[i]>0:
                if self.net[i]>0 and self.net[i-1]<0:
                    self.num1[i]=self.net[i-1]+self.OrginDate[i].number
                else:
                    self.num1[i]=self.OrginDate[i].number
            else:
                if self.net[i]<0 and self.net[i-1]<=0:
                    self.num1[i]=0
                elif self.net[i]<0 and self.net[i-1]>0:
                    self.num1[i]=self.net[i-1]

                elif self.net[i]==0 and self.net[i-1]<0:
                    self.num1[i]=0
                else:
                    self.num1[i]=self.OrginDate[i].number


        self.num2={}
        for i in range(self.rows+self.rows2):
            if i==0:
                if self.net[i]<0:
                    self.num2[i]=self.OrginDate[i].number
                else:
                    self.num2[i]=0
            elif self.net[i]<0:
                if self.net[i]<0 and self.net[i-1]>0:
                    self.num2[i]=-self.net[i-1]+self.OrginDate[i].number
                else:
                    self.num2[i]=self.OrginDate[i].number
            else:
                if self.net[i]>0 and self.net[i-1]>=0:
                    self.num2[i]=0
                elif self.net[i]>0 and self.net[i-1]<0:
                    self.num2[i]=-self.net[i-1]
                elif self.net[i]==0 and self.net[i-1]>0:
                    self.num2[i]=0
                else:
                    self.num2[i]=self.OrginDate[i].number
    def spduo(self):

        self.OrginDate1=[]
        for i in range(self.rows+self.rows2):
            self.OrginDate1.append(outnet(self.OrginDate[i].direction,self.OrginDate[i].offsetFlag,self.OrginDate[i].price,self.num1[i],self.OrginDate[i].time))
        self.OrginDate1.sort(key=lambda x:x.timedate.split(':'))
    def func1(self):
        ad={}
        for i in range(self.rows+self.rows2):
            if self.OrginDate1[i].direction == u''.encode("gbk"):
                ad[i]=1
            else:
                if self.OrginDate1[i].direction == u'买'.encode("gbk"):
                    ad[i]=1
                else:
                    ad[i]=-1

        bd={}
        for i in range(self.rows+self.rows2):
            if self.OrginDate1[i].offsetFlag== u'开仓'.encode("gbk"):
                bd[i]=1
            else:
                bd[i]=-1
        duonumber={}
        for i in range(self.rows+self.rows2):
            if self.OrginDate1[i].direction==u''.encode("gbk"):
                duonumber[i]=0
            else:
                duonumber[i]=self.OrginDate1[i].number
        self.cd={}
        for i in range(self.rows+self.rows2):
            if self.OrginDate1[i].direction == u''.encode("gbk"):
                self.cd[i]=0
            else:
                self.cd[i]=ad[i]*self.OrginDate1[i].number

        nd={}

        for i in range(self.rows+self.rows2):
            nd[i]=bd[i]*ad[i]

        zad=0
        self.dad={}
        for i in range(self.rows+self.rows2):
            zad=zad+self.cd[i]
            self.dad[i]=zad


        self.ed={}
        for i in range(0,self.rows+self.rows2):
            if i==0:
                self.ed[i]=self.OrginDate1[i].price
            elif self.dad[i]==0:
                self.ed[i]=0
            elif self.dad[i-1]==0:
                self.ed[i]=self.OrginDate1[i].price
            elif ad[i]==1:
                self.ed[i]=(self.ed[i-1]*self.dad[i-1]+self.OrginDate1[i].price*duonumber[i])/self.dad[i]
            else:
                self.ed[i]=self.ed[i-1]

        self.jd={}
        for i in range(0,self.rows+self.rows2):
            if i==0:
                self.jd[i]=0
            elif self.dad[i]==0 or self.dad[i-1]==0:
                self.jd[i]=0
            elif ad[i]==1:
                self.jd[i]=(self.OrginDate1[i].price-self.ed[i-1])*self.dad[i-1]*multiple
            else:
                self.jd[i]=(self.OrginDate1[i].price-self.ed[i-1])*self.dad[i-1]*multiple


        self.rd={}
        for i in range(0,self.rows+self.rows2):
            if i==0:
                self.rd[i]=0
            elif ad[i]==1:
                self.rd[i]=self.rd[i-1]
            else:
                self.rd[i]=self.rd[i-1]+(self.OrginDate1[i].price-self.ed[i-1])*self.OrginDate1[i].number*multiple


        self.find={}
        for i in range(0,self.rows+self.rows2):
            if i==0:
                self.find[i]=0
            else:
                self.find[i]=self.rd[i]+self.jd[i]
    def spduokong(self):
        self.final2=[]
        for i in range(self.rows+self.rows2):
            self.final2.append(OutputtradeRecord(self.cd[i],self.dad[i],self.ed[i],self.jd[i],self.rd[i],self.find[i]))
    def spkong(self):

        self.OrginDate2=[]
        for i in range(self.rows+self.rows2):
            self.OrginDate2.append(outnet(self.OrginDate[i].direction,self.OrginDate[i].offsetFlag,self.OrginDate[i].price,self.num2[i],self.OrginDate[i].time,))
        self.OrginDate2.sort(key=lambda x:x.timedate.split(':'))

    def func2(self):
        ak={}
        for i in range(self.rows+self.rows2):
            if self.OrginDate2[i].direction == u''.encode("gbk"):
                ak[i]=-1
            else:
                if self.OrginDate2[i].direction == u'买'.encode("gbk"):
                    ak[i]=1
                else:
                    ak[i]=-1
        bk={}
        for i in range(self.rows+self.rows2):
            if self.OrginDate2[i].offsetFlag== u'开仓'.encode("gbk"):
                bk[i]=1
            else:
                bk[i]=-1

        self.ck={}
        for i in range(self.rows+self.rows2):
            if self.OrginDate2[i].direction == u''.encode("gbk"):
                self.ck[i]=0
            else:
                self.ck[i]=ak[i]*self.OrginDate2[i].number

        nk={}
        for i in range(self.rows+self.rows2):
            nk[i]=bk[i]*ak[i]

        zak=0
        self.dak={}
        for i in range(self.rows+self.rows2):
            zak=zak+self.ck[i]
            self.dak[i]=-zak
        kongnumber={}
        for i in range(self.rows+self.rows2):
            if self.OrginDate2[i].direction==u''.encode("gbk"):
                kongnumber[i]=0
            else:
                kongnumber[i]=self.OrginDate2[i].number

        self.ek={}
        for i in range(0,self.rows+self.rows2):
            if i==0:
                self.ek[i]=self.OrginDate2[i].price
            elif self.dak[i]==0:
                self.ek[i]=0
            elif self.dak[i-1]==0:
                self.ek[i]=self.OrginDate2[i].price
            elif ak[i]==-1:
                self.ek[i]=(self.ek[i-1]*self.dak[i-1]+self.OrginDate2[i].price*kongnumber[i])/self.dak[i]
            else:
                self.ek[i]=self.ek[i-1]

        self.jk={}
        for i in range(0,self.rows+self.rows2):
            if i==0:
                self.jk[i]=0
            elif self.dak[i]==0 or self.dak[i-1]==0:
                self.jk[i]=0
            elif ak[i]==-1:
                self.jk[i]=(self.OrginDate2[i].price-self.ek[i-1])*self.dak[i-1]*multiple*(-1)
            else:
                self.jk[i]=(self.OrginDate2[i].price-self.ek[i-1])*self.dak[i-1]*multiple*(-1)

        self.rk={}
        for i in range(0,self.rows+self.rows2):
            if i==0:
                self.rk[i]=0
            elif ak[i]==-1:
                self.rk[i]=self.rk[i-1]
            else:
                self.rk[i]=self.rk[i-1]+(self.OrginDate2[i].price-self.ek[i-1])*self.OrginDate2[i].number*multiple*(-1)


        self.fink={}
        for i in range(0,self.rows+self.rows2):
            if i==0:
                self.fink[i]=0
            else:
                self.fink[i]=self.rk[i]+self.jk[i]
    def spduokong2(self):
        self.final3=[]
        for i in range(self.rows+self.rows2):
            self.final3.append(OutputtradeRecord(self.ck[i],self.dak[i],self.ek[i],self.jk[i],self.rk[i],self.fink[i]))
    def finalresult(self,strfilename):
        final=[]


        ping[strfilename]=self.fin[self.rows+self.rows2-1]
        for i in range(self.rows+self.rows2):
            final.append(OutputTradeRecord(self.c[i],self.d[i],self.e[i],self.j[i],self.r[i],self.fin[i],self.net[i],self.pouno[i],self.supfino[i]))
            sql="""
                insert into {0} (
                timedate,price,net,result,poundage,pureresult,duonet,duomovement,kongnet,kongmovement)
                values('{1}',{2},{3},{4},{5},{6},{7},{8},{9},{10})
                """.format('out_'+strfilename,self.OrginDate[i].time,self.OrginDate[i].price,final[i].duokong,final[i].amount,final[i].poundage,final[i].finalfinal,self.final2[i].amount,self.final2[i].movement,self.final3[i].amount,self.final3[i].movement)
            cursor3.execute(sql)
            out.commit()
    def everyone(self,k,strfilename):
        c=0
        duocount=0
        kongcount=0
        for i in range(self.rows):
            c=c+self.data1[i][3]
        for i in range(self.rows+self.rows2):
            if self.n[i]==1 and self.OrginDate[i].number!='':
                duocount=duocount+self.OrginDate[i].number
            elif self.n[i]==-1 and self.OrginDate[i].number!='':
                kongcount=kongcount+self.OrginDate[i].number
        sql="""
            insert into [statement(name)] (
            date,pingprofit,volume,pingprofitduo,pingprofitkong,volumeduo,volumekong,tradeday)
            values('{0}',{1},{2},{3},{4},{5},{6},{7})
            """.format(strfilename,ping[strfilename],c,self.duo,self.kong,duocount,kongcount,len(self.daylist))
        cursor3.execute(sql)
        out.commit()  
if __name__=='__main__':

    allsheet=[]
    existsheet=[]


    calount=cursor4.execute("select count(*) from sysobjects where xtype='U'").fetchone()
    print calount[0]
    sheetname=cursor4.execute("select name from sysobjects where xtype='U'").fetchall()
    for i in range(calount[0]):
        allsheet.append(sheetname[i][0])
    print allsheet
    t=Methods()

    sql="delete from [statement(name)]"
    cursor3.execute(sql)
    out.commit()
    ping={}
    existname=cursor3.execute("select name from sysobjects where xtype='U'").fetchall()
    row=cursor3.execute("select count(*) from sysobjects where xtype='U'").fetchone()
    row=row[0]
    for i in range(row):
        if "out" in existname[i][0]:
            existsheet.append(existname[i][0].split('_')[1])
    for k in range(calount[0]):
        strfilename=allsheet[k].encode("gbk")
        t.readExcel(strfilename)
        t.readExcel2()



        if allsheet[k] not in existsheet:

            sql="""
            create table {0}(
            timedate varchar(50),
            price decimal(20, 4),
            result decimal(20, 4),
            net decimal(20, 4),
            poundage decimal(20, 4),
            pureresult decimal(20, 4),
            duonet decimal(20, 4),
            duomovement decimal(20, 4),
            kongnet decimal(20, 4),
            kongmovement decimal(20, 4)
            )
            """.format('out_'+strfilename)
            cursor3.execute(sql)
            out.commit()
        else:
            sql="delete from {0}".format('out_'+strfilename)
            cursor3.execute(sql)
            out.commit()
        t.deal()
        t.spduo()
        t.func1()
        t.spduokong()
        t.spkong()
        t.func2()
        t.spduokong2()
        t.finalresult(strfilename)
        t.everyone(k,strfilename)
