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
multiple=int(raw_input("Y(^o^)Yenter the times "))
poundage=int(raw_input("Y(^o^)Yenter the poundage "))

class OperExcel():
    def rExcel(self):
        file_path2=('C://Users//Administrator//new//all.xls')
        self.rb=open_workbook(file_path2,formatting_info=True)

        self.rs=self.rb.sheet_by_index(0)

        self.rows=self.rs.nrows
        cols=self.rs.ncols

        for i in range(self.rows):
            mylist.append(str(int(self.rs.cell(i,6).value)))

        list(set(mylist))
        oldwb=xlwt.Workbook()
        oldws=oldwb.add_sheet('Sheet1')
        rows=self.rs.nrows
        for i in range(rows):
            oldws.write(i,0,self.rs.cell(i,1).value)
            oldws.write(i,1,self.rs.cell(i,2).value)
            oldws.write(i,2,self.rs.cell(i,3).value)
            oldws.write(i,3,self.rs.cell(i,4).value)
            oldws.write(i,4,str(int(self.rs.cell(i,6).value))+' '+str(self.rs.cell(i,5).value))

        oldwb.save('C://Users//Administrator//new//alldate.xls')


    def deal(self):

        self.ws=newwb.add_sheet(strfilename)

        for i in range(self.rows):
            global ca
            if str(int(self.rs.cell(i,6).value))==strfilename:
                self.ws.write(ca,0,self.rs.cell(i,1).value)
                self.ws.write(ca,1,self.rs.cell(i,2).value)
                self.ws.write(ca,2,self.rs.cell(i,3).value)
                self.ws.write(ca,3,self.rs.cell(i,4).value)
                self.ws.write(ca,4,str(int(self.rs.cell(i,6).value))+' '+str(self.rs.cell(i,5).value))
                self.ws.write(ca,5,self.rs.cell(i,0).value)

                ca=ca+1




mylist=[]
t=OperExcel()
t.rExcel()
newlist=sorted(list(set(mylist)))
newwb=xlwt.Workbook()

for j in range(0,len(newlist)):

    strfilename=newlist[j]
    ca=0
    t.deal()
savepth=('C://Users//Administrator//new//outdate.xls')
newwb.save(savepth)



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
        self.namelist=[]
        self.rs=rb.sheet_by_name(strfilename)
        self.wb=copy(rb)

        self.rows=self.rs.nrows
        cols=self.rs.ncols
        print "%d %d" %(self.rows,cols)
        for i in range(self.rows):
            _record = TradeRecord(self.rs.cell(i,4).value,self.rs.cell(i,1).value,self.rs.cell(i,0).value,self.rs.cell(i,2).value,self.rs.cell(i,3).value)
            self.OrginDate.append(_record)
            self.daylist.append(self.rs.cell(i,4).value.split(' ')[0])
            self.daylist=list(set(self.daylist))
            self.namelist.append(self.rs.cell(i,5).value)
        self.namelist=list(set(self.namelist))


    def readExcel2(self,strfilename):

        file_path=('C://Users//Administrator//new//min.xls')
        rb2=open_workbook(file_path,formatting_info=True)
        self.rs2=rb2.sheet_by_index(0)
        self.wb2=copy(rb2)
        self.ws2=self.wb2.get_sheet(0)
        self.rows2=self.rs2.nrows
        for i in range(self.rows2):
            if self.rs2.cell(i,1).value.split(' ')[0] in self.daylist:
                if self.rs2.cell(i,1).value.split(' ')[0]==strfilename:

                    if int(self.rs2.cell(i,1).value.split(' ')[1].split(':')[0])<17:
                        if int(self.rs2.cell(i,1).value.split(' ')[1].split(':')[0])==9:
                            if int(self.rs2.cell(i,1).value.split(' ')[1].split(':')[1])>14:
                                _record = TradeRecord(self.rs2.cell(i,1).value.split('\x00')[0],'','',self.rs2.cell(i,7).value,'')
                                self.OrginDate.append(_record)
                        else:
                            _record = TradeRecord(self.rs2.cell(i,1).value.split('\x00')[0],'','',self.rs2.cell(i,7).value,'')
                            self.OrginDate.append(_record)


        self.OrginDate.sort(key=lambda x:x.time.split(':'))

    def readExcel3(self,strfilename):

        file_path3=('C://Users//Administrator//new//another.xls')
        rb3=open_workbook(file_path3,formatting_info=True)
        rs3=rb3.sheet_by_index(0)


        self.rows3=rs3.nrows

        for i in range(self.rows3):
            if str(int(rs3.cell(i,6).value))==strfilename:
                self.OrginDate.append(TradeRecord(str(int(rs3.cell(i,6).value))+' '+str(rs3.cell(i,5).value),'','',rs3.cell(i,3).value,''))
        self.OrginDate.sort(key=lambda x:x.time.split(':'))
        self.finalrows=len(self.OrginDate)
    def attendtime(self,k):
        self.number={}
        date1=self.OrginDate[0].time
        kkk=0
        ddd=1
        ppp=1
        volume=0
        for i in range(self.finalrows):

            date2=self.OrginDate[i].time

            a=time.strptime(date1, "%Y%m%d %H:%M:%S")
            b=time.strptime(date2, "%Y%m%d %H:%M:%S")
            starttime=datetime.datetime(a[0],a[1],a[2],a[3],a[4],a[5])
            endtime=datetime.datetime(b[0],b[1],b[2],b[3],b[4],b[5])
            es=(endtime-starttime).seconds
            if self.OrginDate[i].number==u'':
                self.number[i]=0
            else:
                self.number[i]=self.OrginDate[i].number
            if es==900:
                newws6.write(k+1,ppp,volume)
                newws5.write(k+1,ddd,self.fin[i]-self.fin[kkk])
                date1=self.OrginDate[i].time
                kkk=i
                ddd=ddd+1
                ppp=ppp+1
                volume=0
            else:
                if es>900:
                    newws6.write(k+1,ppp,volume)
                    newws5.write(k+1,ddd,self.fin[i]-self.fin[kkk])
                    date1=self.OrginDate[i].time
                    kkk=i
                    ddd=ddd+1
                    ppp=ppp+1
                    volume=0
                else:
                    volume=volume+self.number[i]

        newws5.write(k+1,ddd,self.fin[self.finalrows-1]-self.fin[kkk])
        newws6.write(k+1,ppp,volume)




    def deal(self):
        a={}
        for i in range(self.finalrows):
            if self.OrginDate[i].direction == u'买':
                a[i]=1
            else:
                a[i]=-1
        b={}
        for i in range(self.finalrows):
            if self.OrginDate[i].offsetFlag== u'开仓':
                b[i]=1
            else:
               b[i]=-1

        b1={}
        for i in range(self.finalrows):
            if self.OrginDate[i].direction==u'':
                b1[i]=1
            else:
                if self.OrginDate[i].offsetFlag == u'开仓':
                    b1[i]=1
                else:
                    b1[i]=-1
        self.n={}
        if self.OrginDate[0].direction==u'':

            self.n[0]=1
        else:
            self.n[0]=a[0]*b[0]
        for i in range(1,self.finalrows):
            if self.OrginDate[i].direction==u'':
                self.n[i]=self.n[i-1]
            else:
                self.n[i]=a[i]*b[i]

        abc={}
        for i in range(self.finalrows):
            if self.OrginDate[i].direction==u'':
                abc[i]=0
            else:
                abc[i]=self.OrginDate[i].number

        self.c={}
        for i in range(self.finalrows):
            if self.OrginDate[i].direction == u'':
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
            self.supfino[i]=-self.fin[i]-self.pouno[i]

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
            if self.OrginDate1[i].direction == u'':
                ad[i]=1
            else:
                if self.OrginDate1[i].direction == u'买':
                    ad[i]=1
                else:
                    ad[i]=-1

        bd={}
        for i in range(self.finalrows):
            if self.OrginDate1[i].offsetFlag== u'开仓':
                bd[i]=1
            else:
                bd[i]=-1
        duonumber={}
        for i in range(self.finalrows):
            if self.OrginDate1[i].direction==u'':
                duonumber[i]=0
            else:
                duonumber[i]=self.OrginDate1[i].number
        self.cd={}
        for i in range(self.finalrows):
            if self.OrginDate1[i].direction == u'':
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
        final2=[]
        for i in range(self.finalrows):
            final2.append(OutputtradeRecord(self.cd[i],self.dad[i],self.ed[i],self.jd[i],self.rd[i],self.find[i]))

            #newws.write(i,12,final2[i].change)
            newws.write(i,6,final2[i].amount)
            #newws.write(i,14,final2[i].avgprice)
            #newws.write(i,15,final2[i].position)
            #newws.write(i,16,final2[i].closeprofit)
            newws.write(i,7,final2[i].movement)

    def spkong(self):

        self.OrginDate2=[]
        for i in range(self.finalrows):
            self.OrginDate2.append(outnet(self.OrginDate[i].direction,self.OrginDate[i].offsetFlag,self.OrginDate[i].price,self.num2[i],self.OrginDate[i].time,))
        self.OrginDate2.sort(key=lambda x:x.timedate.split(':'))

    def func2(self):
        ak={}
        for i in range(self.finalrows):
            if self.OrginDate2[i].direction == u'':
                ak[i]=-1
            else:
                if self.OrginDate2[i].direction == u'买':
                    ak[i]=1
                else:
                    ak[i]=-1
        bk={}
        for i in range(self.finalrows):
            if self.OrginDate2[i].offsetFlag== u'开仓':
                bk[i]=1
            else:
                bk[i]=-1

        self.ck={}
        for i in range(self.finalrows):
            if self.OrginDate2[i].direction == u'':
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
            if self.OrginDate2[i].direction==u'':
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
        final3=[]
        for i in range(self.finalrows):
            final3.append(OutputtradeRecord(self.ck[i],self.dak[i],self.ek[i],self.jk[i],self.rk[i],self.fink[i]))

            #newws.write(i,18,final3[i].change)
            newws.write(i,8,final3[i].amount)
            #newws.write(i,20,final3[i].avgprice)
            #newws.write(i,21,final3[i].position)
            #newws.write(i,22,final3[i].closeprofit)
            newws.write(i,9,final3[i].movement)

    def finalresult(self,strfilename):
        final=[]

        noon=0
        ping[strfilename]=self.fin[self.finalrows-1]
        for i in range(self.finalrows):
            if int(self.OrginDate[i].time.split(' ')[1].split(':')[0])<12:
                noon=i
            final.append(OutputTradeRecord(self.c[i],self.d[i],self.e[i],self.j[i],self.r[i],self.fin[i],self.net[i],self.pouno[i],self.supfino[i]))
            #newws.write(i,0,self.OrginDate[i].direction)
            #newws.write(i,1,self.OrginDate[i].offsetFlag)
            newws.write(i,1,self.OrginDate[i].price)
            #newws.write(i,3,self.OrginDate[i].number)
            newws.write(i,0,self.OrginDate[i].time)
            #newws.write(i,5,final[i].change)
            #newws.write(i,6,final[i].openinterest)
            #newws.write(i,7,final[i].avgprice)
            #newws.write(i,8,final[i].avgprofit)
            #newws.write(i,9,final[i].pingprofit)
            newws.write(i,2,final[i].amount)
            newws.write(i,3,final[i].duokong)
            newws.write(i,4,final[i].poundage)
            newws.write(i,5,final[i].finalfinal)
        print noon
        if noon==0:
            self.am=[0]
            self.pm=self.supfino.values()
        else:
            self.am=self.supfino.values()[0:noon+1]
            self.pm=self.supfino.values()[noon+1:]
    def everyone(self,k,strfilename):
        c=0
        duocount=0
        kongcount=0
        newws2.write(k+1,0,strfilename)
        newws2.write(k+1,1,ping[strfilename])
        for i in range(self.rows):
            c=c+self.rs.cell(i,3).value
        newws2.write(k+1,2,c)
        newws2.write(k+1,3,self.duo)
        newws2.write(k+1,4,self.kong)
        for i in range(self.finalrows):
            if self.n[i]==1 and self.OrginDate[i].number!='':
                duocount=duocount+self.OrginDate[i].number
            elif self.n[i]==-1 and self.OrginDate[i].number!='':
                kongcount=kongcount+self.OrginDate[i].number
        newws2.write(k+1,5,duocount)
        newws2.write(k+1,6,kongcount)
        newws2.write(k+1,7,len(self.namelist))
    
    def kline(self,k):
        if len(self.pm)==0:
            self.pm=[0]
            newws3.write(2*k+2,2,max(self.am))
            newws3.write(2*k+2,3,min(self.am))
            newws3.write(2*k+2,4,self.am[-1])
            newws3.write(2*k+2,1,0)
            newws3.write(2*k+3,2,0)
            newws3.write(2*k+3,3,0)
            newws3.write(2*k+3,4,0)
            newws3.write(2*k+3,1,0)
        #newws3.write(k+2,1,max(self.fin.items(), key=lambda x: x[1])[1])
        #newws3.write(k+2,2,min(self.fin.items(), key=lambda x: x[1])[1])
        #newws3.write(k+2,3,self.fin[max(self.fin)])
        #newws3.write(k+3,0,self.fin[max(self.fin)])
        #newws3.write(k+2,9,max(self.fin.items(), key=lambda x: x[1])[1])
        #newws3.write(k+2,10,min(self.fin.items(), key=lambda x: x[1])[1])
        #newws3.write(k+2,11,self.fin[max(self.fin)])
        #newws3.write(k+2,8,0)
            newws3.write(2*k+2,7,(max(self.am))/(len(self.namelist)))
            newws3.write(2*k+2,8,(min(self.am))/(len(self.namelist)))
            newws3.write(2*k+2,9,self.am[-1]/(len(self.namelist)))
            newws3.write(2*k+2,6,0)
            newws3.write(2*k+3,7,0)
            newws3.write(2*k+3,8,0)
            newws3.write(2*k+3,9,0)
            newws3.write(2*k+3,6,0)
        else:
            newws3.write(2*k+2,2,max(self.am))
            newws3.write(2*k+2,3,min(self.am))
            newws3.write(2*k+2,4,self.am[-1])
            newws3.write(2*k+2,1,0)
            newws3.write(2*k+3,2,max(self.pm)-self.am[-1])
            newws3.write(2*k+3,3,min(self.pm)-self.am[-1])
            newws3.write(2*k+3,4,self.pm[-1]-self.am[-1])
            newws3.write(2*k+3,1,0)
            newws3.write(2*k+2,7,(max(self.am))/(len(self.namelist)))
            newws3.write(2*k+2,8,(min(self.am))/(len(self.namelist)))
            newws3.write(2*k+2,9,self.am[-1]/(len(self.namelist)))
            newws3.write(2*k+2,6,0)
            newws3.write(2*k+3,7,(max(self.pm)-self.am[-1])/(len(self.namelist)))
            newws3.write(2*k+3,8,(min(self.pm)-self.am[-1])/(len(self.namelist)))
            newws3.write(2*k+3,9,(self.pm[-1]-self.am[-1])/(len(self.namelist)))
            newws3.write(2*k+3,6,0)
        newws3.write(2*k+2,0,strfilename+'A')
        newws3.write(2*k+3,0,strfilename+'B')
if __name__=='__main__':

    allsheet=[]
    file_path=('C://Users//Administrator//new//outdate.xls')
    rb=open_workbook(file_path,formatting_info=True)
    calount=len(rb.sheets())
    print calount
    for sheet in rb.sheets():
        allsheet.append(sheet.name)
    print allsheet
    t=Methods()

    newwb=xlwt.Workbook()
    newws2=newwb.add_sheet(u'结算单（日期）')
    newws3=newwb.add_sheet(u'开低高收')
    newws2.write(0,0,u'日期')
    newws2.write(0,1,u'平仓盈亏')
    newws2.write(0,2,u'成交量')
    newws2.write(0,3,u'平仓盈亏多头')
    newws2.write(0,4,u'平仓盈亏空头')
    newws2.write(0,5,u'成交量多头')
    newws2.write(0,6,u'成交量空头')
    newws2.write(0,7,u'交易人数')

    ping={}
    newws5=newwb.add_sheet(u'时间段结算(盈亏)')
    newws6=newwb.add_sheet(u'时间段结算（成交量）')
    newws5.write(0,0,u'日期')
    newws6.write(0,0,u'日期')
    newws5.write(0,1,u'09:15:00-09:30:00')
    newws6.write(0,1,u'09:15:00-09:30:00')
    newws5.write(0,2,u'09:30:00-09:45:00')
    newws6.write(0,2,u'09:30:00-09:45:00')
    newws5.write(0,3,u'09:45:00-10:00:00')
    newws6.write(0,3,u'09:45:00-10:00:00')
    newws5.write(0,4,u'10:00:00-10:15:00')
    newws6.write(0,4,u'10:00:00-10:15:00')
    newws5.write(0,5,u'10:15:00-10:30:00')
    newws6.write(0,5,u'10:15:00-10:30:00')
    newws5.write(0,6,u'10:30:00-10:45:00')
    newws6.write(0,6,u'10:30:00-10:45:00')
    newws5.write(0,7,u'10:45:00-11:00:00')
    newws6.write(0,7,u'10:45:00-11:00:00')
    newws5.write(0,8,u'11:00:00-11:15:00')
    newws6.write(0,8,u'11:00:00-11:15:00')
    newws5.write(0,9,u'11:15:00-11:30:00')
    newws6.write(0,9,u'11:15:00-11:30:00')
    newws5.write(0,10,u'11:30:00-11:45:00')
    newws6.write(0,10,u'11:30:00-11:45:00')
    newws5.write(0,11,u'11:45:00-12:00:00')
    newws6.write(0,11,u'11:45:00-12:00:00')
    newws5.write(0,12,u'13:00:00-13:15:00')
    newws6.write(0,12,u'13:00:00-13:15:00')
    newws5.write(0,13,u'13:15:00-13:30:00')
    newws6.write(0,13,u'13:15:00-13:30:00')
    newws5.write(0,14,u'13:30:00-13:45:00')
    newws6.write(0,14,u'13:30:00-13:45:00')
    newws5.write(0,15,u'13:45:00-14:00:00')
    newws6.write(0,15,u'13:45:00-14:00:00')
    newws5.write(0,16,u'14:00:00-14:15:00')
    newws6.write(0,16,u'14:00:00-14:15:00')
    newws5.write(0,17,u'14:15:00-14:30:00')
    newws6.write(0,17,u'14:15:00-14:30:00')
    newws5.write(0,18,u'14:30:00-14:45:00')
    newws6.write(0,18,u'14:30:00-14:45:00')
    newws5.write(0,19,u'14:45:00-15:00:00')
    newws6.write(0,19,u'14:45:00-15:00:00')
    newws5.write(0,20,u'15:00:00-15:15:00')
    newws6.write(0,20,u'15:00:00-15:15:00')
    newws5.write(0,21,u'15:15:00-15:30:00')
    newws6.write(0,21,u'15:15:00-15:30:00')
    newws5.write(0,22,u'15:30:00-15:45:00')
    newws6.write(0,22,u'15:30:00-15:45:00')
    newws5.write(0,23,u'15:45:00-16:00:00')
    newws6.write(0,23,u'15:45:00-16:00:00')
    newws5.write(0,24,u'16:00:00-16:15:00')
    newws6.write(0,24,u'16:00:00-16:15:00')
    for k in range(calount):
        strfilename=allsheet[k]



        t.readExcel(strfilename)
        t.readExcel2(strfilename)
        newws=newwb.add_sheet(strfilename)
        t.readExcel3(strfilename)


        newws5.write(k+1,0,strfilename)
        newws6.write(k+1,0,strfilename)
        t.deal()

        t.finalresult(strfilename)

        t.spduo()
        t.func1()
        t.spduokong()
        t.spkong()
        t.func2()
        t.spduokong2()
        t.everyone(k,strfilename)
        t.kline(k)
        t.attendtime(k)

    savepth=('C://Users//Administrator//new//out.xls')
    newwb.save(savepth)


class OperExcel():
    def rExcel(self):
        file_path=('C://Users//Administrator//new//all.xls')
        rb=open_workbook(file_path,formatting_info=True)

        rs=rb.sheet_by_index(0)
        oldwb=xlwt.Workbook()
        oldws=oldwb.add_sheet("Sheet1")
        rows=rs.nrows
        for i in range(rows):
            oldws.write(i,0,rs.cell(i,1).value)
            oldws.write(i,1,rs.cell(i,2).value)
            oldws.write(i,2,rs.cell(i,3).value)
            oldws.write(i,3,rs.cell(i,4).value)
            oldws.write(i,4,str(int(rs.cell(i,6).value))+' '+rs.cell(i,5).value)
            oldws.write(i,5,rs.cell(i,0).value)
        oldwb.save('C://Users//Administrator//new//allname.xls')
        file_path2=('C://Users//Administrator//new//allname.xls')
        self.rb=open_workbook(file_path2,formatting_info=True)

        self.rs=self.rb.sheet_by_index(0)

        self.rows=self.rs.nrows
        cols=self.rs.ncols

        for i in range(self.rows):
            mylist.append(self.rs.cell(i,5).value)

        list(set(mylist))
    def deal(self):

        self.ws=newwb.add_sheet(strfilename)

        for i in range(self.rows):
            global ca
            if self.rs.cell(i,5).value==strfilename:
                self.ws.write(ca,0,self.rs.cell(i,0).value)
                self.ws.write(ca,1,self.rs.cell(i,1).value)
                self.ws.write(ca,2,self.rs.cell(i,2).value)
                self.ws.write(ca,3,self.rs.cell(i,3).value)
                self.ws.write(ca,4,self.rs.cell(i,4).value)

                ca=ca+1




mylist=[]
t=OperExcel()
t.rExcel()

newwb=xlwt.Workbook()

for j in range(0,len(list(set(mylist)))):

    strfilename=list(set(mylist))[j]
    ca=0
    t.deal()
savepth=('C://Users//Administrator//new//outperson.xls')
newwb.save(savepth)





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

        file_path=('C://Users//Administrator//new//alldate.xls')
        rb=open_workbook(file_path,formatting_info=True)
        rs=rb.sheet_by_index(0)
        self.wb=xlwt.Workbook()
        self.ws=self.wb.add_sheet(u'总表')
        file_path2=('C://Users//Administrator//new//outperson.xls')
        rb2=open_workbook(file_path2,formatting_info=True)
        self.wb2=copy(rb2)
        self.ws2=self.wb2.add_sheet(u'多头')
        self.ws3=self.wb2.add_sheet(u'空头')
        self.rows=rs.nrows
        cols=rs.ncols
        print "%d %d" %(self.rows,cols)
        for i in range(self.rows):
            _record = TradeRecord(rs.cell(i,1).value,rs.cell(i,0).value,rs.cell(i,2).value,rs.cell(i,3).value,rs.cell(i,4).value)
            OrginDate.append(_record)
        OrginDate.sort(key=lambda x:x.time.split(':'))
        for i in range(self.rows):
            self.ws.write(i,0,OrginDate[i].direction)

            self.ws.write(i,1,OrginDate[i].offsetFlag)
            self.ws.write(i,2,OrginDate[i].price)
            self.ws.write(i,3,OrginDate[i].number)
            self.ws.write(i,4,OrginDate[i].time)

    def deal(self):
        a={}
        for i in range(self.rows):
            if OrginDate[i].direction== u'买':
                a[i]=1
            else:
                a[i]=-1

        b={}
        for i in range(self.rows):
            if OrginDate[i].offsetFlag == u'开仓':
                b[i]=1
            else:
                b[i]=-1
        self.c={}
        for i in range(self.rows):
            self.c[i]=b[i]*OrginDate[i].number
        self.n={}
        for i in range(self.rows):
            self.n[i]=a[i]*b[i]

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
            elif a[i]*b[i]==a[i-1]*b[i-1]:
                self.fin[i]=self.r[i]+self.j[i]+sum
            else:
                sum=self.r[i-1]+self.j[i-1]
                self.fin[i]=self.r[i]+self.j[i]+sum


    def finalresult(self):
        for i in range(self.rows):
            final.append(OutputTradeRecord(self.c[i],self.d[i],self.e[i],self.j[i],self.r[i],self.fin[i],self.n[i]))
            self.ws.write(i,5,final[i].change)
            self.ws.write(i,6,final[i].openinterest)
            self.ws.write(i,7,final[i].avgprice)
            self.ws.write(i,8,final[i].avgprofit)
            self.ws.write(i,9,final[i].pingprofit)
            self.ws.write(i,10,final[i].amount)
            self.ws.write(i,11,final[i].duokong)

    def addduokong(self):
        ca=0
        pa=0
        for i in range(self.rows):
            if self.n[i]==1:

                duo.append(OutputTradeRecord(self.c[i],self.d[i],self.e[i],self.j[i],self.r[i],self.fin[i],self.n[i]))
                self.ws2.write(ca,0,OrginDate[i].direction)
                self.ws2.write(ca,1,OrginDate[i].offsetFlag)
                self.ws2.write(ca,2,OrginDate[i].price)
                self.ws2.write(ca,3,OrginDate[i].number)
                self.ws2.write(ca,4,OrginDate[i].time)
                self.ws2.write(ca,5,duo[ca].change)
                self.ws2.write(ca,6,duo[ca].openinterest)
                self.ws2.write(ca,7,duo[ca].avgprice)
                self.ws2.write(ca,8,duo[ca].avgprofit)
                self.ws2.write(ca,9,duo[ca].pingprofit)
                self.ws2.write(ca,10,duo[ca].amount)
                self.ws2.write(ca,11,duo[ca].duokong)
                self.ws2.write(ca,12,duo[ca].avgprofit+duo[ca].pingprofit)
                ca=ca+1
            else:
                kong.append(OutputTradeRecord(self.c[i],self.d[i],self.e[i],self.j[i],self.r[i],self.fin[i],self.n[i]))
                self.ws3.write(pa,0,OrginDate[i].direction)
                self.ws3.write(pa,1,OrginDate[i].offsetFlag)
                self.ws3.write(pa,2,OrginDate[i].price)
                self.ws3.write(pa,3,OrginDate[i].number)
                self.ws3.write(pa,4,OrginDate[i].time)
                self.ws3.write(pa,5,kong[pa].change)
                self.ws3.write(pa,6,kong[pa].openinterest)
                self.ws3.write(pa,7,kong[pa].avgprice)
                self.ws3.write(pa,8,kong[pa].avgprofit)
                self.ws3.write(pa,9,kong[pa].pingprofit)
                self.ws3.write(pa,10,kong[pa].amount)
                self.ws3.write(pa,11,kong[pa].duokong)
                self.ws3.write(pa,12,kong[pa].avgprofit+kong[pa].pingprofit)
                pa=pa+1

        self.wb.save('C://Users//Administrator//new//outall.xls')
        self.wb2.save('C://Users//Administrator//new//outperson.xls')
t=Methods()
t.readExcel()
t.deal()
t.finalresult()
t.addduokong()




OrginDate = []



class TradeRecord(object):
    def __init__(self,time,offsetFlag,direction,price,number):
        self.time=str(time)
        self.offsetFlag = offsetFlag
        self.direction = direction
        self.price = price
        self.number = number

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





class outnet(object):
    def __init__(self,direction,offsetFlag,price,number,time):
        self.direction=direction
        self.offsetFlag=offsetFlag
        self.price=price
        self.number=number
        self.time=time


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




class deal(object):
    def readExcel2(self):
        self.ordata=[]
        self.OrginDate=[]
        self.OrginDate2=[]
        file_path=('C://Users//Administrator//new//min.xls')
        rb2=open_workbook(file_path,formatting_info=True)
        rs2=rb2.sheet_by_index(0)
        self.wb2=copy(rb2)
        self.ws2=self.wb2.get_sheet(0)
        self.rows2=rs2.nrows
        print "%d" %(self.rows2)
        for i in range(self.rows2):
            if rs2.cell(i,1).value.split(' ')[0] in self.daylist:

                if int(rs2.cell(i,1).value.split(' ')[1].split(':')[0])<17:
                    _record = TradeRecord(rs2.cell(i,1).value,'','',rs2.cell(i,7).value,'')
                    self.OrginDate.append(_record)
                    self.OrginDate2.append(_record)
        self.OrginDate.sort(key=lambda x:x.time.split(':'))
        self.OrginDate2.sort(key=lambda x:x.time.split(':'))
        self.rows2=len(self.OrginDate)

    def readExcel(self):
        file_path=('C://Users//Administrator//new//alldate.xls')
        rb=open_workbook(file_path,formatting_info=True)
        rs=rb.sheet_by_index(0)
        file_path2=('C://Users//Administrator//new//out.xls')
        rb2=open_workbook(file_path2,formatting_info=True)
        self.daylist=[]
        self.wb=copy(rb2)
        self.ws=self.wb.add_sheet(u'总表')
        #self.ws2=self.wb.add_sheet(u'开多')
        #self.ws3=self.wb.add_sheet(u'开空')
        self.ws4=self.wb.add_sheet(u'净多')
        self.ws5=self.wb.add_sheet(u'净空')
        #self.ws6=self.wb.add_sheet(u'持仓量盈亏')
        #self.ws7=self.wb.add_sheet(u'持仓量成交量')
        #self.ws8=self.wb.add_sheet(u'持仓量累计时间')
        self.rows=rs.nrows
        for i in range(self.rows):
            _record1 = TradeRecord(rs.cell(i,4).value,rs.cell(i,1).value,rs.cell(i,0).value,rs.cell(i,2).value,rs.cell(i,3).value)
            OrginDate.append(_record1)
            self.daylist.append(rs.cell(i,4).value.split(' ')[0])
        self.daylist=list(set(self.daylist))
        OrginDate.sort(key=lambda x:x.time.split(':'))

    def deal2(self):

        self.a={}
        for i in range(self.rows):
            if OrginDate[i].direction== u'买':
                self.a[i]=1
            else:
                self.a[i]=-1
        b={}
        for i in range(self.rows):
            if OrginDate[i].offsetFlag == u'开仓':
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
            #self.ws.write(i,0,OrginDate[i].direction)
            #self.ws.write(i,1,OrginDate[i].offsetFlag)
            self.ws.write(i,1,OrginDate[i].price)
            #self.ws.write(i,3,OrginDate[i].number)
            self.ws.write(i,0,OrginDate[i].time)
            #self.ws.write(i,5,final[i].change)
            #self.ws.write(i,6,final[i].openinterest)
            #self.ws.write(i,7,final[i].avgprice)
            #self.ws.write(i,8,final[i].avgprofit)
            #self.ws.write(i,9,final[i].pingprofit)
            self.ws.write(i,2,final[i].amount)
            #self.ws.write(i,11,final[i].duokong)
            self.ws.write(i,3,final[i].net)
            self.ws.write(i,4,final[i].poundage)
            self.ws.write(i,5,final[i].finalfinal)

        
    def addchicangliang(self):
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
        for p in range(10):
            self.ws6.write(0,p,write[p])
            self.ws7.write(0,p,down[p])
            self.ws8.write(0,p,timetime[p])
   


    def spduo(self):
        for i in range(self.rows):
            self.OrginDate.append(outnet(OrginDate[i].direction,OrginDate[i].offsetFlag,OrginDate[i].price,self.num1[i],OrginDate[i].time))
        self.OrginDate.sort(key=lambda x:x.time.split(':'))
    def func1(self):
        ad={}
        for i in range(self.rows+self.rows2):
            if self.OrginDate[i].direction == u'':
                ad[i]=1
            else:
                if self.OrginDate[i].direction == u'买':
                    ad[i]=1
                else:
                    ad[i]=-1

        bd={}
        for i in range(self.rows+self.rows2):
            if self.OrginDate[i].offsetFlag== u'开仓':
                bd[i]=1
            else:
                bd[i]=-1
        abc={}
        for i in range(self.rows+self.rows2):
            if self.OrginDate[i].direction==u'':
                abc[i]=0
            else:
                abc[i]=self.OrginDate[i].number

        self.cd={}
        for i in range(self.rows+self.rows2):
            if self.OrginDate[i].direction == u'':
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
        for i in range(self.rows+self.rows2):
            #self.ws4.write(i,0,self.OrginDate[i].direction)
            #self.ws4.write(i,1,self.OrginDate[i].offsetFlag)
            #self.ws4.write(i,2,self.OrginDate[i].number)
            self.ws4.write(i,1,self.OrginDate[i].price)
            self.ws4.write(i,0,self.OrginDate[i].time)
            #self.ws4.write(i,5,final2[i].change)
            self.ws4.write(i,2,final2[i].amount)
            #self.ws4.write(i,7,final2[i].avgprice)
            #self.ws4.write(i,8,final2[i].position)
            #self.ws4.write(i,9,final2[i].closeprofit)
            self.ws4.write(i,3,final2[i].movement)
            self.ws4.write(i,4,final2[i].poundage)
            self.ws4.write(i,5,final2[i].finalfinal)

    def spkong(self):
        for i in range(self.rows):
            self.OrginDate2.append(outnet(OrginDate[i].direction,OrginDate[i].offsetFlag,OrginDate[i].price,self.num2[i],OrginDate[i].time,))
        self.OrginDate2.sort(key=lambda x:x.time.split(':'))

    def func2(self):
        ak={}
        for i in range(self.rows+self.rows2):
            if self.OrginDate2[i].direction == u'':
                ak[i]=-1
            else:
                if self.OrginDate2[i].direction == u'买':
                    ak[i]=1
                else:
                    ak[i]=-1
        bk={}
        for i in range(self.rows+self.rows2):
            if self.OrginDate2[i].offsetFlag== u'开仓':
                bk[i]=1
            else:
                bk[i]=-1

        self.ck={}
        for i in range(self.rows+self.rows2):
            if self.OrginDate2[i].direction == u'':
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
            if self.OrginDate2[i].direction==u'':
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
        for i in range(self.rows+self.rows2):
            #self.ws5.write(i,0,self.OrginDate2[i].direction)
            #self.ws5.write(i,1,self.OrginDate2[i].offsetFlag)
            #self.ws5.write(i,2,self.OrginDate2[i].number)
            self.ws5.write(i,1,self.OrginDate2[i].price)
            self.ws5.write(i,0,self.OrginDate2[i].time)
            #self.ws5.write(i,5,final3[i].change)
            self.ws5.write(i,2,final3[i].amount)
            #self.ws5.write(i,7,final3[i].avgprice)
            #self.ws5.write(i,8,final3[i].position)
            #self.ws5.write(i,9,final3[i].closeprofit)
            self.ws5.write(i,3,final3[i].movement)
            self.ws5.write(i,4,final3[i].poundage)
            self.ws5.write(i,5,final3[i].finalfinal)
        self.wb.save('C://Users//Administrator//new//out.xls')

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
        self.rs=rb.sheet_by_name(strfilename)
        self.wb=copy(rb)

        self.rows=self.rs.nrows
        cols=self.rs.ncols
        print "%d %d" %(self.rows,cols)
        for i in range(self.rows):
            _record = TradeRecord(self.rs.cell(i,4).value,self.rs.cell(i,1).value,self.rs.cell(i,0).value,self.rs.cell(i,2).value,self.rs.cell(i,3).value)
            self.OrginDate.append(_record)
            self.daylist.append(self.rs.cell(i,4).value.split(' ')[0])
            self.daylist=list(set(self.daylist))


    def readExcel2(self):

        file_path=('C://Users//Administrator//new//min.xls')
        rb2=open_workbook(file_path,formatting_info=True)
        self.rs2=rb2.sheet_by_index(0)
        self.wb2=copy(rb2)
        self.ws2=self.wb2.get_sheet(0)
        self.rows2=self.rs2.nrows
        for i in range(self.rows2):
            if self.rs2.cell(i,1).value.split(' ')[0] in self.daylist:

                if int(self.rs2.cell(i,1).value.split(' ')[1].split(':')[0])<17:
                    _record = TradeRecord(self.rs2.cell(i,1).value,'','',self.rs2.cell(i,7).value,'')
                    self.OrginDate.append(_record)
        self.OrginDate.sort(key=lambda x:x.time.split(':'))
        self.rows2=len(self.OrginDate)-self.rows


    def deal(self):
        a={}
        for i in range(self.rows+self.rows2):
            if self.OrginDate[i].direction == u'买':
                a[i]=1
            else:
                a[i]=-1
        b={}
        for i in range(self.rows+self.rows2):
            if self.OrginDate[i].offsetFlag== u'开仓':
                b[i]=1
            else:
               b[i]=-1

        b1={}
        for i in range(self.rows+self.rows2):
            if self.OrginDate[i].direction==u'':
                b1[i]=1
            else:
                if self.OrginDate[i].offsetFlag == u'开仓':
                    b1[i]=1
                else:
                    b1[i]=-1
        self.n={}
        if self.OrginDate[0].direction==u'':
            self.n[0]=1
        else:
            self.n[0]=a[0]*b[0]
        for i in range(1,self.rows+self.rows2):
            if self.OrginDate[i].direction==u'':
                self.n[i]=self.n[i-1]
            else:
                self.n[i]=a[i]*b[i]

        abc={}
        for i in range(self.rows+self.rows2):
            if self.OrginDate[i].direction==u'':
                abc[i]=0
            else:
                abc[i]=self.OrginDate[i].number

        self.c={}
        for i in range(self.rows+self.rows2):
            if self.OrginDate[i].direction == u'':
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
            if self.OrginDate1[i].direction == u'':
                ad[i]=1
            else:
                if self.OrginDate1[i].direction == u'买':
                    ad[i]=1
                else:
                    ad[i]=-1

        bd={}
        for i in range(self.rows+self.rows2):
            if self.OrginDate1[i].offsetFlag== u'开仓':
                bd[i]=1
            else:
                bd[i]=-1
        duonumber={}
        for i in range(self.rows+self.rows2):
            if self.OrginDate1[i].direction==u'':
                duonumber[i]=0
            else:
                duonumber[i]=self.OrginDate1[i].number
        self.cd={}
        for i in range(self.rows+self.rows2):
            if self.OrginDate1[i].direction == u'':
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
        final2=[]
        for i in range(self.rows+self.rows2):
            final2.append(OutputtradeRecord(self.cd[i],self.dad[i],self.ed[i],self.jd[i],self.rd[i],self.find[i]))
            #newws.write(i,12,final2[i].change)
            newws.write(i,6,final2[i].amount)
            #newws.write(i,14,final2[i].avgprice)
            #newws.write(i,15,final2[i].position)
            #newws.write(i,16,final2[i].closeprofit)
            newws.write(i,7,final2[i].movement)

    def spkong(self):

        self.OrginDate2=[]
        for i in range(self.rows+self.rows2):
            self.OrginDate2.append(outnet(self.OrginDate[i].direction,self.OrginDate[i].offsetFlag,self.OrginDate[i].price,self.num2[i],self.OrginDate[i].time,))
        self.OrginDate2.sort(key=lambda x:x.timedate.split(':'))

    def func2(self):
        ak={}
        for i in range(self.rows+self.rows2):
            if self.OrginDate2[i].direction == u'':
                ak[i]=-1
            else:
                if self.OrginDate2[i].direction == u'买':
                    ak[i]=1
                else:
                    ak[i]=-1
        bk={}
        for i in range(self.rows+self.rows2):
            if self.OrginDate2[i].offsetFlag== u'开仓':
                bk[i]=1
            else:
                bk[i]=-1

        self.ck={}
        for i in range(self.rows+self.rows2):
            if self.OrginDate2[i].direction == u'':
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
            if self.OrginDate2[i].direction==u'':
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
        final3=[]
        for i in range(self.rows+self.rows2):
            final3.append(OutputtradeRecord(self.ck[i],self.dak[i],self.ek[i],self.jk[i],self.rk[i],self.fink[i]))
            #newws.write(i,18,final3[i].change)
            newws.write(i,8,final3[i].amount)
            #newws.write(i,20,final3[i].avgprice)
            #newws.write(i,21,final3[i].position)
            #newws.write(i,22,final3[i].closeprofit)
            newws.write(i,9,final3[i].movement)

    def finalresult(self,strfilename):
        final=[]


        ping[strfilename]=self.fin[self.rows+self.rows2-1]
        for i in range(self.rows+self.rows2):
            final.append(OutputTradeRecord(self.c[i],self.d[i],self.e[i],self.j[i],self.r[i],self.fin[i],self.net[i],self.pouno[i],self.supfino[i]))
            #newws.write(i,0,self.OrginDate[i].direction)
            #newws.write(i,1,self.OrginDate[i].offsetFlag)
            newws.write(i,1,self.OrginDate[i].price)
            #newws.write(i,3,self.OrginDate[i].number)
            newws.write(i,0,self.OrginDate[i].time)
            #newws.write(i,5,final[i].change)
            #newws.write(i,6,final[i].openinterest)
            #newws.write(i,7,final[i].avgprice)
            #newws.write(i,8,final[i].avgprofit)
            #newws.write(i,9,final[i].pingprofit)
            newws.write(i,3,final[i].amount)
            newws.write(i,2,final[i].duokong)
            newws.write(i,4,final[i].poundage)
            newws.write(i,5,final[i].finalfinal)
    def runtime(self):
        time1=self.OrginDate[0].time
        self.alltime=0
        for i in range(self.rows+self.rows2):
            if self.net[i]==0:
                time2=self.OrginDate[i].time
                a=time.strptime(time1, "%Y%m%d %H:%M:%S")
                b=time.strptime(time2, "%Y%m%d %H:%M:%S")
                starttime=datetime.datetime(a[0],a[1],a[2],a[3],a[4],a[5])
                endtime=datetime.datetime(b[0],b[1],b[2],b[3],b[4],b[5])
                es=(endtime-starttime).seconds
                self.alltime=es+self.alltime
                if i<self.rows+self.rows2-1:
                    time1=self.OrginDate[i+1].time
                else:
                    time1=time2   
    def everyone(self,k,strfilename):
        c=0
        duocount=0
        kongcount=0
        newws2.write(k+1,0,strfilename)
        newws2.write(k+1,1,ping[strfilename])
        for i in range(self.rows):
            c=c+self.rs.cell(i,3).value
        newws2.write(k+1,2,c)
        newws2.write(k+1,3,self.duo)
        newws2.write(k+1,4,self.kong)
        for i in range(self.rows+self.rows2):
            if self.n[i]==1 and self.OrginDate[i].number!='':
                duocount=duocount+self.OrginDate[i].number
            elif self.n[i]==-1 and self.OrginDate[i].number!='':
                kongcount=kongcount+self.OrginDate[i].number
        newws2.write(k+1,5,duocount)
        newws2.write(k+1,6,kongcount)
        newws2.write(k+1,7,len(self.daylist))
        newws2.write(k+1,8,self.alltime)
if __name__=='__main__':

    allsheet=[]
    file_path=('C://Users//Administrator//new//outperson.xls')
    rb=open_workbook(file_path,formatting_info=True)
    calount=len(rb.sheets())
    print calount
    for sheet in rb.sheets():
        allsheet.append(sheet.name)
    print allsheet
    t=Methods()
    file_path2=('C://Users//Administrator//new//out.xls')
    rb2=open_workbook(file_path2,formatting_info=True)
    newwb=copy(rb2)

    newws2=newwb.add_sheet(u'结算单（姓名）')
    newws2.write(0,0,u'姓名')
    newws2.write(0,1,u'平仓盈亏')
    newws2.write(0,2,u'成交量')
    newws2.write(0,3,u'平仓盈亏多头')
    newws2.write(0,4,u'平仓盈亏空头')
    newws2.write(0,5,u'成交量多头')
    newws2.write(0,6,u'成交量空头')
    newws2.write(0,7,u'交易天数')
    newws2.write(0,8,u'持仓时间')
    ping={}
    for k in range(calount):
        strfilename=allsheet[k]



        t.readExcel(strfilename)
        t.readExcel2()
        newws=newwb.add_sheet(strfilename)
        t.deal()
        t.finalresult(strfilename)
        t.spduo()
        t.func1()
        t.spduokong()
        t.spkong()
        t.func2()
        t.spduokong2()
        t.runtime()
        t.everyone(k,strfilename)


    savepth=('C://Users//Administrator//new//out.xls')
    newwb.save(savepth)



