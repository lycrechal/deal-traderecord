#!/usr/bin/env python
# -*- coding: utf-8 -*-

import pandas as pd
import pyodbc
import datetime
import time
import os
multiple=int(raw_input("Y(^o^)Yenter the times "))
poundage=int(raw_input("Y(^o^)Yenter the poundage "))

tr= pyodbc.connect('DRIVER={SQL Server};SERVER=120.24.68.150,1453;DATABASE=TradeData;UID=dbUser;PWD=db+123-456')
cursor1 = tr.cursor()
outname=pyodbc.connect('DRIVER={SQL Server};SERVER=120.24.68.150,1453;DATABASE=outname;UID=dbUser;PWD=db+123-456')
cursor2=outname.cursor()
out = pyodbc.connect('DRIVER={SQL Server};SERVER=120.24.68.150,1453;DATABASE=out;UID=dbUser;PWD=db+123-456')
cursor3 = out.cursor()
class TradeRecord(object):
    def __init__(self,name,direction,offsetflag,price,number,timedate):
        self.name=name
        self.direction=direction
        self.offsetflag=offsetflag
        self.price=price
        self.number=number
        self.timedate=timedate
class OutPut(object):
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
class Methods(object):
    def loaddata2(self):
        self.OrginDate=[]
        self.daylist=[]
        self.namelist=[]
        
        self.rows=len(df)
        print "%d" %(self.rows)
        for i in range(self.rows):
            self.OrginDate.append(TradeRecord(df.name.values[i],df.direction.values[i],df.offsetflag.values[i],df.price.values[i],df.volume.values[i],df.date.values[i]+" "+df.time.values[i]))
            self.daylist.append(df.date.values[i])
            self.namelist.append(df.name.values[i])
            
        self.daylist=list(set(self.daylist))
        self.namelist=list(set(self.namelist))

        self.OrginDate.sort(key=lambda x:(x.name,x.timedate.split(':')))
        
        self.finalrows=len(self.OrginDate)
    def loaddata(self):
        self.OrginDate=[]
        self.daylist=[]
        self.namelist=[]
        
        self.rows=len(df)
        print "%d" %(self.rows)
        for i in range(self.rows):
            self.OrginDate.append(TradeRecord(df.name.values[i],df.direction.values[i],df.offsetflag.values[i],df.price.values[i],df.volume.values[i],df.date.values[i]+" "+df.time.values[i]))
            self.daylist.append(df.date.values[i])
            self.namelist.append(df.name.values[i])
            
        self.daylist=sorted(list(set(self.daylist)))
        self.namelist=sorted(list(set(self.namelist)))

        self.OrginDate.sort(key=lambda x:x.timedate.split(':'))
        
        self.finalrows=len(self.OrginDate)
    def loadmin(self):
        self.rows2=len(g_MinData)
        for i in range(self.rows2):
            if g_MinData.datetime.values[i].split(' ')[0] in self.daylist:
                if int(g_MinData.datetime.values[i].split(' ')[1].split(':')[0])<17:
                    if int(g_MinData.datetime.values[i].split(' ')[1].split(':')[0])==9:
                        if int(g_MinData.datetime.values[i].split(' ')[1].split(':')[1])>14:
                            _record = TradeRecord('','','',g_MinData.price.values[i],'',g_MinData.datetime.values[i])
                            self.OrginDate.append(_record)
                    else:
                        _record = TradeRecord('','','',g_MinData.price.values[i],'',g_MinData.datetime.values[i])
                        self.OrginDate.append(_record)


        self.OrginDate.sort(key=lambda x:x.timedate.split(':'))
        self.finalrows=len(self.OrginDate)
    def attendtime(self,k,strfilename):
        self.number={}
        date1=self.OrginDate[0].timedate
        kkk=0
        ddd=1
        ppp=1
        volume=0
        tdp=[]
        for i in range(self.finalrows):

            date2=self.OrginDate[i].timedate

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
                tdp.append(float(self.fin[i]-self.fin[kkk]))
                date1=self.OrginDate[i].timedate
                kkk=i
                ddd=ddd+1
                ppp=ppp+1
                volume=0
            else:
                if es>900:
                    tdp.append(float(self.fin[i]-self.fin[kkk]))

                    date1=self.OrginDate[i].timedate
                    kkk=i
                    ddd=ddd+1
                    ppp=ppp+1
                    volume=0
                else:
                    volume=volume+self.number[i]
        tdp.append(float(self.fin[self.finalrows-1]-self.fin[kkk]))
        while len(tdp)<24:
            tdp.append(0)
        sql="""
                insert into [timedivide] (
                date,[1],[2],[3],[4],[5],[6],[7],[8],[9],[10],[11],[12],[13],[14],[15],[16],[17],[18],[19],[20],[21],[22],[23],[24])
                values('{0}',{1},{2},{3},{4},{5},{6},{7},{8},{9},{10},{11},{12},{13},{14},{15},{16},{17},{18},{19},{20},{21},{22},{23},{24})
                """.format(strfilename,tdp[0],tdp[1],tdp[2],tdp[3],tdp[4],tdp[5],tdp[6],tdp[7],tdp[8],tdp[9],tdp[10],tdp[11],tdp[12],tdp[13],tdp[14],tdp[15],tdp[16],tdp[17],tdp[18],tdp[19],tdp[20],tdp[21],tdp[22],tdp[23])
        cursor3.execute(sql)
        out.commit() 

    def deal(self,last1,last2):
        a={}
        for i in range(self.finalrows):
            if self.OrginDate[i].direction == u'买'.encode("gbk"):
                a[i]=1
            else:
                a[i]=-1
        b={}
        for i in range(self.finalrows):
            if self.OrginDate[i].offsetflag== u'开仓'.encode("gbk"):
                b[i]=1
            else:
               b[i]=-1

        b1={}
        for i in range(self.finalrows):
            if self.OrginDate[i].direction==u'':
                b1[i]=1
            else:
                if self.OrginDate[i].offsetflag == u'开仓'.encode("gbk"):
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
        self.pouno[0]=abs(self.c[0])*poundage+last2
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
        if len(d1)==0:
            self.maxduo=0
        else:
            self.maxduo=max(d1.items(), key=lambda x: x[1])[1]
        if len(d2)==0:
            self.maxkong=0
        else:
            self.maxkong=max(d2.items(), key=lambda x: x[1])[1]
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
                self.fin[i]=last1
            elif self.n[i]==self.n[i-1]:
                self.fin[i]=self.r[i]+self.j[i]+sum+last1
            else:
                sum=self.r[i-1]+self.j[i-1]
                self.fin[i]=self.r[i]+self.j[i]+sum+last1

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
    def namediff(self,k):
        sum=0
        self.namefin={}
        self.volume=df.volume.values[0]
        diffday=0
        for i in range(self.finalrows):
            if i==0:
                self.namefin[i]=0
            elif self.n[i]==self.n[i-1]:
                self.volume=self.volume+df.volume.values[i]
                self.namefin[i]=self.r[i]+self.j[i]+sum
                if self.OrginDate[i].timedate.split(' ')[0]!=self.OrginDate[i-1].timedate.split(' ')[0]:
                    
                    sql="""
                        insert into [{0}] (
                        timedate,result,volume)values('{1}',{2},{3})
                        """.format(strfilename,self.OrginDate[i-1].timedate.split(' ')[0],self.namefin[i]-diffday,self.volume)
                    cursor2.execute(sql)
                    outname.commit()
                    self.volume=0
                    diffday=self.namefin[i]
            else:
                self.volume=self.volume+df.volume.values[i]
                sum=self.r[i-1]+self.j[i-1]
                self.namefin[i]=self.r[i]+self.j[i]+sum
                if self.OrginDate[i].timedate.split(' ')[0]!=self.OrginDate[i-1].timedate.split(' ')[0]:
                    sql="""
                        insert into [{0}] (
                        timedate,result,volume)values('{1}',{2},{3})
                        """.format(strfilename,self.OrginDate[i-1].timedate.split(' ')[0],self.namefin[i]-diffday,self.volume)
                    cursor2.execute(sql)
                    outname.commit()
                    self.volume=0
                    diffday=self.namefin[i]
        sql="""
                insert into [{0}] (
                timedate,result,volume)values('{1}',{2},{3})
                """.format(strfilename,self.OrginDate[i-1].timedate.split(' ')[0],self.namefin[i]-diffday,self.volume)
        cursor2.execute(sql)
        outname.commit()
    def daydiff(self,k):
        sum=0
        self.dayfin={}
        self.charge={}

        diffday=0
        charge=0
        for i in range(self.finalrows):
            if i==0:
                self.dayfin[i]=0
                self.charge[i]=0
            elif self.n[i]==self.n[i-1]:
                self.dayfin[i]=self.r[i]+self.j[i]+sum
                self.charge[i]=self.pouno[i]
                if self.OrginDate[i].name!=self.OrginDate[i-1].name:
                    
                    sql="""
                        insert into [{0}] (
                        timedate,result,poundage)values('{1}',{2},{3})
                        """.format(strfilename,self.OrginDate[i-1].name,self.dayfin[i]-diffday,self.charge[i]-charge)
                    cursor2.execute(sql)
                    outname.commit()
                    diffday=self.dayfin[i]
                    charge=self.charge[i]
            else:
                sum=self.r[i-1]+self.j[i-1]
                self.dayfin[i]=self.r[i]+self.j[i]+sum
                self.charge[i]=self.pouno[i]
                if self.OrginDate[i].name!=self.OrginDate[i-1].name:
                    sql="""
                        insert into [{0}] (
                        timedate,result,poundage)values('{1}',{2},{3})
                        """.format(strfilename,self.OrginDate[i-1].name,self.dayfin[i]-diffday,self.charge[i]-charge)
                    cursor2.execute(sql)
                    outname.commit()
                    diffday=self.dayfin[i]
                    charge=self.charge[i]
        sql="""
                insert into [{0}] (
                timedate,result,poundage)values('{1}',{2},{3})
                """.format(strfilename,self.OrginDate[i-1].name,self.dayfin[i]-diffday,self.charge[i]-charge)
        cursor2.execute(sql)
        outname.commit()
    def spduo(self):

        self.OrginDate1=[]
        for i in range(self.finalrows):
            self.OrginDate1.append(TradeRecord('',self.OrginDate[i].direction,self.OrginDate[i].offsetflag,self.OrginDate[i].price,self.num1[i],self.OrginDate[i].timedate))
        self.OrginDate1.sort(key=lambda x:x.timedate.split(':'))
        
    def func1(self,last3):
        ad={}
        for i in range(self.finalrows):
            if self.OrginDate1[i].direction == u'':
                ad[i]=1
            else:
                if self.OrginDate1[i].direction == u'买'.encode("gbk"):
                    ad[i]=1
                else:
                    ad[i]=-1

        bd={}
        for i in range(self.finalrows):
            if self.OrginDate1[i].offsetflag== u'开仓'.encode("gbk"):
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
        self.poundageduo={}
        dage=0
        self.poundageduo[0]=abs(self.cd[0])*poundage
        for i in range(1,self.finalrows):
            dage=abs(self.cd[i])*poundage
            self.poundageduo[i]=self.poundageduo[i-1]+dage
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
                self.find[i]=last3
            else:
                self.find[i]=self.rd[i]+self.jd[i]+last3
        self.supfin={}
        for i in range(self.finalrows):
            self.supfin[i]=self.find[i]-self.poundageduo[i]
    def spduokong(self):
        self.final2=[]
        for i in range(self.finalrows):
            self.final2.append(OutPut(self.cd[i],'',self.ed[i],self.jd[i],self.rd[i],self.dad[i],'','',self.find[i]))

    def netduo(self):
        final2=[]
        for i in range(self.finalrows):
            final2.append(OutPut(self.cd[i],'',self.ed[i],self.jd[i],self.rd[i],-self.find[i],self.dad[i],self.poundageduo[i],-self.find[i]-self.poundageduo[i]))
            sql="""
                insert into {0} (
                timedate,price,net,result,poundage,pureresult)
                values('{1}',{2},{3},{4},{5},{6})
                """.format('netduo',self.OrginDate1[i].timedate,self.OrginDate1[i].price,final2[i].duokong,final2[i].amount,final2[i].poundage,final2[i].finalfinal)
            cursor3.execute(sql)
            out.commit() 

    def spkong(self):

        self.OrginDate2=[]
        for i in range(self.finalrows):
            self.OrginDate2.append(TradeRecord('',self.OrginDate[i].direction,self.OrginDate[i].offsetflag,self.OrginDate[i].price,self.num2[i],self.OrginDate[i].timedate))
        self.OrginDate2.sort(key=lambda x:x.timedate.split(':'))

    def func2(self,last4):
        ak={}
        for i in range(self.finalrows):
            if self.OrginDate2[i].direction == u'':
                ak[i]=-1
            else:
                if self.OrginDate2[i].direction == u'买'.encode("gbk"):
                    ak[i]=1
                else:
                    ak[i]=-1
        bk={}
        for i in range(self.finalrows):
            if self.OrginDate2[i].offsetflag== u'开仓'.encode("gbk"):
                bk[i]=1
            else:
                bk[i]=-1

        self.ck={}
        for i in range(self.finalrows):
            if self.OrginDate2[i].direction == u'':
                self.ck[i]=0
            else:
                self.ck[i]=ak[i]*self.OrginDate2[i].number
        self.poundagekong={}
        dage=0
        self.poundagekong[0]=abs(self.ck[0])*poundage
        for i in range(1,self.finalrows):
            dage=abs(self.ck[i])*poundage
            self.poundagekong[i]=self.poundagekong[i-1]+dage
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
                self.fink[i]=0+last4
            else:
                self.fink[i]=self.rk[i]+self.jk[i]+last4
        self.supfin={}
        for i in range(self.finalrows):
            self.supfin[i]=self.fink[i]-self.poundagekong[i]
    def spduokong2(self):
        self.final3=[]
        for i in range(self.finalrows):
            self.final3.append(OutPut(self.ck[i],'',self.ek[i],self.jk[i],self.rk[i],self.dak[i],'','',self.fink[i]))

    def netkong(self):
        final3=[]
        for i in range(self.finalrows):
            final3.append(OutPut(self.ck[i],'',self.ek[i],self.jk[i],self.rk[i],-self.fink[i],self.dak[i],self.poundagekong[i],-self.fink[i]-self.poundagekong[i]))
            sql="""
                insert into {0} (
                timedate,price,net,result,poundage,pureresult)
                values('{1}',{2},{3},{4},{5},{6})
                """.format('netkong',self.OrginDate2[i].timedate,self.OrginDate2[i].price,final3[i].duokong,final3[i].amount,final3[i].poundage,final3[i].finalfinal)
            cursor3.execute(sql)
            out.commit() 

    def runtime(self):
        time1=self.OrginDate[0].timedate
        self.alltime=0
        for i in range(self.finalrows):
            if self.net[i]==0:
                time2=self.OrginDate[i].timedate
                a=time.strptime(time1, "%Y%m%d %H:%M:%S")
                b=time.strptime(time2, "%Y%m%d %H:%M:%S")
                starttime=datetime.datetime(a[0],a[1],a[2],a[3],a[4],a[5])
                endtime=datetime.datetime(b[0],b[1],b[2],b[3],b[4],b[5])
                es=(endtime-starttime).seconds
                self.alltime=es+self.alltime
                if i<self.finalrows-1:
                    time1=self.OrginDate[i+1].timedate
                else:
                    time1=time2
    def finalresult(self,strfilename):
        final=[]
        ping[strfilename]=self.fin[self.finalrows-1]
        for i in range(self.finalrows):
            final.append(OutPut(self.c[i],self.d[i],self.e[i],self.j[i],self.r[i],self.fin[i],self.net[i],self.pouno[i],self.supfino[i]))
            sql="""
                insert into [{0}] (
                timedate,price,net,result,poundage,pureresult,duonet,duomovement,kongnet,kongmovement)
                values('{1}',{2},{3},{4},{5},{6},{7},{8},{9},{10})
                """.format('out_'+strfilename,self.OrginDate[i].timedate,self.OrginDate[i].price,final[i].duokong,final[i].amount,final[i].poundage,final[i].finalfinal,self.final2[i].amount,self.final2[i].finalfinal,self.final3[i].amount,self.final3[i].finalfinal)
            cursor3.execute(sql)
            out.commit()
    def finalresult2(self,strfilename,sheetname):
        final=[]
        ping[strfilename]=self.fin[self.finalrows-1]
        for i in range(self.finalrows):
            final.append(OutPut(self.c[i],self.d[i],self.e[i],self.j[i],self.r[i],self.fin[i],self.net[i],self.pouno[i],self.supfino[i]))
            sql="""
                insert into [{0}] (
                timedate,price,net,result,poundage,pureresult,duonet,duomovement,kongnet,kongmovement,name)
                values('{1}',{2},{3},{4},{5},{6},{7},{8},{9},{10},'{11}')
                """.format('out_'+sheetname,self.OrginDate[i].timedate,self.OrginDate[i].price,final[i].duokong,final[i].amount,final[i].poundage,final[i].finalfinal,self.final2[i].amount,self.final2[i].finalfinal,self.final3[i].amount,self.final3[i].finalfinal,strfilename)
            cursor3.execute(sql)
            out.commit()
    def everyone(self,k,strfilename,NOD):
        c=0
        duocount=0
        kongcount=0

        for i in range(self.rows):
            c=c+df.volume.values[i]

        for i in range(self.finalrows):
            if self.n[i]==1 and self.OrginDate[i].number!='':
                duocount=duocount+self.OrginDate[i].number
            elif self.n[i]==-1 and self.OrginDate[i].number!='':
                kongcount=kongcount+self.OrginDate[i].number


        if NOD==0:
            self.instrument=[]
            for i in range(self.rows):
                self.instrument.append(df.instrument.values[i])
            self.instrument=list(set(self.instrument))
            self.instrument = ','.join(self.instrument)
            sql="""
            insert into [statement(name)] (
            date,pingprofit,volume,maxduo,maxkong,maxprofit,minprofit,tradeday,start,[end],tradeplace,instrument)
            values('{0}',{1},{2},{3},{4},{5},{6},{7},'{8}','{9}','{10}','{11}')
            """.format(strfilename,ping[strfilename],c,self.maxduo,self.maxkong,max(self.fin.items(), key=lambda x: x[1])[1],min(self.fin.items(), key=lambda x: x[1])[1],len(self.daylist),self.daylist[0],self.daylist[-1],df.tradeplace.values[0],self.instrument)
            cursor3.execute(sql)
            out.commit() 
        elif NOD==3:

            sql="""
            update [statement(name)] set
            [pingprofit]={1},[volume]=[volume]+{2},[tradeday]=[tradeday]+{3},[end]='{4}',[tradeplace]='{5}'
            where [date]='{0}'
            
            """.format(strfilename,ping[strfilename],c,len(self.daylist),self.daylist[-1],df.tradeplace.values[0])
            cursor3.execute(sql)
            out.commit()             
        else:
            sql="""
            insert into [statement(date)] (
            date,pingprofit,volume,maxduo,maxkong,maxprofit,minprofit,tradercount,runtime)
            values('{0}',{1},{2},{3},{4},{5},{6},{7},{8})
            """.format(strfilename,self.fin[self.finalrows-1],c,self.maxduo,self.maxkong,max(self.fin.items(), key=lambda x: x[1])[1],min(self.fin.items(), key=lambda x: x[1])[1],len(self.namelist),self.alltime)
            cursor3.execute(sql)
            out.commit() 
    def updateeveryone(self,strfilename,NOD):
        c=0
        for i in range(self.rows):
            c=c+df.volume.values[i]
        if NOD==0:
            sql="""
            update [statement(name)]  SET [simulated profit] = {1},[simulated volume]={2},[simulated daycount]={3} WHERE [date] = '{0}'
            """.format(strfilename,self.fin[self.finalrows-1],c,len(self.daylist))
            cursor3.execute(sql)
            out.commit() 
        else:
            sql="""
            update [statement(name)]  SET [actual profit] = {1},[actual volume]={2},[actual daycount]={3} WHERE [date] = '{0}'
            """.format(strfilename,self.fin[self.finalrows-1],c,len(self.daylist))
            cursor3.execute(sql)
            out.commit() 
if __name__=='__main__':
    sql="select [name],[direction],[offsetflag],[price],[volume],[time],[date],[tradeplace],[instrument] from [new]"
    g_TradeData=pd.read_sql(sql,tr)

    mydist=sorted(list(set(g_TradeData['date'])))
    print mydist
    calculatedate=Methods()
    ping={}
    calount=cursor3.execute("select count(*) from sysobjects where xtype='U'").fetchone()
    print calount[0]
    allsheet=[]
    sheetname=cursor3.execute("select name from sysobjects where xtype='U'").fetchall()
    for i in range(calount[0]):
        allsheet.append(sheetname[i][0])
    for k in range(len(mydist)):
        if 'out_'+mydist[k] not in allsheet:

            df=g_TradeData[(g_TradeData['date']==mydist[k])].loc[:,['name','direction','offsetflag','price','volume','time','date']]
            strfilename=str(mydist[k])
            calculatedate.loaddata2()
            #calculatedate.loadmin()
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
            sql="""
            create table [{0}](timedate varchar(50),
            result decimal(20,4),
            poundage decimal(20,4))
            """.format(strfilename)
            cursor2.execute(sql)
            outname.commit()

            calculatedate.deal(0,0)
            calculatedate.daydiff(k)
            calculatedate.loaddata()
            calculatedate.deal(0,0)
            calculatedate.spduo()
            calculatedate.func1(0)
            calculatedate.spduokong()
            calculatedate.spkong()
            calculatedate.func2(0)
            calculatedate.spduokong2()
            calculatedate.finalresult(strfilename)
            calculatedate.runtime()
            calculatedate.everyone(k,strfilename,2)
            calculatedate.attendtime(k,strfilename)
    calculateall=Methods()
    df=g_TradeData.loc[:,['name','direction','offsetflag','price','volume','time','date']]
    ping={}
    last=cursor3.execute("SELECT top 1 [result],[poundage],[duomovement],[kongmovement]FROM [out].[dbo].[out_summary statement]order by timedate desc,net asc").fetchall()
    last1=float(last[0][0])
    last2=float(last[0][1])
    last3=float(last[0][2])
    last4=float(last[0][3])
    calculateall.loaddata()
    #calculateall.loadmin()

    calculateall.deal(last1,last2)

    calculateall.spduo()
    calculateall.func1(last3)
    calculateall.spduokong()
    calculateall.spkong()
    calculateall.func2(last4)
    calculateall.spduokong2()
    calculateall.finalresult('summary statement')
    #calculateall.spduo()
    #calculateall.func1()
    #calculateall.netduo()
    #calculateall.spkong()
    #calculateall.func2()
    #calculateall.netkong()

    mylist=list(set(g_TradeData['name']))
    print mylist
    calculatename=Methods()
    df=g_TradeData.loc[:,['name','direction','offsetflag','price','volume','time','date','tradeplace','instrument']]
    ping={}
    namecount=cursor2.execute("select count(*) from sysobjects where xtype='U'").fetchone()
    namesheet=[]
    sheetname=cursor2.execute("select name from sysobjects where xtype='U'").fetchall()
    for i in range(namecount[0]):
        namesheet.append(sheetname[i][0])

    for k in range(len(mylist)):
        strfilename=mylist[k]
        if(mylist[k].decode("gbk")) not in namesheet:
            sql="""
            create table {0}(timedate varchar(50),
            result decimal(20,4),
            volume int)
            """.format(strfilename)
            cursor2.execute(sql)
            outname.commit()
            df=g_TradeData[(g_TradeData['name']==mylist[k])].loc[:,['name','direction','offsetflag','price','volume','time','date','tradeplace','instrument']]
            calculatename.loaddata()
            #calculatename.loadmin()
            calculatename.deal(0,0)
            calculatename.namediff(k)
            calculatename.spduo()
            calculatename.func1(0)
            calculatename.spduokong()
            calculatename.spkong()
            calculatename.func2(0)
            calculatename.spduokong2()
            calculatename.finalresult2(strfilename,'summary(name)')
            calculatename.everyone(k,strfilename,0)           
        else:
            df=g_TradeData[(g_TradeData['name']==mylist[k])].loc[:,['name','direction','offsetflag','price','volume','time','date','tradeplace','instrument']]
            sql="""
        SELECT top 1 [result],[poundage],[duomovement],[kongmovement]
        FROM [out].[dbo].[out_summary(name)]
        where [name]= '{0}'
        order by timedate desc,net asc
        """.format(strfilename)
            last=cursor3.execute(sql).fetchall()
        
            last1=float(last[0][0])
            last2=float(last[0][1])
            last3=float(last[0][2])
            last4=float(last[0][3])
            calculatename.loaddata()
            #calculatename.loadmin()
            calculatename.deal(last1,last2)
            calculatename.namediff(k)
            calculatename.spduo()
            calculatename.func1(last3)
            calculatename.spduokong()
            calculatename.spkong()
            calculatename.func2(last4)
            calculatename.spduokong2()
            calculatename.finalresult2(strfilename,'summary(name)')
            calculatename.everyone(k,strfilename,3)
        

