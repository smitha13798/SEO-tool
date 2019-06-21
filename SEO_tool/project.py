import re
from urllib.request import urlopen
from bs4 import BeautifulSoup
import xlsxwriter
from pyexcel_xlsx import read_data
import sqlite3
def readfromexcel(filename):
    mlist=[]
    data=read_data(filename) 
    for x in  data['Sheet1']:
         mlist.append(x)
    return mlist

def splitwords(l):
    wordlist=l.split(",")
    return wordlist



def searchfunc(pat,text):
      p=re.compile(pat)
      print(p)
      keylist=p.findall(text)
      print(keylist)
      num1=len(keylist)
      num2=len(text.split())
      freq=(num1/num2)
      return freq
    
def filterurl(url):
    data=urlopen(url)
    b=data.read()
    bs=BeautifulSoup(b,"html5lib")
    for x in bs(["script"],["style"]):
        x.extract()
    mydata=bs.get_text()
    print(mydata)
    return mydata

dbobject=''
def createdb():
     global dbobject
     dbobject=sqlite3.connect(":memory:")
     dbobject.execute("create table seo(url text,keyword text,density float)")
     
def databaseop(s):
    dbobject.execute("insert into seo values(?,?,?)",s)
    dbobject.commit()
    curr=dbobject.execute("select * from seo")
    return curr;

def dbtoexcel(c):
    workbook=xlsxwriter.Workbook("myoutput.xlsx")
    sheet=workbook.add_worksheet()
    rows=0
    col=0
    for x in c:
        col=col+1
        sheet.write(rows,col,x[0])
        col=col+1
        sheet.write(rows,col,x[1])
        col=col+1
        sheet.write(rows,col,x[2])
    workbook.close()

l=readfromexcel("D://pythonproj//data.xlsx")
url=l[0][0]
keywords=l[0][1]
print(keywords)
keylist=splitwords(keywords)
print("keylist",keylist)
datatext=filterurl(url)
createdb()
for x in keylist:
    print(x)
    i=searchfunc(x,datatext)
    mytuple=(url,x,i)
    c=databaseop(mytuple)
    dbtoexcel(c)
 
