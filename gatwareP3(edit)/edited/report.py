
#from xlwt import *
import csv

from tkinter import *
from tkinter.filedialog import askopenfilename
#from filedialog import *
root=Tk()
root.withdraw()

global datafile,f,header

datafile=[]
data=[]

file=askopenfilename(parent=root,title='Openfile',filetypes=[('txt files', '*.txt')])
#print file

f=open(file).readlines()

print(f)

header=f[0]
header=header.rstrip()
header=header.split()
header=header[0:-2]
#print header

item='Req'

def replacer(item):
    if item in header:
        start=header.index(item)
        nextitem=header[start+1]
        finish=start+2
        JoinedItems=' '.join(header[start:finish])
        header.remove(item)
        header.remove(nextitem)
        header.insert(start,JoinedItems)

def replaceTwo(item):
    if item in header:
        start=header.index(item)
        nextitem=header[start+1]
        finish=start+2
        JoinedItems=' '.join(header[start:finish])
        header.remove(item)
        header.remove(nextitem)
        header.insert(start,JoinedItems)
        item=JoinedItems
        start=header.index(item)
        nextitem=header[start+1]
        finish=start+2
        JoinedItems=' '.join(header[start:finish])
        header.remove(item)
        header.remove(nextitem)
        header.insert(start,JoinedItems)
      
replacer('Req')
replacer('Iss')
replacer('Cum')
replacer('Lvl.')
replacer('Comp')
replaceTwo('Total')
replaceTwo('Total')
replacer('Grand')
replaceTwo('No')
replaceTwo('UNIT')
replaceTwo('Cost')
replacer('Cost')

GoodRowNumber=3
BadRowNumber=19


for line in f:
    if len(line)>0:
        line=line.split()
        datafile.append(line)
        
print(datafile)

noemptyrowdata=[]

for line in datafile:
    if not line==([]):
        noemptyrowdata.append(line)
        
#print noemptyrowdata
        
lasttwoentrydeleteddata=[]

#for line in noemptyrowdata: #strip off the last two items in each row
    #line=line[:-2]
    #asttwoentrydeleteddata.append(line)


badrow=datafile[BadRowNumber]
#print "badrow=",badrow
goodrow=datafile[GoodRowNumber]
#print "goodrow=",goodrow

list=[]
descr=[]
everyrow=[]
deltalength=[]
lengthnoempty=[]
lengthdeleted=[]

for everyrow in noemptyrowdata:
    lengtheveryrow=len(everyrow)
    lengthnoempty.append(lengtheveryrow)

#for everyrow in lasttwoentrydeleteddata:
    #lengtheveryrow=len(everyrow)
    #lengthdeleted.append(lengtheveryrow)

for eachline in datafile:
    if 'P' in eachline:
        description=' '.join(eachline[3:eachline.index('P')])
        newline=eachline[0:3]+eachline[eachline.index('P'):]
        newline.insert(3,description)
        list.append(newline)
    elif 'M' in eachline:
        description=' '.join(eachline[3:eachline.index('M')])
        newline=eachline[0:3]+eachline[eachline.index('M'):]
        newline.insert(3,description)
        list.append(newline)
    

list.insert(0,header)


file=file[12:]
file=file.replace('txt','csv')

f=open(file,'wb')
c=csv.writer(f)

for row in list:
    c.writerow(row)
f.close()



        
