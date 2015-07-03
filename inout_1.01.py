import os
import time
import sys
import commands
import xlwt
import datetime
#creating an alias for running bash commands so i dont get bored writing it
# as well as creating the xml workbook
Workbook=xlwt.Workbook(encoding="utf-8")
run = commands.getoutput
sheet=Workbook.add_sheet("Sheet1")


#set commands
firstCommand="netstat -ib | grep -e "+'"'+"en1"+'"'+" -m 1 | awk "+"'"+'{print $7}'+"'"
secondCommand="netstat -ib | grep -e "+'"'+"en1"+'"'+" -m 1 | awk "+"'"+'{print $10}'+"'"
thirdCommand="netstat -ib | grep -e "+'"'+"en1"+'"'+" -m 1 | awk "+"'"+'{print $7}'+"'"
fourthCommand="netstat -ib | grep -e "+'"'+"en1"+'"'+" -m 1 | awk "+"'"+'{print $10}'+"'"
# get current number of bytes and write headers to the excell spreadsheet
counter = 0
sheet.write(0,0,"In")
sheet.write(0,1,"Out")
sheet.write(0,2,"Timestamp")
while True:
    time.sleep(30)
    BytesI = float(run(firstCommand))#bytes in
    BytesO = float(run(secondCommand))#bytes out


    #sleep one second
    time.sleep(1)

    #get current number of bytes after one second

    BytesI_1SD=float(run(thirdCommand)) #bytes in after one second delay
    BytesO_1SD=float(run(fourthCommand)) #bytes out after one second delay

    #get difference between bytes in and out during that one second
    BdifI=(BytesI_1SD-BytesI) #difference between bytes in 
    BdifO=BytesO_1SD-BytesO #difference between bytes out

    #convert to kilobytes and write to the spreadsheet the In speed
    #out speed and hour:minute:time 
    
    KdifI= BdifI/1024.0
    KdifO= BdifO/1024.0
    sheet.write(1+counter,0,KdifI)
    sheet.write(1+counter,1,KdifO)
    
    #writes date and time to file 
	#go to http://docs.python.org/2/library/datetime.html for a complete list of 
	#formatting options if you need to change anything yourself
    sheet.write(1+counter,2,datetime.datetime.now().strftime("%B %d %H:%M:%S"))
    #print results
    print "in: "+str(round(KdifI,2))+" Kb/sec"
    print "out: "+str(round(KdifO,2))+" Kb/sec"
    Workbook.save("Networking.xls")
    #how many times the loop has gone through [determines what collumn the data
    #for this time through the loop goes on]
    counter+=1


