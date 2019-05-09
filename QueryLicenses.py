'''
Created on Mar 19, 2019
@author: 310290474
'''
from datetime import datetime

import requests,json
from simple_salesforce import Salesforce
from decimal import Decimal
import datetime
from email.policy import SMTP
from smtplib import SMTP_SSL
import smtplib
import ssl
from smtplib import SMTP_SSL
from lib2to3.fixes.fix_input import context
from email.mime.text import MIMEText
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
import pyodbc 


conn = pyodbc.connect('Driver={SQL Server};'
                      'Server=localhost;'
                      'Database=generic;'
                      'Trusted_Connection=yes;')

cursor = conn.cursor()

timeNow = datetime.datetime.now()

#File that will be written in this location
file = open('(Enter File Name Here)' + timeNow.strftime("%Y%m%d %H%M%S") + '.csv' , 'w')

#This is the file that the program will get the salesforce login from. The format should be 'username/password/securitytoken' seperated by a newline
credFile = open('(Enter File Name Here)','r')

#Initialize email message
emailMessage = ''

timeNow = datetime.datetime.now().replace(microsecond=0).isoformat()

#This important variable determines at what limit what the weekly usage threshold can be. This number is represented as a percentage out of 100.
#If an org uses more than this threshold, support will be notified.
percentThreshold = 30

#Method to send an email with the body containing the message
def sendEmail(message):
    toAddr = '(Address the report will be set to)'
	fromAddr = '(Adress the report will be sent from)'
    msg = MIMEMultipart()
    msg['From'] = 'BHQOPSRPT01@visicu.com'
    msg['To'] = toAddr
    msg['Subject'] = 'License Report Attention Needed'
    body = message
    msg.attach(MIMEText(body,'plain'))
    text = msg.as_string()
    server = smtplib.SMTP('(enter IP adress of smtp server)' , 25)
    server.ehlo()
    server.sendmail(fromAddr, toAddr, text)
    server.quit()

#Creates the headers of the csv file
def setupFile():
    file.write('OrgName,OrgID,Data Storage Limit,Data Storage Used,Data Storage %,File Storage Limit,File Storage Used,File Storage %,' +
    'SF Licenses Total,SF Licenses Used,SF Licenses Available,Community Licenses Total,Community Licenses Used,Community Licenses Available\t\n')

#These to methods are used to convert the org limits, given in mb, to gb.
def mbToGBString(data):
    return str(round(float(data)/1024, 1))

def mbToGB(data):
    return round(float(data)/1024, 1)
setupFile()

def calculatePercentToStr(used,total):
    return str(round((float(used)/float(total))*100, 1))
    
def calculateRoCtoFloat(pastValue,presentValue):
    return round((float(pastValue) - float(presentValue))/pastValue * 100,1)

for line in credFile:
    
    #Boolean to determine if each try statement runs successfully. If not, 
    successful = True

    stringList = str(line).split('/')
    usr = str(stringList[0])
    pwd = str(stringList[1])
    sec = str(stringList[2])
    domain = usr.split('ecc.')[1]
    #Removes endline character in file
    
    if(sec.endswith('\n')):
        sec = sec[:-1]
    
    
    try:
        session = requests.Session()
        sf = Salesforce(username= usr, password=pwd, security_token=sec, sandbox = False, session = session)
        query = sf.query("SELECT name, totalLicenses,UsedLicenses FROM UserLicense")
        
        #Extract License data from query
        splitUser = str(query).split('\'Salesforce\'), (\'TotalLicenses\', ')
        userTotal = splitUser[1].split(')')[0]
        userUsed = splitUser[1].split('\'UsedLicenses\', ')[1].split(')')[0]
        userAvailable = str(int(userTotal) - int(userUsed))
        orgId = sf.query('SELECT id FROM Organization')
        orgId = str(orgId).split("Id', '")[1].split("'")[0]
       
        splitCom = str(query).split('Customer Community\'), (\'TotalLicenses\', ')
        comTotal = splitCom[1].split(')')[0]
        comUsed = splitCom[1].split('\'UsedLicenses\', ')[1].split(')')[0]
        comAvailable = str(int(comTotal) - int(comUsed))
        availableUserPercent = calculatePercentToStr(userAvailable, userTotal)
        availableComPercent = calculatePercentToStr(comAvailable, comTotal)


        #Get data from one 7 to 12 days ago, depending on which is available first.
        #This is put into a try block, because data from the last week will not always be available (like for new orgs)
        try:
            result = cursor.execute('''SELECT top 1 ID
                        ,theDate, dateadd(day,-12,getdate()) startdate, dateadd(day,-6,getdate())  enddate
                        ,orgID,orgName,comTotal,comAvailable,comUsed,availableComPercent,userTotal,userAvailable,userUsed
                        ,availableUserPercent,fileTotal,fileAvailable,fileUsed,availableFilePercent,dataTotal
                        ,dataAvailable,dataUsed,availableDataPercent
                        FROM OrgLimits where orgid =  ''' + "'" + str(orgId) + "'" +''' and
                        thedate between dateadd(day,-12,getdate()) and dateadd(day,-6,getdate())
                        order by thedate desc''')
            data= cursor.fetchall()
            #WeeklyComPercent and user percent will calucate the rate of change between 7 days ago and today.
            #This is done to see if the data has been changing, if it hasn't, then there is no need to notify support or order more licenses.
            weeklyComPercent = calculateRoCtoFloat(data[0].comAvailable, comAvailable)
            print("Weekly Com Percent: " + str(weeklyComPercent))
            weeklyUserPercent = calculateRoCtoFloat(data[0].userAvailable, userAvailable)
            print("User User Percent: " + str(weeklyUserPercent))

            #If the availble number of licenses are lower then 20% AND it's been changing recently, then notify support.
            #if(float(availableUserPercent) < 20):     
            if((float(availableUserPercent) < 20 and weeklyUserPercent >= percentThreshold) or float(availableUserPercent) <= 5):
                print(usr + ' needs more SF Licenses')
                emailMessage= emailMessage + domain + ' (' + orgId + ') '+ ' may need more SF user licenses ' + userAvailable + '/' +  userTotal + ' (' + str(availableUserPercent) + '%) left. \t\n'
                emailMessage = emailMessage + domain + "'s sf licenses has decreased from " + str(int(data[0].userAvailable)) + ' to ' + userAvailable + ' (' + str(weeklyUserPercent) + '%) since ' + str(theDate) + str("\t\n")
            if((float(availableComPercent) < 20 and weeklyComPercent >= percentThreshold) or float(availableComPercent) <= 5):
            #if(float(availableComPercent) < 20):
                print(usr + ' needs more com Licenses')  
                emailMessage= emailMessage + domain + ' (' + orgId + ') '+ ' may need more community user licenses ' + comAvailable + '/' + comTotal + ' (' + availableComPercent + '%)  left.\t\n'
                emailMessage = emailMessage + domain + "'s community licenses has decreased from " + str(int(data[0].comAvailable)) + ' to ' + comAvailable + ' (' + str(weeklyComPercent) + '%) since ' + str(data[0].theDate) + str('\t\n')
        
        
        except IndexError as e:
            #If no data is found for last week, we will only evaluate the available percentage, since the 'data' variable will not be available
            print('Could not fetch data from last week for ' + domain)
            weeklyComPercent = percentThreshold
            weeklyUserPercent = percentThreshold
            if(float(availableUserPercent) < 20):
                print(usr + ' needs more SF Licenses')
                emailMessage= emailMessage + domain + ' (' + orgId + ') '+ ' may need more SF user licenses ' + userAvailable + '/' +  userTotal + ' (' + str(availableUserPercent) + '%) left. \t\n'
            if(float(availableComPercent) < 20):
                print(usr + ' needs more com Licenses')  
                emailMessage= emailMessage + domain + ' (' + orgId + ') '+ ' may need more community user licenses ' + comAvailable + '/' + comTotal + ' (' + availableComPercent + '%)  left.\t\n'
    
        #API call to get data limits       
        headers = {'authorization': 'Bearer ' + sf.session_id}
        response = session.get(url = 'https://' + domain + '.my.salesforce.com/services/data/v43.0/limits', headers = headers)
        
        #Extract data and storage limits from response      
        totalFileStorageMB = str(response.content).split('FileStorageMB":{"Max":')[1].split(',')[0]
        remainingFileStorageMB = str(response.content).split('FileStorageMB":{"Max":')[1].split(':')[1].split('}')[0]
        usedFileMB = str(float(totalFileStorageMB) - float(remainingFileStorageMB))
        usedFilePercent = calculatePercentToStr(usedFileMB,totalFileStorageMB)
        availableFilePercent = str(round(100 - float(usedFilePercent),1))
                
        totalDataStorageMB = str(response.content).split('"DataStorageMB":{"Max":')[1].split(',')[0]
        remainingDataStorageMB = str(response.content).split('DataStorageMB":{"Max":')[1].split(':')[1].split('}')[0]
        usedDataMB =  (str(float(totalDataStorageMB) - float(remainingDataStorageMB)))
        usedDataPercent = calculatePercentToStr(usedDataMB, totalDataStorageMB)
        availableDataPercent = str(round(100 -float(usedDataPercent),1))


        #Again, Get data from one 7 to 12 days ago for file/storage data, depending on which is available first.
        #no need to query odbc again, we'll just reuse the 'data' variable.
        try:
            #Calculate weekly rate of change for storage and file limits
            weeklyFilePercent = calculateRoCtoFloat(data[0].fileAvailable,mbToGB(remainingFileStorageMB))
            print("Weekly File Percent: " + str(weeklyFilePercent))
            weeklyDataPercent = calculateRoCtoFloat(data[0].dataAvailable, mbToGB(remainingDataStorageMB))
            print("Weekly Data Percent: " + str(weeklyDataPercent))
            
            #Notify support if file or data usage is below 20% AND weekly rate of change is over the percentage threshold
            if((float(availableFilePercent) <20 and weeklyFilePercent >= percentThreshold) or float(availableFilePercent) <= 5):
            #if(float(availableFilePercent) <20):
                print(usr + ' (' + orgId + ')'+ ' needs more file space')
                emailMessage= emailMessage + domain  + ' (' + orgId + ') '+' may need more file space ' + usedFileMB + '/' + totalFileStorageMB + ' (' + availableFilePercent+ '%) left.\t\n'
                emailMessage = emailMessage + domain + ' file usage has decreased from ' + str(data[0].fileAvailable) + ' to ' + remainingFileStorageMB + ' (' + str(weeklyFilePercent) + '%) since ' + str(data[0].theDate)+ str('\t\n')
                    
            if((float(availableDataPercent)  <20 and weeklyDataPercent >= percentThreshold) or float(availableDataPercent) <= 5):
            #if(float(availableDataPercent)  <20):
                print(usr + ' (' + orgId + ')'+' needs more data space')  
                emailMessage= emailMessage + domain + ' (' + orgId + ') '+ ' may need more data space ' + usedDataMB + '/' + totalDataStorageMB + ' (' +availableDataPercent+ '%) left.\t\n'
                emailMessage = emailMessage + domain + ' data usage has decreased from ' + str(data[0].dataAvailable) + ' to ' + remainingDataStorageMB + ' (' + str(weeklyDataPercent) + '%) since ' +str(data[0].theDate)+ str('\t\n')        

        
        except IndexError as e:
            #If no data is found for last week, set the weekly percents to the threshold,
            #that way support will always be notified about the org limits until 7 days data is collected.
            print('Could not fetch file/storage data from last week for ' + domain)
            weeklyFilePercent = percentThreshold
            weeklyDataPercent = percentThreshold
        
            #Notify support if file or data usage is below 20% AND weekly rate of change is over the percentage threshold
            if(float(availableFilePercent) <20):
                print(usr + ' (' + orgId + ')'+ ' needs more file space')
                emailMessage= emailMessage + domain  + ' (' + orgId + ') '+' may need more file space ' + usedFileMB + '/' + totalFileStorageMB + ' (' + availableFilePercent+ '%) left.\t\n'
                    
            if(float(availableDataPercent)  <20):
                print(usr + ' (' + orgId + ')'+' needs more data space')  
                emailMessage= emailMessage + domain + ' (' + orgId + ') '+ ' may need more data space ' + usedDataMB + '/' + totalDataStorageMB + ' (' +availableDataPercent+ '%) left.\t\n'
                
    
        #Write to our file the stats of the current org.   
        file.write(domain + ','+ orgId +',' +mbToGBString(totalDataStorageMB) +',' + mbToGBString(usedDataMB) +',' +usedDataPercent + ',' + mbToGBString(totalFileStorageMB) +',' + mbToGBString(usedFileMB) + ','+  usedFilePercent +',' + userTotal +','  +userUsed + ',' +userAvailable + ',' + comTotal + ','+ comUsed + ',' + comAvailable + '\t\n')
        print('Data extracted succesfully for ' + orgId)
        
        #Insert stats into our obdc database
        insertStatement = '''INSERT INTO OrgLimits
                       ([theDate]
                       ,[orgID]
                       ,[orgName]
                       ,[comTotal]
                       ,[comAvailable]
                       ,[comUsed]
                       ,[availableComPercent]
                       ,[userTotal]
                       ,[userAvailable]
                       ,[userUsed]
                       ,[availableUserPercent]
                       ,[fileTotal]
                       ,[fileAvailable]
                       ,[fileUsed]
                       ,[availableFilePercent]
                       ,[dataTotal]
                       ,[dataAvailable]
                       ,[dataUsed]
                       ,[availableDataPercent])
                 VALUES
                       ('''+ "'" + str(timeNow) +"'" + '''
                       ,''' +"'"+ str(orgId) +"'" '''
                       ,'''  +"'"+ str(domain) +"'"+ '''
                       ,'''  + str(comTotal)   + '''
                       , '''   + str(comAvailable)  + '''
                       , '''  + str(comUsed)  + '''
                       , ''' + str(availableComPercent) + '''
                      , ''' + str(userTotal) + '''
                      , ''' + str(userAvailable) + '''
                      , ''' + str(userUsed) + '''
                      , ''' + str(availableUserPercent) + '''
                      , ''' + str(mbToGB(totalFileStorageMB)) + '''
                      , ''' + str(mbToGB(remainingFileStorageMB)) + '''
                      , ''' + str(mbToGB(usedFileMB)) + '''
                      , ''' + str(availableFilePercent)+ '''
                      , ''' + str(mbToGB(totalDataStorageMB)) + '''
                      , ''' + str(mbToGB(remainingDataStorageMB)) + '''
                      , ''' + str(mbToGB(usedDataMB)) + '''
                      , ''' + str(availableDataPercent) + ''')
    
                    '''
        cursor.execute(insertStatement)
        conn.commit()
    
    #Except any error and include the error in the email    
    except Exception as e:
        try:
            orgId
        except NameError as n:
            print('login no succesful for ' + usr)
            orgId = usr
        print('Data extraction not Succesful for ' + orgId + '!!!!!!!!!! Returned: ' + str(e) + ' \t\n')
        emailMessage = emailMessage + 'Data extraction not Succesful for ' + domain + ' (' + orgId + ') !!!!!!!!!! Returned: ' + str(e) +  ' \t\n'
    print('')

#Check to see if message is empty, no need to send empty message
if(emailMessage != ''):
    print('Sending the following message: \t\n' + emailMessage)
    sendEmail(emailMessage)
else:
    print('Message empty, not sending message')

#Cleanup our opened resources
conn.close()
del conn
file.close()
