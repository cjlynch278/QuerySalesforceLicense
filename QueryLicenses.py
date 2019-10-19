'''
Created on Oct 11, 2019

@author: 310290474
'''
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
import smtplib, ssl
import email
import mime
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

import sys

#conn = pyodbc.connect('Driver={SQL Server};'
 #                     'Server=localhost;'
 #                     'Database=generic;'
 #                     'Trusted_Connection=yes;')

#cursor = conn.cursor()

timeNow = datetime.datetime.now()

file = open('LicenseReports' + timeNow.strftime("%Y%m%d %H%M%S") + '.csv' , 'w')
credFile = open('Credentials.txt','r')

#Testing File
#credFile = open('D:\\Python\\TestingFiles\\Reporting Credentials.txt','r')

#Jim *insists* that the message follows this format:
#org1 SF Licenses
#org2 SF Licenses...
#    .... org N SF Licenses
#org1 Community Licenses 
#org2 Community Licesnses...
#    .... org N Community Licenses.

#Because of the above format, I will need a list variable to keep track of all of the community license
#   strings to print out.


#This important variable determines at what limit what the weekly usage threshold can be. This number is represented as a percentage out of 100.
#If an org uses more than this threshold, support will be notified.

#Method to send an email with the body containing the message
class Backend:
    def testEmail(self,message):
        port = 465  # For SSL
        userName = 'pythonautomationmail4@gmail.com'
        password = 'PythonAuto7'
    
    # Create a secure SSL context
        context = ssl.create_default_context()
        toAddr = 'pythonautomationmail4@gmail.com'
        msg = MIMEMultipart()
        msg['From'] = 'pythonautomationmail4@gmail.com'
        msg['To'] = toAddr
        msg['Subject'] = 'Daily Salesforce Report'
        body = message
        msg.attach(MIMEText(body,'plain'))
        text = msg.as_string()
        with smtplib.SMTP_SSL("smtp.gmail.com", port, context=context) as server:
            server.login(userName, password)
            server.sendmail(userName, userName, text)
    
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
    def calculateRoCtoFloat(pastValue,presentValue):
        if(pastValue == 0):
            return 0
        
        return round((float(pastValue) - float(presentValue))/pastValue * 100,1)
    
    
    def calculatePercentToStr(used,total):
        if(float(total) == 0 or float(used) == 0):
            return str(0)
            
        return str(round((float(used)/float(total))*100, 1))
    
    def getResponseUsage(self, attribute, response):
        totalUsage = str(response[str(attribute)]['Max'])
        remainingUsage= str(response[str(attribute)]['Remaining'])
        message = str(remainingUsage) + '/' + str(totalUsage) + ' ' +attribute + ' left'
        
        return message
        
    def getSoqlData(self, sf, licenseType):
            queryText = str("SELECT name, totalLicenses,UsedLicenses FROM UserLicense where name = '"+  str(licenseType) + "'")
            query = sf.query(queryText)
            total = query['records'][0]['TotalLicenses']
            used = query['records'][0]['UsedLicenses']
            available = str(int(total) - int(used))
            return( str(available) + '/' + str(total) + ' ' + licenseType +"Licenses available")
    
    def run(self, restQueryList, soqlQueryList):
        credFile = open('Credentials.txt','r')

        sfLicenseMessage = ""
        dataMessage = ""
        fileMessage = ""
        commLicenseMessage = ""
        emailMessage = ''
        platLicenseMessage =''
        timeNow = datetime.datetime.now().replace(microsecond=0).isoformat()
    
        for line in credFile:
            #Boolean to determine if each try statement runs successfully. If not, 
            successful = True
            stringList = str(line).split('/')
            usr = str(stringList[0])
            pwd = str(stringList[1])
            sec = str(stringList[2])
            domain = "na53.salesforce.com"
        
            if(sec.endswith('\n')):
                sec = sec[:-1]
           
            try:
                session = requests.Session()
                sf = Salesforce(username= usr, password=pwd, security_token=sec, sandbox = False, session = session)
                
                nameQuery = sf.query("SELECT ID, Name FROM Organization")
                name= str(nameQuery['records'][0]['Name'])
                emailMessage = emailMessage + "Stats for " + name + ':\n'
        
                headers = {'authorization': 'Bearer ' + sf.session_id}
        
                response = session.get(url = 'https://' + domain + '/services/data/v43.0/limits', headers = headers)
                response = json.loads(response.text)
        
                for item in soqlQueryList:
                    emailMessage = emailMessage + self.getSoqlData(sf, str(item)) +'\n'
                for item in restQueryList:
                    emailMessage = emailMessage + self.getResponseUsage(str(item), response) + '\n'


        
        
            
            #Except any error and include the error in the email    
            except Exception as e:
        
                print('Data extraction not Succesful for ' + usr + '!!!!!!!!!! Returned: ' + str(e) )
                emailMessage = emailMessage + 'Data extraction not Succesful for ' + name + '!!!!!!!!!! Returned: ' + str(e) +  ' \t\n'
            emailMessage = emailMessage + '\n'
            print('')
        
        
        
         
        #Check to see if message is empty, no need to send empty message
        if(emailMessage != ''):
         
            print('Sending the following message: \t\n' + emailMessage)
            self.testEmail(emailMessage)        
        else:
            print('Message empty, not sending message')
        
        return emailMessage

        file.close()
  
    def runOrig(self):
        sfLicenseMessage = ""
        dataMessage = ""
        fileMessage = ""
        commLicenseMessage = ""
        emailMessage = ''
        platLicenseMessage =''
    
        timeNow = datetime.datetime.now().replace(microsecond=0).isoformat()
    
        for line in credFile:
            #Boolean to determine if each try statement runs successfully. If not, 
            successful = True
            stringList = str(line).split('/')
            usr = str(stringList[0])
            pwd = str(stringList[1])
            sec = str(stringList[2])
            domain = "na53.salesforce.com"
        
            if(sec.endswith('\n')):
                sec = sec[:-1]
           
            try:
                session = requests.Session()
                sf = Salesforce(username= usr, password=pwd, security_token=sec, sandbox = False, session = session)
                
                nameQuery = sf.query("SELECT ID, Name FROM Organization")
                name= str(nameQuery['records'][0]['Name'])
                emailMessage = emailMessage + "Stats for " + name + ':\n'
        
                headers = {'authorization': 'Bearer ' + sf.session_id}
        
                response = session.get(url = 'https://' + domain + '/services/data/v43.0/limits', headers = headers)
                response = json.loads(response.text)
        
                    
                #Extract data and storage limits from SOQL Queries
                emailMessage = emailMessage + self.getSoqlData(sf, 'Salesforce') +'\n'
                emailMessage = emailMessage + self.getSoqlData( sf, 'Salesforce Platform ') +'\n'
                emailMessage = emailMessage + self.getSoqlData(sf, "Community Community Login") + '\n'
                #Extract data and storage limits from REST Responses
                emailMessage = emailMessage + self.getResponseUsage('AnalyticsExternalDataSizeMB', response, name) + '\n'
                emailMessage = emailMessage + self.getResponseUsage('FileStorageMB', response, name)+ '\n'
                emailMessage = emailMessage + self.getResponseUsage('DailyApiRequests', response, name)+ '\n'
        
        
        
            
            #Except any error and include the error in the email    
            except Exception as e:
        
                print('Data extraction not Succesful for ' + usr + '!!!!!!!!!! Returned: ' + str(e) )
                emailMessage = emailMessage + 'Data extraction not Succesful for ' + name + '!!!!!!!!!! Returned: ' + str(e) +  ' \t\n'
            emailMessage = emailMessage + '\n'
            print('')
        
        
        
         
        #Check to see if message is empty, no need to send empty message
        if(emailMessage != ''):
         
            print('Sending the following message: \t\n' + emailMessage)
            self.testEmail(emailMessage)        
        else:
            print('Message empty, not sending message')
        
        file.close()
  
