from tkinter import *
import tkinter as tk
from QueryLicenses import Backend

backend =  Backend()


def enterText(textArea,text):
    textArea.insert('1.0', str(text) + "\n")

def clearText(textArea, restQueryList, soqlQueryList):
    textArea.delete('1.0', END)
    restQueryList.clear()
    soqlQueryList.clear()

def executeQuery(restQueryList, soqlQueryList, resultPane):
    resultPane.insert("1.0", backend.run(restQueryList, soqlQueryList))
    
def addSoqlItems(button, itemList, textArea, soqlQueryList):
    for item in button.curselection():
        soqlQueryList.append(str(itemList[item]))
        textArea.insert(END, str(itemList[item]) + "\n")
        
def addRestItems(button, itemList, textArea, restQueryList):
    print(button.curselection())
    for item in button.curselection():
        restQueryList.append(str(itemList[item]))
        textArea.insert(END, str(itemList[item]) + "\n")
           

restQueryList = []
soqlQueryList = []
           
root = Tk()
root.title('Salesforce Data Extraction')
root.resizable(width=FALSE, height=FALSE)
root.geometry('{}x{}'.format(540, 810))





soqlItems = [
    'Salesforce',
    'Salesforce Platform',
    'Customer Community Login',
    'Customer Community Plus Login',
    'Gold Partner',
    'Chatter Free',
    'Chatter External',
    'Partner App Subscription',
    'External Identity',
    'Force.com - Free',
    'Partner Community',
    'Partner Community Login',
    'Silver Partner',
    'Analytics Cloud Integration User'
    ]
restItems = [    
    'AnalyticsExternalDataSizeMB',
    'FileStorageMB',
    'DailyApiRequests',
    'AnalyticsExternalDataSizeMB',
    'ConcurrentAsyncGetReportInstances',
    'ConcurrentEinsteinDataInsightsStoryCreation',
    'ConcurrentEinsteinDiscoveryStoryCreation',
    'ConcurrentSyncReportRuns',
    'DailyAnalyticsDataflowJobExecutions',
    'DailyAnalyticsUploadedFilesSizeMB',
    'DailyApiRequests',
    'DailyAsyncApexExecutions',
    'DailyBulkApiRequests',
    'DailyDurableGenericStreamingApiEvents',
    'DailyDurableStreamingApiEvents',
    'DailyEinsteinDataInsightsStoryCreation',
    'DailyEinsteinDiscoveryPredictAPICalls',
    'DailyEinsteinDiscoveryPredictionsByCDC',
    'DailyEinsteinDiscoveryStoryCreation',
    'DailyGenericStreamingApiEvents',
    'DailyStandardVolumePlatformEvents',
    'DailyStreamingApiEvents',
    'DailyWorkflowEmails',
    'DataStorageMB',
    'DurableStreamingApiConcurrentClients',
    'FileStorageMB',
    'HourlyAsyncReportRuns',
    'HourlyDashboardRefreshes',
    'HourlyDashboardResults',
    'HourlyDashboardStatuses',
    'HourlyLongTermIdMapping',
    'HourlyODataCallout',
    'HourlyPublishedPlatformEvents',
    'HourlyPublishedStandardVolumePlatformEvents',
    'HourlyShortTermIdMapping',
    'HourlySyncReportRuns',
    'HourlyTimeBasedWorkflow',
    'MassEmail',
    'MonthlyEinsteinDiscoveryStoryCreation',
    'MonthlyPlatformEvents',
    'Package2VersionCreates',
    'PermissionSets',
    'SingleEmail',
    'StreamingApiConcurrentClients'
    ]

soqlPickList = Listbox(root, selectmode=MULTIPLE, width=40, height=10)
restPickList = Listbox(root,selectmode=MULTIPLE, width=40, height=10)
for item in soqlItems:
    soqlPickList.insert(END, item)
for item in restItems:
    restPickList.insert(END, item)
    
soqlTitle = Label(root, text="SOOQL Query Options", font=("Courier", 12)) 
soqlTitle.place  (x = 20, y = 15) 
soqlPickList.place(x= 20, y = 40)

soqlTitle = Label(root, text="REST Query Options", font=("Courier", 12)) 
soqlTitle.place  (x = 275, y = 15) 
restPickList.place(x = 275, y = 40)

queriesTitle = Label(root, text="Selected Queries", font=("Courier", 12)) 
queriesTitle.place  (x = 100, y = 240) 

selectedText = Text(root, width = 40, height = 10)
selectedText.place(x = 100, y = 265)

clearButton = Button(text = "Clear", command = lambda: clearText(selectedText,restQueryList,soqlQueryList))
clearButton.place(x = 100, y = 435)

executeButton = Button(text = "Execute", command = lambda: (executeQuery(restQueryList, soqlQueryList, resultText)))
executeButton.place(x = 150, y = 435)

soqlAddButton = Button(text = "Add Soql Items", command = lambda: addSoqlItems(soqlPickList,soqlItems,selectedText,soqlQueryList)) 
soqlAddButton.place(x = 20, y = 210)

restAddButton = Button(text = "Add Rest Items", command = lambda: addRestItems(restPickList,restItems,selectedText,restQueryList)) 
restAddButton.place(x = 275, y = 210)

resultsTitle = Label(root, text="Results", font=("Courier", 12)) 
resultsTitle.place  (x = 25, y = 465) 

resultText = Text(root, width = 60, height = 17)
resultText.place(x = 25, y = 490)

clearResults = Button(text = "Clear", command = lambda: clearText(resultText,restQueryList,soqlQueryList))
clearResults.place(x = 25, y = 770)

root.mainloop()

