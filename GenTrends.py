# First, we'll import pandas, a data processing and CSV/Excel file I/O library
from calendar import c, week
import pandas as pd
from matplotlib import pyplot as plt
import glob #python pattern matching files 
import os
from datetime import datetime
import time

#pip3 install pandas
#pip3 install openpyxl 

# Set the figure size - handy for larger output
plt.rcParams["figure.figsize"] = [10, 6]

class Appointment:
    #Appointment class for plotting afternoon appointments. Could also be used for other purposes. 
    def __init__(self,day, time):
        self.day = day
        self.time = time

def getFiles():

    #Inputs: no inputs
    #Purpose: to return two lists containing the file paths for the data files for current semester data and 
    #         past semester data. Past semester data paths are sorted based on semester date. 
    #Return values: Two lists embedded within a list -> [currSem[],pastSem[]]

    #retrieving file paths for xlsx data files
    currSemDataPath = glob.glob("Data/Curr_semester_totals/*.xlsx")
    pastSemFoldersPath = glob.glob("Data/Prev_semester_totals/*")
    
    pastDataPathDict = {}
    for path in pastSemFoldersPath:

        dataPaths = glob.glob(path+"/*.xlsx")
        path_ID = path[len(path)-4:len(path)]
        if "fall" in path.lower():
            path_ID += "1"
        else:
            path_ID+= "0"

        path_ID = int(path_ID)
        pastDataPathDict[path_ID] = dataPaths
    
    pastSemDataPaths = []
    for key in sorted(pastDataPathDict):
        pastSemDataPaths.append(pastDataPathDict[key])

    return [currSemDataPath,pastSemDataPaths]

def apptsBySem(currSemData,pastSemPaths):

    appts = list(currSemData["Email Address"])
    filterPlaceHolder(appts)
    currTotal = len(appts)
    pastSemsTotals = []
    for paths in pastSemPaths:
        for path in paths:
            if "Condensed" not in path:
                semData = pd.read_excel(path) #opening data file to read
                appts = list(semData["Email Address"])
                filterPlaceHolder(appts)
                pastSemsTotals.append(len(appts))
    
    allSemTotals = pastSemsTotals
    allSemTotals.append(currTotal)
    print(allSemTotals)

    plotdata = pd.DataFrame(
        {"Total Number of Conferences": allSemTotals})
    plotdata.plot(kind="bar")
    plt.xticks(rotation=0, horizontalalignment="center")
    plt.title("Total conferences by semester")
    plt.xlabel("Semester")
    plt.ylabel("Number of Conferences")
    plt.show()

def filterPlaceHolder(list):

    for name in list:
        if "jmullin2" in name:
            list.remove(name)


def getAppointments(data,time_threshold,lower):

    #returns a list of appointment objects that have a time below or above the provided time_threshold 
    #        value. (wether to go above or below is determined by the lower flag). This function is 
    #        used to get the number of afternoon and evening appointments 

    listDates = list(data["Appointment Date"])
    listTimes = list(data["Start Time"])

    currSemAfternoonAppts = []
    
    for i in range(len(listDates)):
        # %I:%M%p
        time =  datetime.strptime(listTimes[i],'%I:%M%p').hour
        if lower:
            if time < time_threshold:
                if listDates[i].weekday() < 4:
                    currSemAfternoonAppts.append(Appointment(listDates[i].weekday(),time))
        else:
            if time > time_threshold:
                currSemAfternoonAppts.append(Appointment(listDates[i].weekday(),time))

    return currSemAfternoonAppts

def plotEveningAppt(currSemData,pastSemPaths):

    #plots the number of evening appointments by weekday 

    time_threshold = datetime.strptime('5:00pm','%I:%M%p').hour
    currSemAppts = getAppointments(currSemData,time_threshold,False)
    histSemAppts = []
    for paths in pastSemPaths:
        for path in paths:
            if "Condensed" not in path:
                semData = pd.read_excel(path) #opening data file to read
                histSemAppts.extend(getAppointments(semData,time_threshold,False))
    
    weekdays = []
    for appt in currSemAppts:
        weekdays.append(appt.day)
    for appt in histSemAppts:
        weekdays.append(appt.day)
    
    weekdays = set(weekdays)

    dictFreq = {}
    dictFreq_hist = {}

    for d in weekdays:
        dictFreq[d] = 0
        dictFreq_hist[d] = 0
    
    for a in currSemAppts:
        dictFreq[a.day]+=1
    for a in histSemAppts:
        dictFreq_hist[a.day]+=1/(len(pastSemPaths))
    
    x_values = list(dictFreq.values())
    x_values_hist = list(dictFreq_hist.values())

    week = {0:'Monday',1:'Tuesday',2:'Wednesday',3:'Thursday',4:'Friday',5:'Saturday',6:'Sunday'}
    y_values = []
    for d in weekdays:
        y_values.append(week[d])

    plotdata = pd.DataFrame(
        {"Number of Conferences": x_values, "Average" : x_values_hist}, index=y_values)
    plotdata.plot(kind="bar")
    plt.xticks(rotation=0, horizontalalignment="center")
    plt.title("Number of evening appointments by weekday")
    plt.xlabel("Week Day")
    plt.ylabel("Number of Conferences")
    plt.show()


def plotAfternoonAppt(currSemData,pastSemPaths):

    #plots the number of afternoon appointments by weekday

    time_threshold = datetime.strptime('5:00pm','%I:%M%p').hour
    currSemAppts = getAppointments(currSemData,time_threshold,True)
    histSemAppts = []
    for paths in pastSemPaths:
        for path in paths:
            if "Condensed" not in path:
                semData = pd.read_excel(path) #opening data file to read
                histSemAppts.extend(getAppointments(semData,time_threshold,True))
    
    weekdays = []
    for appt in currSemAppts:
        weekdays.append(appt.day)
    for appt in histSemAppts:
        weekdays.append(appt.day)
    
    weekdays = set(weekdays)

    dictFreq = {}
    dictFreq_hist = {}

    for d in weekdays:
        dictFreq[d] = 0
        dictFreq_hist[d] = 0
    
    for a in currSemAppts:
        dictFreq[a.day]+=1
    for a in histSemAppts:
        dictFreq_hist[a.day]+=1/(len(pastSemPaths))
    
    x_values = list(dictFreq.values())
    x_values_hist = list(dictFreq_hist.values())

    week = {0:'Monday',1:'Tuesday',2:'Wednesday',3:'Thursday'}
    y_values = []
    for d in weekdays:
        y_values.append(week[d])

    plotdata = pd.DataFrame(
        {"Number of Conferences": x_values, "Average" : x_values_hist}, index=y_values)
    plotdata.plot(kind="bar")
    plt.xticks(rotation=0, horizontalalignment="center")
    plt.title("Number of afternoon appointments by weekday")
    plt.xlabel("Week Day")
    plt.ylabel("Number of Conferences")
    plt.show()

def extractWeekNumbers(data):

    #returns a list of week numbers that represent each appointment date. 
    # we need this to be able to know what week number each appointment date corresponds to

    listDates = list(data["Appointment Date"])
    listDates.sort() #sorting dates in ascending order
    
    weekNumbers = []

    for i in range(len(listDates)):

        weekNumbers.append(listDates[i].date().isocalendar()[1])
    
    dictKeys = list(set(weekNumbers))
    dictValues = list(range(len(dictKeys)))
    dict = {}
    for i in range(len(dictKeys)):
        dict[dictKeys[i]] = dictValues[i]
    newWeekNums = []
    print(weekNumbers)
    for num in weekNumbers:
        newWeekNums.append(dict[num])

    return newWeekNums

def apptByWeek(currSemData,pastSemPaths):

    #plots the number of appointments by weekday

    weekNumbers = extractWeekNumbers(currSemData)
    print(weekNumbers)

    dictFreq = {}
    
    for week in set(weekNumbers):
        dictFreq[week] = 0
    
    for num in weekNumbers:
        dictFreq[num]+=1

    x_values = list(dictFreq.values())
 

    plotdata = pd.DataFrame(
        {"Number of Conferences": x_values})
    plotdata.plot(kind="bar")
    plt.xticks(rotation=0, horizontalalignment="center")
    plt.title("Writing center usage by week")
    plt.xlabel("Week Number")
    plt.ylabel("Number of Conferences")
    plt.show()

def historicalAppointments(pastSemPaths,currSemAppts):

    semDatas = []
    appointmentTotals = []
    for paths in pastSemPaths:
        for path in paths:
            if "Condensed" in path:
                print(path)
                semDataCompr = pd.read_excel(path) #opening data file to read
                sem_totals = list(semDataCompr["Total Appointments"])
                removeNAN(sem_totals) 
                appointmentTotals.extend(sem_totals)
                semDatas.append(len(sem_totals))
    
    appointmentTotals.extend(currSemAppts)
    y_values = list(set(appointmentTotals))

    dictFreq = {}
    for y in y_values:
        dictFreq[y] = 0

    k = 0
    for i in range(len(semDatas)):
        for j in range(semDatas[i]):
            dictFreq[appointmentTotals[k]]+=((1/semDatas[i])*100)/len(semDatas)
            k+=1

    x_values = list(dictFreq.values())
    return [x_values,y_values]


def appointmentFreq(dataCompr,pastSemPaths):

    totalAppointmentsPerClient = list(dataCompr["Total Appointments"])
    removeNAN(totalAppointmentsPerClient)
    numClients = len(totalAppointmentsPerClient)

    historical_values = historicalAppointments(pastSemPaths,totalAppointmentsPerClient)

    y_values = historical_values[1]
    histX_values = historical_values[0]

    for i in range(len(y_values)): y_values[i] = int(y_values[i])

    dictFreq = {}
    for y in y_values:
        dictFreq[y] = 0
    print(dictFreq)
    for num in totalAppointmentsPerClient:
        dictFreq[int(num)]+=(1/numClients)*100

    x_values = list(dictFreq.values())
    for i in range(len(y_values)): y_values[i] = str(y_values[i])

    # print("sum x values: ", sum(x_values))
    # print("y values: ",y_values)

    # Create a data frame with one column, "Num Conferences"
    plotdata = pd.DataFrame(
        {"Number of Conferences": x_values, "Average": histX_values},
         index=y_values)
    plotdata.plot(kind="bar")
    plt.xticks(rotation=0, horizontalalignment="center")
    plt.title("Writing center usage by conference frequency")
    plt.xlabel("Total number of conferences")
    plt.ylabel("Percentage of WC users")
    plt.show()

def removeNAN(list):

    #removes NAN entries from given list
    for elem in list: #removing NAN (empty) entries
        if pd.isna(elem):
            list.remove(elem) 

def main():

    #retreiving file paths for current semester data and past semester data
    paths = getFiles()

    currSemPaths = paths[0] 
    pastSemPaths = paths[1]


    #extracting current semester data from xlsx file paths
    currSemData = pd.read_excel(currSemPaths[0])
    currSemDataCompr = pd.read_excel(currSemPaths[1])

    #plotting number of conferences by frequency for current semester vs historical
    appointmentFreq(currSemDataCompr,pastSemPaths)

    #extracting Week numbers from current semester data (uncompressed)
    apptByWeek(currSemData,pastSemPaths)

    #plotting number of afternoon appointments by weekday
    plotAfternoonAppt(currSemData,pastSemPaths)

    #plotting number of evening appointments by weekday
    plotEveningAppt(currSemData,pastSemPaths)

    #plotting total number of conferences by semester
    apptsBySem(currSemData,pastSemPaths)

main()