import xlwings as xw
import math
import datetime as dt

# getting the patient file
patientFile = xw.Book("00897741.xlsx")

# to access the tabs
infoAll = patientFile.sheets['Info_All']

basalDay1 = patientFile.sheets['Basal_Day1']
cgmDay1 = patientFile.sheets['CGM_Day1']
infoDay1 = patientFile.sheets['Info_Day1']

basalDay2 = patientFile.sheets['Basal_Day2']
cgmDay2 = patientFile.sheets['CGM_Day2']
infoDay2 = patientFile.sheets['Info_Day2']

# setting the pre-programmed rate
preprogrammedRate = infoAll.range('H3').options(numbers=float).value
relativeRate = (preprogrammedRate/60)*5

# getting the line number
infoAllLineNumber = infoAll.range('H2').options(numbers=int).value

basalDay1LineNumber = basalDay1.range('F2').options(numbers=int).value
cgmDay1LineNumber = cgmDay1.range('F2').options(numbers=int).value
infoDay1LineNumber = infoDay1.range('H2').options(numbers=int).value

basalDay2LineNumber = basalDay2.range('F2').options(numbers=int).value
cgmDay2LineNumber = cgmDay2.range('F2').options(numbers=int).value
infoDay2LineNumber = infoDay2.range('H2').options(numbers=int).value


timeBasal1 = basalDay1.range('C2:C{}'.format(basalDay1LineNumber)).options(numbers=int).value
rateBasal1 = basalDay1.range('D2:D{}'.format(basalDay1LineNumber)).value


timeAll1 = infoDay1.range('D2:D{}'.format(infoDay1LineNumber)).options(numbers=int).value


timeCGM1 = cgmDay1.range('C2:C{}'.format(cgmDay1LineNumber)).options(numbers=int).value
rateCGM1 = cgmDay1.range('D2:D{}'.format(cgmDay1LineNumber)).options(numbers=int).value


timeBasal2 = basalDay2.range('C2:C{}'.format(basalDay2LineNumber)).options(numbers=int).value
rateBasal2 = basalDay2.range('D2:D{}'.format(basalDay2LineNumber)).value


timeAll2 = infoDay2.range('D2:D{}'.format(infoDay2LineNumber)).options(numbers=int).value


timeCGM2 = cgmDay2.range('C2:C{}'.format(cgmDay2LineNumber)).options(numbers=int).value
rateCGM2 = cgmDay2.range('D2:D{}'.format(cgmDay2LineNumber)).options(numbers=int).value



def calculateRates(tab, tabLineNumber, basalDayLineNumber, cgmDayLineNumber, relativeRate, preprogrammedRate, timeAll, timeBasal, rateBasal, timeCGM, rateCGM):

    j = 0
    i = 0

    for cell in tab.range('E2:E{}'.format(tabLineNumber)): 
        if j == (basalDayLineNumber - 1):
            cell.value = relativeRate
            continue
        # if the time matches, change rate 
        if math.floor(timeAll[i]) == math.floor(timeBasal[j]):
            
            relativeRate = (rateBasal[j]/60)*5
            cell.value = relativeRate        
            j+=1
        else:
            
            cell.value = relativeRate
        i+=1

    #for cell in tab.range('E2:E{}'.format(tabLineNumber)): 
    #    if cell.value == 0:
    #        cell.value = preprogrammedRate/60*5
            

    m = 0
    k = 0
    relativeCGM = 0

    for cell in tab.range('F2:F{}'.format(tabLineNumber)): 
        if k == (cgmDayLineNumber - 1):
            cell.value = relativeCGM
            continue
        if math.floor(timeAll[m]) == math.floor(timeCGM[k]):
            
            relativeCGM = rateCGM[k]
            cell.value = relativeCGM   
            k+=1
        else:
            
            cell.value = relativeCGM
        m+=1

    for cell in tab.range('F2:F{}'.format(tabLineNumber)): 
        if cell.value == 0:
            cell.value = relativeCGM

calculateRates(infoDay1, infoDay1LineNumber, basalDay1LineNumber, cgmDay1LineNumber, relativeRate, preprogrammedRate, timeAll1, timeBasal1, rateBasal1, timeCGM1, rateCGM1)

calculateRates(infoDay2, infoDay2LineNumber, basalDay2LineNumber, cgmDay2LineNumber, relativeRate, preprogrammedRate, timeAll2, timeBasal2, rateBasal2, timeCGM2, rateCGM2)
