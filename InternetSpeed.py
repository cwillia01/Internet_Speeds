import speedtest
import datetime
import openpyxl

#provide information about your excel file below
excelLoc = "InternetSpeed.xlsx"
sheetName = "Speeds"
dateCol = 'A'
downloadCol = 'C'
uploadCol = 'B'

#gets the date and time
def today_date():
    today = datetime.datetime.now()
    todayFormat = today.strftime("%d/%m/%Y")
    todayTime = today.strftime("%H:%M")
    thisDate = todayFormat + " " + todayTime
    return thisDate

#gets the next available row in the sheet
#avoids rows being overwritten
def get_next_row(sheet):
    maxRow = sheet.max_row
    currentRow = str(maxRow+1)
    return currentRow

#get both download and upload speed
def return_internet_speed():
    st = speedtest.Speedtest()
    download = st.download()/1000000
    upload = st.upload()/1000000
    return [download, upload]

#add the infromation to the excel sheet
def add_today_speed(filename, sheetName, dateCol, downloadCol, uploadCol):
    xfile = openpyxl.load_workbook(filename)
    sheet = xfile[sheetName]

    #gather both download and upload speed
    speeds = return_internet_speed()
    downloadSpeed = speeds[0]
    uploadSpeed = speeds[1]


    #determine next available row and map columns
    nextRow = get_next_row(sheet)

    #send date, download and upload information to the next available row
    sheet[dateCol + nextRow] = today_date()
    sheet[downloadCol + nextRow] = downloadSpeed
    sheet[uploadCol + nextRow] = uploadSpeed

    xfile.save(filename)

#####execute
add_today_speed(excelLoc, sheetName, dateCol, downloadCol, uploadCol)