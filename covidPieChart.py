#!python3
#Covid cases pie chart by county, Romania
import json, requests, openpyxl, os, sys
import pyinputplus as pyip


#Create a folder where we'll put the file with data
abspath = os.path.abspath(sys.argv[0])
dname = os.path.dirname(abspath)
os.chdir(dname)
os.makedirs('covidData', exist_ok=True)

#Create the excel file
wb = openpyxl.Workbook()
sheet = wb.active
counties = []
cases = []

#Get the JSON data
url = 'https://www.graphs.ro/json.php'
res = requests.get(url)
try:
    res.raise_for_status()
except:
    print('Could not connect to server.')
w = json.loads(res.text)

#Select the date you're interested in

userDate = pyip.inputDate(prompt= 'Enter a date (yyyy-mm-dd) format: ',formats= ['%Y-%m-%d'])
userDate = str(userDate)
year, month, day = userDate.replace(' ', '').split('-')

userCounty = pyip.inputStr(prompt= 'You are interested in a specific county? ')
userCounty = userCounty.lower()
userCounty = userCounty[0].upper() + userCounty[1:]

#Creating a list with all the county names from the data
countyDict = w['covid_romania'][0]['county_data']
totalCounties = [ county['county_name'].lower() for county in countyDict] #Replacing the for loop with a list comprehension

def daySelector():
    for i in w['covid_romania']:
        for v in i.values():
            if v == userDate:
                dayData = i
                cases = i['new_cases_today']
                tests = i['new_tests_today']
                print(str(cases) + ' cases confirmed from '+ str(tests) + ' tests, that means ' + str(round(cases / (tests / 100),1)) + '%') 
                return dayData['county_data']
    

def timeWalking():
    intDay = int(day)
    intMonth = int(month)
    intYear = int(year)
    days = []

    for i in range(14):
        if (intDay == 0 and intMonth == 10) or (intDay == 0 and intMonth == 12) or (intDay == 0 and intMonth == 5) or (intDay == 0 and intMonth == 7):
            intDay = 30
            intMonth -= 1
        elif intDay == 0 and intMonth == 3:
            if intYear % 4 == 0:
                intDay =2
                intMonth -= 1
            else:
                intDay = 28
                intMonth -= 1
        elif intDay == 0:
            intDay = 31
            intMonth -= 1
        days.append(year + '-' + str(intMonth).rjust(2, '0')+ '-' + str(intDay).rjust(2,'0')) #the rjust(2, '0') puts a 0 on the left if the value is single digit.
        intDay -= 1
    return days
    
def timelineSelector():
    cases = {}
    timeline = timeWalking()
    for i in w['covid_romania']:
        for v in i.values():
            if v in timeline:
                dayData = i
                day = i['reporting_date']
                countyList = i['county_data']
                for j in countyList:
                    if j['county_name'] == userCounty:
                        cases[v] = j['total_cases']
    return cases 

#if user enters a valid county, create new sheet and fill it with data from 14 days.
if userCounty.lower() in totalCounties:
    timeCases = timelineSelector()
    countySheet = wb.create_sheet(userCounty + 'Sheet')
    i = 1
    for k, v in sorted(timeCases.items()):
        countySheet.cell(row= i, column= 1).value = k
        countySheet.cell(row= i, column= 2).value = v
        i = i + 1
    
    countyChart = openpyxl.chart.LineChart()
    countyChart.style = 12
    countyChart.title = 'Evolution of cases'
    countyChart.x_axis.title = 'Last 14 days'
    countyChart.y_axis.title = 'Confirmed cases'

    countyData = openpyxl.chart.Reference(countySheet, min_col= 2, min_row= 1, max_row = countySheet.max_row, max_col= 2)
    countyDates = openpyxl.chart.Reference(countySheet, min_col= 1, min_row= 1, max_row= countySheet.max_row)
    countyChart.add_data(countyData, titles_from_data= False)
    countyChart.set_categories(countyDates)
    countyChart.legend = None
    countySheet.add_chart(countyChart, 'E5')
    
data = daySelector()
try:
    for county in data:
        counties.append(county['county_name'])
        cases.append(county['total_cases'])
except:
    print('Data not found.')
    sys.exit(1)
#writing data to excel file
for index, county in enumerate(counties, 1):
    sheet.cell(row= index, column= 1).value = county

for index, case in enumerate(cases, 1):
    sheet.cell(row= index, column= 2).value = case

#Creating the pie chart and saving the excel file
pieChart = openpyxl.chart.PieChart()
labels = openpyxl.chart.Reference(sheet, min_col= 1, min_row= 1, max_row= sheet.max_row)
data = openpyxl.chart.Reference(sheet, min_col= 2, min_row= 1, max_row= sheet.max_row)
pieChart.add_data(data, titles_from_data= False)
pieChart.set_categories(labels)
pieChart.title = 'Covid cases for ' + userDate
pieChart.height = 12
pieChart.width = 15
sheet.add_chart(pieChart, 'E5')
wb.save(os.path.join('covidData', str(userDate + '.xlsx'))) # the file will be named after the date.



