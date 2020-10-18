#!python3
#Covid cases pie chart by county, Romania
import json, requests, openpyxl, os, sys
import pyinputplus as pyip

#Create a folder where we'll put the file with data
abspath = os.path.abspath(sys.argv[0])
dname = os.path.dirname(abspath)
os.chdir(dname)
os.makedirs('covidData', exist_ok=True)


#Get the JSON data
url = 'https://www.graphs.ro/json.php'
res = requests.get(url)
try:
    res.raise_for_status()
except:
    print('Could not connect to server.')
w = json.loads(res.text)

#Select the date you're interested in

userDate = str(pyip.inputDate(prompt= 'Enter a date (yyyy-mm-dd) format: ',formats= ['%Y-%m-%d']))

def daySelector():
    for i in w['covid_romania']:
        for v in i.values():
            if v == userDate:
                dayData = i
                cases = i['new_cases_today']
                tests = i['new_tests_today']
                print(str(cases) + ' cases confirmed from '+ str(tests) + ' tests, that means ' + str(round(cases / (tests / 100),1)) + '%') 
                return dayData
    

#Create the excel file
wb = openpyxl.Workbook()
sheet = wb.active
counties = []
cases = []

data = daySelector()
try:
    for county in data['county_data']:
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


