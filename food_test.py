import myfitnesspal
import datetime
from openpyxl import load_workbook


client = myfitnesspal.Client('USERNAME', 'PASSWORD')
theDate = datetime.date.today()
wb = load_workbook('TheNameOfYourFile.xlsx')
ws2 = wb['TheNameOfTheSheet']

for day in range (1, theDate.day + 1):
    today = client.get_date(theDate.year, theDate.month, day)
    curr = datetime.date(theDate.year, theDate.month, day)
    weight = client.get_measurements('Weight', curr, curr)
    print("Printing " + str(day))
    if list(weight.values()):
        try:
            macros = [curr, today.totals['calories'], today.totals['carbohydrates'], today.totals['fat'],
                      today.totals['protein'], today.totals['sodium'], today.totals['sugar'], list(weight.values())[0]]
            ws2.append(macros)
        except:
            print("Empty")
    else:
        try:
            macros = [curr, today.totals['calories'], today.totals['carbohydrates'], today.totals['fat'],
                      today.totals['protein'], today.totals['sodium'], today.totals['sugar']]
            ws2.append(macros)
        except:
            print("Empty")


wb.save('TheNameOfYourFile.xlsx')
