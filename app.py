from openpyxl import load_workbook
import random 
wb = load_workbook(filename = 'MQD.xlsx')
for sheet in wb:
    rr = random.choice(list(sheet))
    print(rr[0].value + ' :: ' + rr[1].value)



        