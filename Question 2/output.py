### QUESTION-2 ###

from datafile import *
from formulation import *

### OUTPUT ###

if model.status == grb.GRB.OPTIMAL:
    obj = round(model.objVal, 2)
    print('\nTotal cost -', obj, '\n')

    FinalQuantity = [[round(QFinal[v,i].x,2) for i in range(num_items)] for v in range(num_vendors)]
    UnfulfilledDemand = [round(DUnfulfilled[i].x,2) for i in range(num_items)]      
    
    print(pd.DataFrame(FinalQuantity, index = V, columns = I),'\n')
    print(pd.DataFrame(UnfulfilledDemand, index = I, columns = ['Unfulfilled Demand']),'\n')
    print(pd.DataFrame([DiscountSD[v].x for v in range(num_vendors)], index = V, columns = ['Spend Discount']),'\n')
    print(pd.DataFrame([DiscountQD[v].x for v in range(num_vendors)], index = V, columns = ['Shipment Discount']),'\n')

    print('Quantity left with vendor -\n')
    print(pd.DataFrame([[round(Quantity[v][i] - FinalQuantity[v][i],2) for i in range(num_items)] for v in range(num_vendors)], index = V, columns = I),'\n')

    print('\nS1 -',S1)
    print('\nS2 -',S2)
    print('\n',pd.DataFrame([sum(round(Price[v][i]*FinalQuantity[v][i],2) for i in range(num_items)) for v in range(num_vendors)], index = V, columns = ['Cost to vendor']))

    print('\nQ1 -',Q1)
    print('\nQ2 -',Q2)
    print('\n',pd.DataFrame([sum(round(FinalQuantity[v][i],2) for i in range(num_items)) for v in range(num_vendors)], index = V, columns = ['Quantity from vendor']))

else:
    print('The problem does not have an optimal solution.')

# END TIME
now = datetime.now()
print("\nEnded at -", now.strftime("%H:%M:%S"))
end = time.time()
print(f"Runtime of the program is {round(((end - start)/60), 2)} minutes\n")

# Saving data to a new worksheet
workbook = openpyxl.Workbook()
sheet = workbook.active

for i in range(num_items):
    for v in range(num_vendors):
        sheet.cell(row = 1, column = 1, value = 'DEMAND FROM VENDORS')
        sheet.cell(row = 1, column = i+2, value = I[i])
        sheet.cell(row = v+2, column = 1, value = V[v])
        sheet.cell(row = v+2, column = i+2, value = FinalQuantity[v][i])

    sheet.cell(row = num_vendors+3, column = 1, value = 'VENDOR')
    sheet.cell(row = num_vendors+3, column = 2, value = 'DEMAND UNFULFILLED')
    sheet.cell(row = num_vendors+4+i, column = 1, value = I[i])
    sheet.cell(row = num_vendors+4+i, column = 2, value = UnfulfilledDemand[i])

workbook.save('output.xlsx')
os.startfile('output.xlsx') # opens the output excel file