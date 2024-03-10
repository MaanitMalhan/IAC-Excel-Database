import openpyxl
from assessSheetExtraction import count_assem
from reccSheetExtraction import count_recc
#Add recommended savings, plant energy costs and number of reccs

def calculations(workbook):
    destination_workbook = workbook
    calc_sheet = destination_workbook.create_sheet(title="calculation")
    destination_sheet = destination_workbook['calculation']

    #Add recommended savings, plant energy costs and number of reccs
    destination_sheet['A1'] = 'Total_number_of_assessments'
    destination_sheet['B1'] = 'Total_number_of_recommendations'
    destination_sheet['C1'] = 'Total_recommended_savings'
    destination_sheet['D1'] = 'Total_plant_energy_costs'
    destination_sheet['E1'] = 'Total_number_of_recommendations_per_assessment'
    destination_sheet['F1'] = 'Total_recommended_savings_per_assessment'

    #Add recommended savings, plant energy costs and number of reccs
    destination_sheet['A2'] = count_assem(destination_workbook['ASSESS'])
    destination_sheet['B2'] = count_recc(destination_workbook['RECC'])

