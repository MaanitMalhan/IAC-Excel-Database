import openpyxl 
from openpyxl.styles import Font

def arc_code_sheet(workbook):
    # Load the destination workbook
    destination_workbook = workbook
    arc_sheet = destination_workbook.create_sheet(title="ARC_CODES")
    
    
    arc_sheet['A1'] = "ARC_CODE"
    arc_sheet['B1'] = "ARC_NAME"
    arc_sheet['A1'].font = Font(bold=True)
    arc_sheet['B1'].font = Font(bold=True)
    #2
    arc_sheet['A2'] = "2"
    arc_sheet['B2'] = "Energy Management"
    arc_sheet['A2'].font = Font(bold=True)
    arc_sheet['B2'].font = Font(bold=True)
    #2.1
    arc_sheet['A3'] = "2.1"
    arc_sheet['B3'] = "Combustion Systems"
    arc_sheet['A4'] = "2.11"
    arc_sheet['B4'] = "Furnaces, Ovens & Direct Fired Operations"
    arc_sheet['A5'] = "2.12"
    arc_sheet['B5'] = "Boilers"
    arc_sheet['A6'] = "2.13"
    arc_sheet['B6'] = "Fuel Switching"
    #2.2
    arc_sheet['A7'] = "2.2"
    arc_sheet['B7'] = "Thermal Systems"
    arc_sheet['A8'] = "2.21"
    arc_sheet['B8'] = "Steam"
    arc_sheet['A9'] = "2.22"
    arc_sheet['B9'] = "Heating"
    arc_sheet['A10'] = "2.23"
    arc_sheet['B10'] = "Heat Treating"
    arc_sheet['A11'] = "2.24"
    arc_sheet['B11'] = "Heat Recovery"
    arc_sheet['A12'] = "2.25"
    arc_sheet['B12'] = "Heat Containment"
    arc_sheet['A13'] = "2.26"
    arc_sheet['B13'] = "Cooling"
    arc_sheet['A14'] = "2.27"
    arc_sheet['B14'] = "Drying"
    #2.3
    arc_sheet['A15'] = "2.3"
    arc_sheet['B15'] = "Electrical Power"
    arc_sheet['A16'] = "2.31"
    arc_sheet['B16'] = "Demand Management"
    arc_sheet['A17'] = "2.32"
    arc_sheet['B17'] = "Power Factor"
    arc_sheet['A18'] = "2.33"
    arc_sheet['B18'] = "Generation of Power"
    arc_sheet['A19'] = "2.34"
    arc_sheet['B19'] = "Cogeneration"
    arc_sheet['A20'] = "2.35"
    arc_sheet['B20'] = "Transmission"
    #2.4
    arc_sheet['A21'] = "2.4"
    arc_sheet['B21'] = "Motor Systems"
    arc_sheet['A22'] = "2.41"
    arc_sheet['B22'] = "Motors"
    arc_sheet['A23'] = "2.42"
    arc_sheet['B23'] = "Air Compressors"
    arc_sheet['A24'] = "2.43"
    arc_sheet['B24'] = "Other equipment"
    #2.5
    arc_sheet['A25'] = "2.5"
    arc_sheet['B25'] = "Industrial Design"
    arc_sheet['A26'] = "2.51"
    arc_sheet['B26'] = "Systems"
    #2.6
    arc_sheet['A27'] = "2.6"
    arc_sheet['B27'] = "Operations"
    arc_sheet['A28'] = "2.61"
    arc_sheet['B28'] = "Maintenance"
    arc_sheet['A29'] = "2.62"
    arc_sheet['B29'] = "Equipment Control"
    #2.7
    arc_sheet['A30'] = "2.7"
    arc_sheet['B30'] = "Building and Grounds"
    arc_sheet['A31'] = "2.71"
    arc_sheet['B31'] = "Lighting"
    arc_sheet['A32'] = "2.72"
    arc_sheet['B32'] = "Space Conditioning"
    arc_sheet['A33'] = "2.73"
    arc_sheet['B33'] = "Ventilation"
    arc_sheet['A34'] = "2.74"
    arc_sheet['B34'] = "Building Envelope"
    #2.8
    arc_sheet['A35'] = "2.8"
    arc_sheet['B35'] = "Ancillary Costs"
    arc_sheet['A36'] = "2.81"
    arc_sheet['B36'] = "Administrative"
    arc_sheet['A37'] = "2.82"
    arc_sheet['B37'] = "Shipping, Distribution, and Transportation"
    #2.9
    arc_sheet['A38'] = "2.9"
    arc_sheet['B38'] = "Alternative Energy Usage"
    arc_sheet['A39'] = "2.91"
    arc_sheet['B39'] = "General"
    #3
    arc_sheet['A40'] = "3"
    arc_sheet['B40'] = "Waste Minimization/POllution Prevention"
    arc_sheet['A40'].font = Font(bold=True)
    arc_sheet['B40'].font = Font(bold=True)
    #3.1
    arc_sheet['A41'] = "3.1"
    arc_sheet['B41'] = "Operations"
    arc_sheet['A42'] = "3.11"
    arc_sheet['B42'] = "Procedures"
    arc_sheet['A43'] = "3.12"
    arc_sheet['B43'] = "Waste Stream Contamination"
    #3.2
    arc_sheet['A44'] = "3.2"
    arc_sheet['B44'] = "Equipment"
    arc_sheet['A45'] = "3.21"
    arc_sheet['B45'] = "General"
    #3.3
    arc_sheet['A46'] = "3.3"
    arc_sheet['B46'] = "Post Generation Treatment/Minimization"
    arc_sheet['A47'] = "3.31"
    arc_sheet['B47'] = "General"
    #3.4
    arc_sheet['A48'] = "3.4"
    arc_sheet['B48'] = "Water Use"
    arc_sheet['A49'] = "3.41"
    arc_sheet['B49'] = "General"
    #3.5
    arc_sheet['A50'] = "3.5"
    arc_sheet['B50'] = "Recycling"
    arc_sheet['A51'] = "3.51"
    arc_sheet['B51'] = "Liquid Waste"
    arc_sheet['A52'] = "3.52"
    arc_sheet['B52'] = "Solid Waste"
    arc_sheet['A53'] = "3.53"
    arc_sheet['B53'] = "Other Materials"
    #3.6
    arc_sheet['A54'] = "3.6"
    arc_sheet['B54'] = "Waste Disposal"
    arc_sheet['A55'] = "3.61"
    arc_sheet['B55'] = "General"
    #3.7
    arc_sheet['A56'] = "3.7"
    arc_sheet['B56'] = "Matinenance"
    arc_sheet['A57'] = "3.71"
    arc_sheet['B57'] = "Cleaning/Degreasing"
    arc_sheet['A58'] = "3.72"
    arc_sheet['B58'] = "Spillage"
    arc_sheet['A59'] = "3.73"
    arc_sheet['B59'] = "Other"
    #3.8
    arc_sheet['A60'] = "3.8"
    arc_sheet['B60'] = "Raw Materials"
    arc_sheet['A61'] = "3.81"
    arc_sheet['B61'] = "Solvents"
    arc_sheet['A62'] = "3.82"
    arc_sheet['B62'] = "Other Solutions"
    arc_sheet['A63'] = "3.83"
    arc_sheet['B63'] = "Solids"
    #4
    arc_sheet['A64'] = "4"
    arc_sheet['B64'] = "Direct Productivity Enhancements"
    arc_sheet['A64'].font = Font(bold=True)
    arc_sheet['B64'].font = Font(bold=True)
    #4.1
    arc_sheet['A65'] = "4.1"
    arc_sheet['B65'] = "Manufacturing Enhancements"
    arc_sheet['A66'] = "4.11"
    arc_sheet['B66'] = "Botteneck reduction"
    arc_sheet['A67'] = "4.12"
    arc_sheet['B67'] = "Defect Reduction"
    arc_sheet['A68'] = "4.13"
    arc_sheet['B68'] = "Material Reduction"
    #4.2
    arc_sheet['A69'] = "4.2"
    arc_sheet['B69'] = "Purchasing"
    arc_sheet['A70'] = "4.21"
    arc_sheet['B70'] = "Raw Materials"
    arc_sheet['A71'] = "4.22"
    arc_sheet['B71'] = "Ancillary Materials"
    arc_sheet['A72'] = "4.23"
    arc_sheet['B72'] = "Capital"
    #4.3
    arc_sheet['A73'] = "4.3"
    arc_sheet['B73'] = "Inventory"
    arc_sheet['A74'] = "4.31"
    arc_sheet['B74'] = "Just in Time"
    arc_sheet['A75'] = "4.32"
    arc_sheet['B75'] = "Other inventory controls"
    #4.4
    arc_sheet['A74'] = "4.4"
    arc_sheet['B74'] = "Labor Optimization"
    arc_sheet['A75'] = "4.42"
    arc_sheet['B75'] = "Practice/Procedure"
    arc_sheet['A76'] = "4.43"
    arc_sheet['B76'] = "Training"
    arc_sheet['A77'] = "4.44"
    arc_sheet['B77'] = "Automation"
    arc_sheet['A78'] = "4.45"
    arc_sheet['B78'] = "Scheduling"
    arc_sheet['A79'] = "4.46"
    arc_sheet['B79'] = "Maintenance"  
    #4.5
    arc_sheet['A80'] = "4.5"
    arc_sheet['B80'] = "Space Utilization"
    arc_sheet['A81'] = "4.51"
    arc_sheet['B81'] = "Floor Layout"
    arc_sheet['A82'] = "4.52"
    arc_sheet['B82'] = "Rental Space"
    #4.6
    arc_sheet['A83'] = "4.6"
    arc_sheet['B83'] = "Reduction of Downtime"
    arc_sheet['A84'] = "4.61"
    arc_sheet['B84'] = "Maintenance"
    arc_sheet['A85'] = "4.62"
    arc_sheet['B85'] = "Quick Change"
    arc_sheet['A86'] = "4.63"
    arc_sheet['B86'] = "Power Conditioning"
    arc_sheet['A87'] = "4.64"
    arc_sheet['B87'] = "Alarms"
    arc_sheet['A88'] = "4.65"
    arc_sheet['B88'] = "Other Equipment"
    arc_sheet['A89'] = "4.66"
    arc_sheet['B89'] = "Industrial Internet of Things(IIOT)"
    #4.7
    arc_sheet['A90'] = "4.7"
    arc_sheet['B90'] = "Management Practices"
    arc_sheet['A91'] = "4.71"
    arc_sheet['B91'] = "Total Quality Management"
    arc_sheet['A92'] = "4.72"
    arc_sheet['B92'] = "Certifications"
    arc_sheet['A93'] = "4.73"
    arc_sheet['B93'] = "Marketing"
    #4.8
    arc_sheet['A94'] = "4.8"
    arc_sheet['B94'] = "Other Administrative Savings"
    arc_sheet['A95'] = "4.81"
    arc_sheet['B95'] = "Taxes"
    arc_sheet['A96'] = "4.82"
    arc_sheet['B96'] = "Fees"

    #arc_destination_sheet = destination_workbook['ARC_CODES']  
    #destination_workbook.save('/Users/maanitmalhan/Documents/IAC_Center/excel-data-iac/files/test.xlsx')
    #print("ARC_CODES sheet created successfully!")
    
