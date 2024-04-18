import plotly.express
import plotly.io as pio
import plotly.graph_objects as go
import openpyxl
from app import *


def plot_creation(workbook):
    #do a sunburst graph, parent: Tot_ASSM, Tot_RECC Child: median, mean, max, min Total saving type Primary, Secondary, Tertiary, Quaternary
    calc_sheet = workbook['calculation']
    recc_sheet = workbook['RECC']
    count = 0
    column_to_check = "G"
    populated_rows = 0
    # Iterate over rows in the column and count populated ones
    for row in recc_sheet.iter_rows(min_row=1, max_row=recc_sheet.max_row, min_col=1, max_col=1):
        cell_value = row[0].value
        if cell_value is not None and str(cell_value).strip() != '':
            populated_rows += 1

    data = dict(
        character=["Total Assessment", "Total Recommendations","Avg # of Recommendations per assessment","Total Savings", "Avg Savings from Recommendations","Primary Savings", "Secondary Savings", "Tertiary Savings", "Quaternary Savings", "Total Implementation Cost"],
        parent=["", "", "Total Recommendations", "Total Assessment", "Total Recommendations", "Total Recommendations", "Total Recommendations", "Total Recommendations", "Total Recommendations", "Total Assessment"],
        value=[calc_sheet['A2'].value,calc_sheet['B2'].value,calc_sheet['C2'].value,recc_sheet[f'O{populated_rows+4}'].value,recc_sheet[f'W{populated_rows+4}'].value,recc_sheet[f'K{populated_rows+2}'].value,recc_sheet[f'O{populated_rows+2}'].value,recc_sheet[f'S{populated_rows+2}'].value,recc_sheet[f'W{populated_rows+2}'].value,recc_sheet[f'G{populated_rows+2}'].value]
    )
    
    fig = plotly.express.sunburst(
        data,
        names='character',
        parents='parent',
        values='value'
    )
    #fig.update_traces(textinfo="label+percent entry")




    pio.write_html(fig, file=f'{universal_dir}SNE_IAC_Database_plots.html', auto_open=True)



