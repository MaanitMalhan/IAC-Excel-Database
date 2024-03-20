import plotly
import openpyxl


def plot_creation():
    pass#do a sunburst graph


def plotly_graphs_in_excel(workbook):
    destination_workbook = workbook
    destination_sheet = destination_workbook.create_sheet(title="Graphs")
    graph_sheet = destination_workbook['Graphs']
    
    