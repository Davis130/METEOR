from attr import field
import pandas as pd
import win32com.client as win32
from tkinter import filedialog as fd
from pathlib import Path 


def clear_tcd(ws):
    for tcd in ws.PivotTables():
        tcd.TableRange2.Clear()

def insert_tcd_field_set_bruts(tcd):
    """ blablabla"""
    field_rows = {}
    field_rows['Stock'] = tcd.PivotFields('Stock')
    field_rows['Etat'] = tcd.PivotFields('Etat')

    field_values = {}
    field_values['Code'] = tcd.PivotFields('Code')
    field_values['Palier inventaire'] = tcd.PivotFields('Palier inventaire')
    field_values['Alerte Palier inventaire'] = tcd.PivotFields('Alerte Palier inventaire')
    
    #insert row fields to pivot table design
    field_rows['Stock'].Orientation = 1
    field_rows['Stock'].Position  = 1

    field_rows['Etat'].Orientation = 1
    field_rows['Etat'].Position  = 2

    #insert data field
    field_values['Palier inventaire'].Orientation = 2
    field_values['Alerte Palier inventaire'].Orientation = 2
    field_values['Code'].Orientation = 4
   
    #field_rows['Code'].Position  = -4112

def insert_tcd_field_set_bruts_4B_INV_AZU(tcd):

    field_rows = {}
    field_rows['Site'] = tcd.PivotFields('Site')

    field_values = {}
    field_values['Code'] = tcd.PivotFields('Code')
    field_values['Palier inventaire'] = tcd.PivotFields('Palier inventaire')

    #insert row fields to pivot table design
    field_rows['Site'].Orientation = 1
    field_rows['Site'].Position  = 1

    #insert data field
    field_values['Palier inventaire'].Orientation = 2
    field_values['Code'].Orientation = 4






def insert_tcd_field_set_prioritaires(tcd):
    """ blablabla"""
    field_rows = {}
    field_rows['Site'] = tcd.PivotFields('Site')
    field_rows['Etat'] = tcd.PivotFields('Etat')

    field_values = {}
    field_values['Code'] = tcd.PivotFields('Code')
    field_values['Palier'] = tcd.PivotFields('Palier')
    field_values['Alerte_Palier'] = tcd.PivotFields('Alerte_Palier')
    
    #insert row fields to pivot table design
    field_rows['Site'].Orientation = 1
    field_rows['Site'].Position  = 1

    field_rows['Etat'].Orientation = 1
    field_rows['Etat'].Position  = 2

    #insert data field
    field_values['Palier'].Orientation = 2
    field_values['Alerte_Palier'].Orientation = 2
    field_values['Code'].Orientation = 4    

def choix_fichier():
    return fd.askopenfilename(initialdir = __file__)
  
def addpivot(wb,sourcedata,title,filters=(),columns=(),
             rows=(),sumvalue=(),sortfield=""):
    """Build a pivot table using the provided source location data
    and specified fields
    """
    ...
    for fieldlist,fieldc in ((filters ,win32.xlPageField),
                            (columns  ,win32.xlColumnField),
                            (rows     ,win32.xlRowField)):
        for i,val in enumerate(fieldlist):
            wb.ActiveSheet.PivotTables(tname).PivotFields(val).Orientation = fieldc
        wb.ActiveSheet.PivotTables(tname).PivotFields(val).Position = i+1

def autres(file,source,tcd,tcd_2=''):

    #Launch Application
    xlApp = win32.Dispatch('Excel.Application')
    xlApp.visible = True

    #Reference Workbook
    wb = xlApp.Workbooks.Open(file)
    wb.Sheets.Add(xlApp.ActiveSheet).name = tcd
    wb.Sheets.Add(xlApp.ActiveSheet).name = tcd_2

    #Reference Worksheets
    ws_data = wb.Worksheets(source)
    ws_report = wb.Worksheets(tcd)
    ws_report2 = wb.Worksheets(tcd_2)


    #Clear PivotTables
    clear_tcd(ws_report)
    clear_tcd(ws_report2)

    #create tcd cache connection
    tcd_cache = wb.PivotCaches().Create(1,ws_data.Range("A1").CurrentRegion)
    tcd_cache2 = wb.PivotCaches().Create(1,ws_data.Range("A1").CurrentRegion)

    # create TCD
    #tcd = tcd_cache.createPivotTable(ws_report.Range("B4"), "En cours de Blanchiment")
    tcd = tcd_cache.createPivotTable(ws_report.Range("B4"), "Tableau")
    tcd_2 = tcd_cache2.createPivotTable(ws_report2.Range("B4"), "Tableau")

    #Toggle grand totals
    tcd.ColumnGrand =  True
    tcd.RowGrand = True
    tcd_2.ColumnGrand =  True
    tcd_2.RowGrand = True

    # Change subtotal Location
    tcd.SubtotalLocation(1) # bottom (1 = top)
    tcd_2.SubtotalLocation(1) # bottom (1 = top)

    #change report Layout
    #tcd.RowAxisLayout(1)

    #change pivot table style
    #tcd.TableStyle2 = 'PivotStyleMedium9'

    REP = Path(__file__).parent
    fichier = Path(file).name

    sauve = REP / fichier

   

    #create report 
    if 'Bruts' in file:
        insert_tcd_field_set_bruts(tcd)
        insert_tcd_field_set_bruts_4B_INV_AZU(tcd_2)
    elif 'Prioritaires'in file:
        insert_tcd_field_set_prioritaires(tcd)
        #insert_tcd_field_set_bruts_4B_INV_AZU(tcd_2)

    wb.SaveAs(str(sauve))
    xlApp.Application.Quit() 





if __name__ == '__main__':


    fichier = choix_fichier() 
    if fichier and Path(fichier).exists():
        print(fichier)
        if 'Bruts' in fichier:
            autres(fichier,'Etats Météor Bruts','4 Inv AZU','4B Inv AZU')
            

        elif 'Prioritaires'in fichier:
            autres(fichier,'Etats Météor Prioritaires','Etats Météor Prioritaires_tcd','4B Inv AZU')
            
    else:
        print('pas de choix')
    #print("test")