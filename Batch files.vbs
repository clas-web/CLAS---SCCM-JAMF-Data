https://docs.google.com/document/d/1L2IAqDDRlHqEAHEQSDapt9iG5RNfLrkedQoxTykGZjQ/edit#

Set objExcel = CreateObject("Excel.Application")

Set objWB = objExcel.Workbooks.Open("G:\My Drive\IT Support\SCCM ATOM Feeds\Excel Feeds\All_Machines.xlsx")
objExcel.Visible = True
objWB.RefreshAll

objExcel.DisplayAlerts = False
objWB.Close True 'False to not save
objExcel.DisplayAlerts = True 

WScript.Sleep 40000
objExcel.Quit 

'***********************************************************************************************************************************

Set objExcel = CreateObject("Excel.Application")

Set objWB = objExcel.Workbooks.Open("G:\My Drive\IT Support\SCCM ATOM Feeds\Excel Feeds\casper_xml_import-from-web.xlsx")
objExcel.Visible = True
objWB.RefreshAll

objExcel.DisplayAlerts = False
objWB.Close True 'False to not save
objExcel.DisplayAlerts = True 

WScript.Sleep 20000
objExcel.Quit 
