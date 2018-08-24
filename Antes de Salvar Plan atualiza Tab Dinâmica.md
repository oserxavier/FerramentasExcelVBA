# FerramentasExcelVBA
Nessa pasta disponho minhas ferramentas que utilizo habitualmente para tratamento de diversas planilhas. 

'Site :http://inanyplace.blogspot.com.br/2012/09/excel-vba-refresh-pivot-tables.html

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
 Dim nPT As PivotTable
 Dim nWS As Worksheet
 For Each nWS In ActiveWorkbook.Worksheets
 For Each nPT In nWS.PivotTables
 nPT.RefreshTable
 Next nPT
 Next nWS

End Sub
